"""Procesa los requerimientos de operaciones, genera un reporte Excel y lo envía por correo.

Flujo: carga -> validación de umbral de error -> cuarentena -> imputación
       -> exportación a Excel -> envío por correo.
"""

import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

ARCHIVO_ENTRADA = "datos_operaciones.csv"
ARCHIVO_CUARENTENA = "cuarentena.csv"
ARCHIVO_EXCEL = "reporte_operaciones.xlsx"
UMBRAL_MAXIMO_FILAS_CON_NULOS = 0.40  # 40%
ESTADO_POR_DEFECTO = "Sin clasificar"
# Formato CSV compatible con Excel en español
SEPARADOR_CSV = ";"
DECIMAL_CSV = ","
CODIFICACION_CSV = "utf-8-sig"


class ArchivoCorruptoError(Exception):
    """Error crítico: el archivo de origen supera el umbral de filas con nulos."""


def cargar_datos(nombre_archivo: str) -> pd.DataFrame:
    """Lee el CSV y detiene la ejecución si más del 40% de las filas tiene algún nulo."""
    tabla_requerimientos = pd.read_csv(nombre_archivo, sep=SEPARADOR_CSV, decimal=DECIMAL_CSV, encoding=CODIFICACION_CSV)

    porcentaje_filas_con_nulos = tabla_requerimientos.isna().any(axis=1).mean()
    if porcentaje_filas_con_nulos > UMBRAL_MAXIMO_FILAS_CON_NULOS:
        raise ArchivoCorruptoError(
            "Alerta: El archivo de origen está altamente corrupto "
            f"({porcentaje_filas_con_nulos:.1%} de las filas tiene valores nulos)."
        )
    return tabla_requerimientos


def separar_cuarentena(tabla_requerimientos: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Separa las filas sin ID_Requerimiento del DataFrame principal."""
    filtro_sin_id = tabla_requerimientos["ID_Requerimiento"].isna()
    tabla_cuarentena = tabla_requerimientos[filtro_sin_id]
    tabla_valida = tabla_requerimientos[~filtro_sin_id].copy()
    return tabla_valida, tabla_cuarentena


COLUMNAS_NUMERICAS_A_IMPUTAR = ["Horas_Invertidas", "Costo_Materiales", "Costo_Mano_Obra"]


def imputar_datos(tabla_valida: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Rellena Estado faltante y horas/costos faltantes (media por Tipo_Servicio).

    Devuelve la tabla imputada y la cantidad de filas que tuvieron al menos una imputación.
    """
    columnas_a_revisar = ["Estado"] + COLUMNAS_NUMERICAS_A_IMPUTAR
    cantidad_filas_imputadas = int(tabla_valida[columnas_a_revisar].isna().any(axis=1).sum())

    tabla_valida["Estado"] = tabla_valida["Estado"].fillna(ESTADO_POR_DEFECTO)

    for nombre_columna in COLUMNAS_NUMERICAS_A_IMPUTAR:
        promedio_por_servicio = tabla_valida.groupby("Tipo_Servicio")[nombre_columna].transform("mean")
        tabla_valida[nombre_columna] = tabla_valida[nombre_columna].fillna(promedio_por_servicio.round(2))
    return tabla_valida, cantidad_filas_imputadas


def procesar_datos(tabla_valida: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Calcula requerimientos por estado y horas invertidas por tipo de servicio."""
    total_requerimientos_por_estado = (
        tabla_valida.groupby("Estado")["ID_Requerimiento"].count().rename("Total_Requerimientos").reset_index()
    )
    total_horas_por_tipo_servicio = (
        tabla_valida.groupby("Tipo_Servicio")["Horas_Invertidas"].sum().rename("Total_Horas").reset_index()
    )
    return total_requerimientos_por_estado, total_horas_por_tipo_servicio


def exportar_excel(
    total_requerimientos_por_estado: pd.DataFrame,
    total_horas_por_tipo_servicio: pd.DataFrame,
    tabla_datos_limpios: pd.DataFrame,
    tabla_cuarentena: pd.DataFrame,
    filas_imputadas: int,
    nombre_archivo: str,
) -> None:
    """Guarda las métricas, los datos limpios y la cuarentena en tres hojas de un Excel."""
    resumen_ejecucion = pd.DataFrame({
        "Indicador": ["Filas procesadas con éxito", "Filas imputadas", "Filas en cuarentena"],
        "Cantidad": [len(tabla_datos_limpios), filas_imputadas, len(tabla_cuarentena)],
    })

    with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as escritor_excel:
        # Hoja "Métricas": tres tablas apiladas, separadas por una fila en blanco
        fila_inicio = 0
        for tabla_metrica in (resumen_ejecucion, total_requerimientos_por_estado, total_horas_por_tipo_servicio):
            tabla_metrica.to_excel(escritor_excel, sheet_name="Métricas", index=False, startrow=fila_inicio)
            fila_inicio += len(tabla_metrica) + 2

        tabla_datos_limpios.to_excel(escritor_excel, sheet_name="Datos_Limpios", index=False)
        tabla_cuarentena.to_excel(escritor_excel, sheet_name="Cuarentena", index=False)

        # Ajusta el ancho de cada columna al contenido
        for hoja in escritor_excel.sheets.values():
            for columna in hoja.columns:
                ancho_maximo = max(len(str(celda.value)) for celda in columna if celda.value is not None)
                hoja.column_dimensions[columna[0].column_letter].width = ancho_maximo + 3


def enviar_reporte(ruta_archivo: str) -> None:
    """Envía por correo el reporte adjunto usando credenciales del archivo .env.

    Variables requeridas: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, DESTINATARIO.
    """
    load_dotenv()
    servidor_smtp = os.getenv("SMTP_HOST")
    puerto_smtp = os.getenv("SMTP_PORT")
    usuario_smtp = os.getenv("SMTP_USER")
    clave_smtp = os.getenv("SMTP_PASS")
    destinatario = os.getenv("DESTINATARIO")

    variables_faltantes = [
        nombre for nombre, valor in [
            ("SMTP_HOST", servidor_smtp), ("SMTP_PORT", puerto_smtp), ("SMTP_USER", usuario_smtp),
            ("SMTP_PASS", clave_smtp), ("DESTINATARIO", destinatario),
        ] if not valor
    ]
    if variables_faltantes:
        raise ValueError(f"Faltan variables en el archivo .env: {', '.join(variables_faltantes)}")

    mensaje = EmailMessage()
    mensaje["Subject"] = "Reporte de operaciones"
    mensaje["From"] = usuario_smtp
    mensaje["To"] = destinatario
    mensaje.set_content("Adjunto el reporte consolidado de operaciones.")

    ruta = Path(ruta_archivo)
    mensaje.add_attachment(
        ruta.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=ruta.name,
    )

    with smtplib.SMTP(servidor_smtp, int(puerto_smtp), timeout=30) as conexion_smtp:
        conexion_smtp.starttls()
        conexion_smtp.login(usuario_smtp, clave_smtp)
        conexion_smtp.send_message(mensaje)


def main() -> None:
    """Orquesta todo el flujo y maneja los errores esperables."""
    try:
        tabla_requerimientos = cargar_datos(ARCHIVO_ENTRADA)
        tabla_valida, tabla_cuarentena = separar_cuarentena(tabla_requerimientos)
        tabla_cuarentena.to_csv(
            ARCHIVO_CUARENTENA, index=False, sep=SEPARADOR_CSV, decimal=DECIMAL_CSV, encoding=CODIFICACION_CSV
        )

        tabla_imputada, cantidad_filas_imputadas = imputar_datos(tabla_valida)
        requerimientos_por_estado, horas_por_tipo = procesar_datos(tabla_imputada)
        exportar_excel(
            requerimientos_por_estado,
            horas_por_tipo,
            tabla_imputada,
            tabla_cuarentena,
            cantidad_filas_imputadas,
            ARCHIVO_EXCEL,
        )
        print(f"Reporte '{ARCHIVO_EXCEL}' y '{ARCHIVO_CUARENTENA}' generados con éxito.")
    except ArchivoCorruptoError as error:
        print(error)
        raise SystemExit(1)
    except FileNotFoundError:
        print(f"Error: no se encontró '{ARCHIVO_ENTRADA}'. Ejecuta primero crear_datos.py.")
        return
    except pd.errors.EmptyDataError:
        print(f"Error: '{ARCHIVO_ENTRADA}' está vacío.")
        return
    except KeyError as error:
        print(f"Error: falta la columna esperada {error} en los datos.")
        return
    except PermissionError:
        print(f"Error: no se pudo escribir un archivo de salida (¿'{ARCHIVO_EXCEL}' está abierto en Excel?).")
        return
    except Exception as error:
        print(f"Error inesperado al generar el reporte: {error}")
        return

    # El envío de correo es independiente: si falla, el reporte ya quedó generado
    try:
        enviar_reporte(ARCHIVO_EXCEL)
        print("Correo enviado con éxito.")
    except ValueError as error:
        print(f"Correo no enviado: {error}")
    except (smtplib.SMTPException, OSError) as error:
        print(f"Correo no enviado: error de conexión o autenticación SMTP ({error}).")


if __name__ == "__main__":
    main()
