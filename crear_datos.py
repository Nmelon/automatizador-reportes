"""Genera datos de prueba de requerimientos de operaciones técnicas."""

import numpy as np
import pandas as pd

NOMBRE_ARCHIVO_SALIDA = "datos_operaciones.csv"
CANTIDAD_FILAS = 30
TIPOS_DE_SERVICIO = ["Mantención", "Instalación", "Soporte", "Inspección"]
ESTADOS_POSIBLES = ["Pendiente", "En Proceso", "Resuelto"]
SUBESTACIONES = ["Alto Jahuel", "Polpaico", "Charrúa", "Ancoa", "Quillota"]
NIVELES_RIESGO = ["Bajo", "Medio", "Alto", "Crítico"]
DIAS_VENTANA_FECHAS = 30  # las fallas ocurren en el último mes

# Rango de horas por nivel de riesgo: los más críticos consumen más horas
RANGO_HORAS_POR_RIESGO = {"Bajo": (1, 11), "Medio": (8, 21), "Alto": (15, 31), "Crítico": (28, 49)}


def generar_datos_prueba(cantidad_filas: int) -> pd.DataFrame:
    """Crea un DataFrame corporativo con datos aleatorios y algunos valores nulos a propósito."""
    generador_aleatorio = np.random.default_rng(seed=42)

    niveles_riesgo = generador_aleatorio.choice(NIVELES_RIESGO, cantidad_filas)
    horas_invertidas = np.array([
        generador_aleatorio.integers(*RANGO_HORAS_POR_RIESGO[nivel]) for nivel in niveles_riesgo
    ], dtype=float)

    # Fechas aleatorias del último mes; la resolución ocurre entre 0 y 5 días después (sin pasar de hoy)
    hoy = pd.Timestamp.today().normalize()
    fechas_falla = hoy - pd.to_timedelta(generador_aleatorio.integers(0, DIAS_VENTANA_FECHAS, cantidad_filas), unit="D")
    dias_hasta_resolucion = pd.to_timedelta(generador_aleatorio.integers(0, 6, cantidad_filas), unit="D")
    fechas_resolucion = pd.DatetimeIndex(np.minimum((fechas_falla + dias_hasta_resolucion).to_numpy(), hoy.to_datetime64()))

    # Costos en miles de dólares; el riesgo crítico encarece materiales y mano de obra
    factor_costo = pd.Series(niveles_riesgo).map({"Bajo": 1.0, "Medio": 1.5, "Alto": 2.2, "Crítico": 3.5}).to_numpy()
    costo_materiales = (generador_aleatorio.uniform(2, 15, cantidad_filas) * factor_costo).round(2)
    costo_mano_obra = (horas_invertidas * generador_aleatorio.uniform(0.15, 0.35, cantidad_filas)).round(2)

    tabla_requerimientos = pd.DataFrame({
        "ID_Requerimiento": [f"REQ-{numero:03d}" for numero in range(1, cantidad_filas + 1)],
        "Fecha_Falla": fechas_falla.strftime("%Y-%m-%d"),
        "Fecha_Resolucion": fechas_resolucion.strftime("%Y-%m-%d"),
        "Subestacion": generador_aleatorio.choice(SUBESTACIONES, cantidad_filas),
        "Tipo_Servicio": generador_aleatorio.choice(TIPOS_DE_SERVICIO, cantidad_filas),
        "Nivel_Riesgo": niveles_riesgo,
        "Estado": generador_aleatorio.choice(ESTADOS_POSIBLES, cantidad_filas),
        "Cumple_SLA": generador_aleatorio.choice(["Sí", "No"], cantidad_filas, p=[0.75, 0.25]),
        "Horas_Invertidas": horas_invertidas,
        "Costo_Materiales": costo_materiales,
        "Costo_Mano_Obra": costo_mano_obra,
    })

    # Escenarios de prueba (posiciones fijas). Filas con algún nulo: 11 de 30 (36,7%),
    # por debajo del umbral de 40% del generador de reportes.
    posiciones_sin_id = [2, 13, 22]                # van a cuarentena
    posiciones_sin_estado = [5, 9, 17, 26]         # se imputan con "Sin clasificar"
    posiciones_sin_horas = [4, 11, 20, 26]         # se imputan con el promedio por servicio
    posiciones_sin_costo_materiales = [5, 14]      # se imputan con la media por servicio
    posiciones_sin_costo_mano_obra = [11, 17]      # se imputan con la media por servicio

    tabla_requerimientos.loc[posiciones_sin_id, "ID_Requerimiento"] = np.nan
    tabla_requerimientos.loc[posiciones_sin_estado, "Estado"] = np.nan
    tabla_requerimientos.loc[posiciones_sin_horas, "Horas_Invertidas"] = np.nan
    tabla_requerimientos.loc[posiciones_sin_costo_materiales, "Costo_Materiales"] = np.nan
    tabla_requerimientos.loc[posiciones_sin_costo_mano_obra, "Costo_Mano_Obra"] = np.nan

    return tabla_requerimientos


def guardar_datos(tabla_requerimientos: pd.DataFrame, nombre_archivo: str) -> None:
    """Guarda el DataFrame en un archivo CSV usando pandas."""
    # Formato compatible con Excel en español: separador ";", decimal "," y BOM para las tildes
    tabla_requerimientos.to_csv(nombre_archivo, index=False, sep=";", decimal=",", encoding="utf-8-sig")


def main() -> None:
    """Orquesta la generación y el guardado de los datos de prueba."""
    try:
        tabla_requerimientos = generar_datos_prueba(CANTIDAD_FILAS)
        guardar_datos(tabla_requerimientos, NOMBRE_ARCHIVO_SALIDA)
        print(f"Archivo '{NOMBRE_ARCHIVO_SALIDA}' creado con {len(tabla_requerimientos)} filas.")
    except PermissionError:
        print(f"Error: no hay permisos para escribir '{NOMBRE_ARCHIVO_SALIDA}' (¿está abierto en otro programa?).")
    except Exception as error:
        print(f"Error inesperado al crear los datos: {error}")


if __name__ == "__main__":
    main()
