# Automatizador de Reportes de Operaciones

Herramienta en Python que toma un CSV de requerimientos de operaciones técnicas, lo valida, limpia y resume, y genera un **reporte Excel consolidado** con tres pestañas. Opcionalmente lo envía por correo.

El Excel resultante es la fuente de datos del [Dashboard de Operaciones](https://github.com/Nmelon/dashboard-operaciones).

```
datos_operaciones.csv ──► generador_reportes.py ──► reporte_operaciones.xlsx ──► Dashboard web
                               │                           │
                               └─► cuarentena.csv          └─► (opcional) correo SMTP
```

## Qué hace

1. **Carga** el CSV (separador `;`, decimal `,`, UTF-8 con BOM, compatible con Excel en español).
2. **Valida el archivo completo:** si más del 40 % de las filas tiene algún valor nulo, se detiene con una alerta de archivo corrupto.
3. **Separa la cuarentena:** las filas sin `ID_Requerimiento` salen del flujo y se guardan en `cuarentena.csv`.
4. **Imputa los datos faltantes:**
   - `Estado` vacío → `Sin clasificar`.
   - `Horas_Invertidas`, `Costo_Materiales` y `Costo_Mano_Obra` vacíos → promedio del mismo `Tipo_Servicio`.
5. **Calcula métricas:** requerimientos por estado y horas por tipo de servicio.
6. **Exporta** `reporte_operaciones.xlsx` con tres pestañas:

   | Pestaña | Contenido |
   | --- | --- |
   | `Métricas` | Filas procesadas, imputadas y en cuarentena; requerimientos por estado; horas por tipo de servicio |
   | `Datos_Limpios` | Requerimientos válidos, ya imputados |
   | `Cuarentena` | Filas descartadas por no tener ID |

7. **Envía el reporte por correo** (opcional). Es un paso independiente: si falla, el Excel ya quedó generado.

## Cómo ejecutarlo

Requiere Python 3.10 o superior.

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (en Linux/macOS: source .venv/bin/activate)
pip install -r requirements.txt

python crear_datos.py           # genera datos de prueba en datos_operaciones.csv
python generador_reportes.py    # genera reporte_operaciones.xlsx y cuarentena.csv
```

`crear_datos.py` produce 30 filas sintéticas con casos de prueba fijos: 3 sin ID (cuarentena) y otras con estado, horas o costos vacíos (imputación). Usa una semilla fija, pero las fechas se calculan respecto al día en que se ejecuta, así que cambian entre ejecuciones.

## Envío por correo (opcional)

El envío usa variables de entorno leídas desde un archivo `.env`, que **nunca se sube al repositorio**.

```bash
copy .env.example .env          # en Linux/macOS: cp .env.example .env
```

Completa en `.env`:

| Variable | Descripción |
| --- | --- |
| `SMTP_HOST` | Servidor SMTP (por ejemplo `smtp.gmail.com`) |
| `SMTP_PORT` | Puerto (usa STARTTLS, normalmente `587`) |
| `SMTP_USER` | Cuenta que envía |
| `SMTP_PASS` | Contraseña o contraseña de aplicación |
| `DESTINATARIO` | Quien recibe el reporte |

Si falta alguna variable, el script lo avisa, no envía el correo y deja el Excel generado igualmente.

> **Atención:** con las cinco variables completas, cada ejecución de `generador_reportes.py` envía un correo real.

## Conexión con el dashboard

El dashboard lee `reporte_operaciones.xlsx` desde su propia carpeta `web/`. Después de generar el reporte, copia el archivo allí:

```bash
copy reporte_operaciones.xlsx ..\Dashboard\web\
```

Las columnas que el dashboard exige en `Datos_Limpios` están documentadas en el README del dashboard.

## Estructura

```
.
├── crear_datos.py          Genera el CSV de prueba
├── generador_reportes.py   Flujo principal: validar, limpiar, resumir, exportar, enviar
├── requirements.txt        pandas, openpyxl, python-dotenv
└── .env.example            Plantilla de variables para el correo
```

El código usa `pandas` para todo el manejo de datos, está dividido en funciones pequeñas (`cargar_datos`, `separar_cuarentena`, `imputar_datos`, `procesar_datos`, `exportar_excel`, `enviar_reporte`) y captura los errores esperables: archivo faltante o vacío, columna ausente, Excel abierto en otro programa y fallos SMTP.

## Licencia

[MIT](LICENSE) © 2026 Aron Saavedra
