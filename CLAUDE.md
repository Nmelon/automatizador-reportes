# Contexto del Proyecto
Herramienta de automatización en Python para procesar datos operativos, generar métricas de control de procesos y exportar reportes consolidados.

# Stack Tecnológico
- Python 3
- Pandas (Obligatorio para manipulación de datos)

# Reglas de Código (¡Estrictas!)
- Usa SIEMPRE la librería `pandas` para leer, limpiar, agrupar y escribir archivos de datos (CSV/Excel). Nunca uses el módulo `csv` nativo de Python.
- Escribe código modular. Separa la lógica en funciones (ej. `cargar_datos`, `procesar_datos`, `generar_reporte`).
- Implementa siempre bloques `try/except` para capturar errores comunes (ej. archivo no encontrado, datos nulos o corruptos).
- Usa docstrings y comentarios breves pero claros en español.
- Nombra las variables y funciones en español usando `snake_case` de manera muy descriptiva (ej. `total_horas_invertidas`, no `total`).

# Comandos de Ejecución
- Instalar dependencias: `pip install -r requirements.txt`
- Ejecutar script principal: `python generador_reportes.py`
