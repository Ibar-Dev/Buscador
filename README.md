# Buscador 
readme_content = """# Buscador v0.7.5

![Python](https://img.shields.io/badge/Python-3.6%2B-blue)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green)
![Pandas](https://img.shields.io/badge/Data-Pandas-yellow)

## 📋 Descripción

Buscador es una aplicación de escritorio para búsquedas avanzadas en archivos Excel. Permite realizar consultas complejas utilizando operadores lógicos, comparaciones numéricas, rangos y negaciones. La aplicación está diseñada para trabajar con dos tipos de archivos:

- **Diccionario**: Archivo Excel que contiene términos de referencia.
- **Descripciones**: Archivo Excel con datos que se desean consultar.

La aplicación puede buscar directamente en las descripciones o utilizar el diccionario como intermediario para encontrar coincidencias más relevantes.

## ✨ Características

- Interfaz gráfica intuitiva basada en Tkinter
- Carga y visualización de archivos Excel (.xlsx, .xls)
- Búsqueda avanzada con múltiples operadores:
  - Operadores lógicos: AND (`+`), OR (`|` o `/`)
  - Comparaciones numéricas: `>`, `<`, `>=`, `<=`, `=`
  - Rangos numéricos: `num1 - num2`
  - Negación (exclusión): `#palabra` o `#"frase completa"`
- Normalización de texto para búsquedas insensibles a mayúsculas/minúsculas y acentos
- Exportación de resultados a Excel o CSV
- Guardado de reglas/búsquedas para uso posterior
- Configuración persistente entre sesiones

## 🔧 Requisitos

- Python 3.6 o superior
- Dependencias:
  - pandas
  - openpyxl (para archivos .xlsx)
  - tkinter (incluido en la mayoría de instalaciones de Python)

## 📦 Instalación

1. Asegúrese de tener Python 3.6+ instalado
2. Instale las dependencias:

```bash
pip install pandas openpyxl
