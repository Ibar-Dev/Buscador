
# Buscador Avanzado de Excel

Aplicación de escritorio desarrollada en Python para búsquedas avanzadas en archivos Excel, especializada en relacionar términos entre diccionarios y descripciones.

![Demo Interface](placeholder.jpg) <!-- Agregar imagen de demo si está disponible -->

## Características Principales

### 🖥 Interfaz Gráfica
- Desarrollada con `tkinter` y `ttk`
- Adaptable al tema del sistema operativo
- Vista previa de datos integrada

### 🔍 Búsqueda Avanzada
- Búsqueda simple de términos
- Operadores lógicos:
  - **AND**: `palabra1+palabra2`
  - **OR**: `palabra1-palabra2` 
- Búsqueda directa en descripciones

### 📁 Manejo de Archivos
- Soporte para formatos:
  - `.xlsx` (OpenPyXL)
  - `.xls` (xlwt)
- Exportación de resultados múltiples:
  - Excel (.xlsx)
  - CSV (UTF-8)
  - Excel 97-2003 (.xls)

## Requisitos del Sistema

### Dependencias principales
```bash
pip install pandas openpyxl
