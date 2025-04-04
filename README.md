# Buscador Avanzado Optimizado

## Descripción
Buscador Avanzado Optimizado es una aplicación de escritorio desarrollada en Python que permite cargar, buscar y comparar datos entre archivos Excel. Está diseñada para facilitar la búsqueda de términos específicos en grandes conjuntos de datos y exportar los resultados.

## Características
- Carga de archivos Excel como diccionario principal y archivo de descripciones para comparar
- Búsqueda avanzada con operadores '+' (AND) y '-' (OR)
- Visualización de datos en tablas con desplazamiento horizontal y vertical
- Exportación de resultados en formatos Excel (.xlsx, .xls) y CSV
- Interfaz gráfica intuitiva con barras de estado informativas

## Requisitos
- Python 3.6 o superior
- Bibliotecas:
  - tkinter
  - pandas
  - xlsxwriter

## Instalación
1. Asegúrese de tener Python instalado en su sistema
2. Instale las dependencias necesarias:
```
pip install pandas xlsxwriter
```
3. Ejecute la aplicación:
```
python buscador_avanzado.py
```

## Uso
1. **Cargar Diccionario**: Cargue el archivo Excel principal que servirá como base de datos para las búsquedas.
2. **Cargar Descripciones**: (Opcional) Cargue un segundo archivo Excel para comparar con el diccionario.
3. **Buscar**: Introduzca términos de búsqueda en el campo "REGLAS a ensayar".
   - Use '+' entre palabras para buscar coincidencias que contengan TODAS las palabras (operador AND)
   - Use '-' entre palabras para buscar coincidencias que contengan ALGUNA de las palabras (operador OR)
4. **Exportar REGLAS**: Exporte los resultados de la búsqueda a un archivo Excel o CSV.

## Sintaxis de búsqueda
- **Búsqueda simple**: Escriba una palabra o frase para encontrar coincidencias exactas.
- **Búsqueda AND**: Use '+' entre términos (ej. "término1+término2") para encontrar filas que contengan ambos términos.
- **Búsqueda OR**: Use '-' entre términos (ej. "término1-término2") para encontrar filas que contengan al menos uno de los términos.

## Estructura del código
- **ManejadorExcel**: Clase para operaciones con archivos Excel.
- **MotorBusqueda**: Gestiona la lógica de búsqueda y manipulación de datos.
- **InterfazGrafica**: Maneja la interfaz gráfica de la aplicación.

## Licencia
Este proyecto está disponible como software de código abierto.

## Contacto
Para soporte o sugerencias, por favor contacte al desarrollador.
