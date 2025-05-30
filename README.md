# Buscador Avanzado Modularizado

## Descripción General

El Buscador Avanzado Modularizado es una aplicación de escritorio desarrollada en Python con una interfaz gráfica construida en Tkinter. Su propósito principal es permitir a los usuarios realizar búsquedas complejas y detalladas dentro de archivos de datos en formato Excel.

La aplicación utiliza dos archivos Excel principales:
1.  **Archivo de Diccionario**: Este archivo no solo contiene términos de búsqueda (Formas Canónicas del Diccionario - FCDs) y sus sinónimos, sino que también se utiliza para definir y normalizar unidades y magnitudes, mejorando la precisión de las búsquedas numéricas.
2.  **Archivo de Descripciones**: Es el archivo de datos principal donde se realizan las búsquedas. Los resultados finales se extraen de este archivo.

El motor de búsqueda soporta una sintaxis rica que incluye operadores lógicos (AND, OR), negación de términos, búsquedas numéricas específicas (comparaciones como >, <, >=, <=, =) con reconocimiento de unidades, búsquedas por rangos numéricos y búsqueda de frases exactas.

## Características Principales

* **Carga Dinámica de Archivos**: Permite al usuario cargar archivos Excel (`.xlsx`, `.xls`) para el diccionario y los datos de descripción en tiempo de ejecución.
* **Motor de Búsqueda Potente**:
    * **Operadores Lógicos**: Soporte para `+` (AND) y `|` (OR) para combinar términos de búsqueda.
    * **Negación**: Exclusión de términos o frases usando el prefijo `#` (ej., `#obsoleto` o `# "fuera de stock"`).
    * **Búsquedas Numéricas**:
        * Comparaciones: `>`, `<`, `>=`, `<=`, `=` (ej., `>100W`, `<=50V`). Las unidades son opcionales pero mejoran la precisión si se definen en el diccionario.
        * Rangos: Definición de rangos numéricos (ej., `10-20KG`, `2.5-3.0mm`).
    * **Frases Exactas**: Búsqueda de secuencias literales de palabras encerrándolas entre comillas dobles (ej., `"rack 19 pulgadas"`).
* **Normalización de Texto**: Para búsquedas más robustas, el texto de las consultas y de los datos se normaliza (conversión a mayúsculas, eliminación de tildes y caracteres especiales no relevantes).
* **Reconocimiento y Normalización de Magnitudes**: Utiliza el archivo de diccionario para construir un mapeo dinámico de unidades y sus sinónimos a una forma canónica, permitiendo búsquedas como `>10 voltios` aunque en los datos aparezca como `>10v`.
* **Interfaz Gráfica de Usuario (GUI)**:
    * Desarrollada con Tkinter, proporcionando una experiencia de usuario clara.
    * Visualización interactiva de los datos del diccionario y los resultados de la búsqueda en tablas.
    * Resaltado de las Formas Canónicas del Diccionario (FCDs) que coinciden con la consulta en la vista previa del diccionario.
* **Flujo de Búsqueda Flexible**:
    * **Vía Diccionario (por defecto)**: La consulta se busca primero en el diccionario. Los sinónimos encontrados se usan para buscar en las descripciones. Se aplica la condición numérica original (si la hubo en la query) y negaciones globales.
    * **Manejo de AND complejo**: Para consultas como `A + B`, se buscan sinónimos para A y para B en el diccionario, y luego se buscan descripciones que contengan sinónimos de A *Y* sinónimos de B.
    * **Flujo Alternativo por Unidad**: Si una búsqueda numérica con unidad (ej. `>10V`) no encuentra FCDs directos, el sistema intenta buscar FCDs que contengan solo la unidad (ej. `V`), y luego aplica el filtro numérico original (`>10`) a los sinónimos encontrados en las descripciones.
    * **Búsqueda Directa**: Opción para buscar la consulta original directamente en el archivo de descripciones si la búsqueda vía diccionario no produce los resultados esperados.
* **Exportación de Resultados**: Los resultados de la búsqueda (de la tabla de descripciones) pueden ser exportados a formatos `.xlsx` o `.csv`.
* **Configuración Persistente**: Guarda la ruta de los últimos archivos cargados y otras configuraciones de la UI en un archivo `config_buscador_avanzado_ui.json`.
* **Logging Detallado**: Registra las operaciones, advertencias y errores en un archivo de log (`Buscador_Avanzado_App_v1.10.3_Mod.log`) para facilitar la depuración y el seguimiento.

## Estructura del Proyecto

El proyecto está organizado en un paquete principal `buscador_app` y un script de entrada `main.py`.
