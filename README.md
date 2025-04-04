# Buscador Avanzado Optimizado de Archivos Excel

Este proyecto es una aplicación de escritorio desarrollada en Python utilizando la biblioteca Tkinter para la interfaz gráfica y pandas para el manejo de archivos Excel. Permite a los usuarios cargar archivos Excel, realizar búsquedas avanzadas dentro de ellos y comparar el contenido de dos archivos.

## Características Principales

* **Carga de Archivos Excel:** Permite cargar dos archivos Excel: uno para realizar la búsqueda y otro opcional para comparar y priorizar resultados.
* **Búsqueda Avanzada:**
    * Búsqueda simple de un término en todas las columnas del archivo cargado.
    * Búsqueda con operador **AND** utilizando el símbolo `+` (ejemplo: `palabra1+palabra2`). Encontrará filas que contengan todas las palabras especificadas.
    * Búsqueda con operador **OR** utilizando el símbolo `-` (ejemplo: `palabra1-palabra2`). Encontrará filas que contengan al menos una de las palabras especificadas.
    * La búsqueda no distingue entre mayúsculas y minúsculas.
* **Comparación de Archivos:** Permite cargar un segundo archivo Excel para comparar su contenido con el primero. Se muestra un mensaje indicando si los archivos son idénticos o diferentes. La búsqueda priorizará los resultados encontrados en el archivo de comparación (si está cargado).
* **Visualización de Datos:** Muestra los datos cargados y los resultados de la búsqueda en tablas interactivas dentro de la aplicación.
* **Exportación de Resultados:** Permite exportar los resultados de la búsqueda a archivos Excel (`.xlsx`, `.xls`) o CSV (`.csv`).
* **Interfaz Gráfica Intuitiva:** Interfaz de usuario fácil de usar para una experiencia fluida.
* **Barra de Estado:** Proporciona información sobre el estado de la aplicación y las operaciones en curso.

## Cómo Utilizar

### Requisitos

* **Python 3.x** instalado en su sistema.
* Las siguientes bibliotecas de Python deben estar instaladas:
    * **tkinter:** Generalmente viene preinstalado con Python.
    * **pandas:** Para el manejo de archivos Excel. Puede instalarlo con: `pip install pandas`
    * **openpyxl:** Necesario para leer y escribir archivos Excel `.xlsx`. Puede instalarlo con: `pip install openpyxl`
    * **xlsxwriter:** (Opcional, pero recomendado para una mejor compatibilidad al exportar a Excel). Puede instalarlo con: `pip install xlsxwriter`

### Ejecución

1.  Guarde el código Python proporcionado en un archivo con extensión `.py` (por ejemplo, `buscador_excel.py`).
2.  Abra una terminal o símbolo del sistema.
3.  Navegue hasta el directorio donde guardó el archivo.
4.  Ejecute la aplicación con el comando: `python buscador_excel.py`

### Pasos

1.  **Cargar Archivo Buscador:** Haga clic en el botón "Cargar Excel Buscador" y seleccione el archivo Excel en el que desea realizar la búsqueda. Los datos del archivo se mostrarán en la tabla superior.
2.  **Cargar Archivo a Comparar (Opcional):** Si desea comparar un segundo archivo o priorizar resultados de otro archivo, haga clic en "Cargar Excel a Comparar" y seleccione el archivo. Se mostrará un mensaje indicando si los archivos son idénticos o diferentes, y los datos del segundo archivo se mostrarán en la tabla inferior (inicialmente).
3.  **Introducir Término de Búsqueda:** Escriba el término o términos que desea buscar en el campo de texto "Término/s de búsqueda:".
    * Para búsqueda simple, escriba una sola palabra o frase.
    * Para búsqueda con **AND**, separe las palabras con el símbolo `+` (ej: `nombre+apellido`).
    * Para búsqueda con **OR**, separe las palabras con el símbolo `-` (ej: `producto-servicio`).
4.  **Buscar:** Haga clic en el botón "Buscar" o presione la tecla Enter. Los resultados de la búsqueda se mostrarán en la tabla inferior.
5.  **Exportar Resultados (Opcional):** Si desea guardar los resultados de la búsqueda, haga clic en el botón "Exportar Resultados". Se le pedirá que elija un nombre de archivo y el formato (Excel o CSV).

## Estructura del Código

El código se organiza en las siguientes clases:

* **`ManejadorExcel`:** Contiene métodos estáticos para cargar archivos Excel de forma segura utilizando pandas.
* **`MotorBusqueda`:** Gestiona la lógica de búsqueda y almacenamiento de los datos cargados. Implementa la funcionalidad de búsqueda con los operadores AND (`+`) y OR (`-`).
* **`InterfazGrafica`:** Crea la ventana principal de la aplicación y define todos los widgets (botones, etiquetas, tablas, etc.) y su comportamiento.

## Notas Adicionales

* La aplicación muestra un máximo de 3 resultados iniciales en la tabla de resultados para mejorar el rendimiento con grandes conjuntos de datos. La barra de estado indicará el número total de coincidencias encontradas.
* Se utilizan mensajes de error y advertencia para proporcionar retroalimentación al usuario en caso de problemas al cargar archivos o durante la búsqueda.
* La interfaz gráfica está diseñada para ser responsive dentro de los límites definidos.

Este proyecto proporciona una herramienta útil para buscar y comparar información dentro de archivos Excel de manera eficiente.
