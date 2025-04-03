# Buscador
Agregado el archivo principal del buscador de Excel a DataFrame con Tkinter.


# Buscador en Excel con Interfaz Gráfica

Este proyecto consiste en una aplicación de escritorio desarrollada en Python utilizando la biblioteca Tkinter. Permite a los usuarios cargar archivos Excel, realizar búsquedas avanzadas dentro de sus datos y comparar dos archivos Excel.

## Funcionalidades Principales

* **Cargar Archivos Excel:** Permite seleccionar y cargar archivos Excel (`.xlsx` o `.xls`).
* **Comparar Archivos Excel:** Carga un segundo archivo Excel y compara si su contenido es idéntico al primero.
* **Búsqueda Avanzada:**
    * Realiza búsquedas de texto dentro de todas las columnas del archivo Excel cargado.
    * Soporta búsquedas simples.
    * Implementa búsquedas con operadores "AND" (utilizando el símbolo `+` entre las palabras clave).
    * Implementa búsquedas con operadores "OR" (utilizando el símbolo `-` entre las palabras clave).
* **Visualización de Resultados:** Los resultados de la búsqueda se muestran en una tabla interactiva dentro de la aplicación.
* **Exportar Resultados (No Implementado):** El botón para exportar los resultados a un archivo Excel está presente, pero la funcionalidad aún no ha sido implementada.

## Dependencias

Para ejecutar esta aplicación, necesitas tener instaladas las siguientes bibliotecas de Python:

* **pandas:** Para la manipulación y análisis de datos tabulares (archivos Excel).
* **tkinter:** Para la creación de la interfaz gráfica (generalmente viene incluido con la instalación de Python).
* **httpx:** Aunque importado, no parece ser utilizado en la versión actual del código. Podría ser una dependencia de desarrollo o para una funcionalidad futura.

Puedes instalar estas dependencias utilizando pip:

```bash
pip install [dependencia deseada]
