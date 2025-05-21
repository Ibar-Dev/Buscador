# 📊 Buscador(v0.6.5)

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange)
![Dependencies](https://img.shields.io/badge/dependencies-pandas%20%7C%20openpyxl-brightgreen)
![Version](https://img.shields.io/badge/version-0.5.8-informational)

Aplicación de escritorio con interfaz gráfica (Tkinter) diseñada para realizar búsquedas complejas y estructuradas en archivos Excel. Facilita el cruce de información entre un archivo "Diccionario" de referencia y un archivo "Descripciones" de datos, permitiendo guardar y exportar los resultados y las reglas de búsqueda.

---

## 📖 Tabla de Contenidos

- [✨ Características Destacadas](#-características-destacadas)
- [🏗️ Arquitectura del Software](#️-arquitectura-del-software)
  - [Clases Principales](#clases-principales)
  - [Enumeraciones](#enumeraciones)
- [🛠️ Requisitos Previos](#️-requisitos-previos)
- [⚙️ Instalación](#️-instalación)
- [🔧 Configuración](#️-configuración)
  - [Archivo de Configuración](#archivo-de-configuración-config_buscador_json)
  - [Columnas del Diccionario](#columnas-del-diccionario)
- [🚀 Ejecución de la Aplicación](#️-ejecución-de-la-aplicación)
- [🗺️ Guía de Uso](#️-guía-de-uso)
  - [1. Carga de Archivos](#1-carga-de-archivos)
  - [2. Realización de Búsquedas](#2-realización-de-búsquedas)
  - [Lógica de Búsqueda Detallada](#lógica-de-búsqueda-detallada)
  - [3. Guardar Reglas de Búsqueda](#3-guardar-reglas-de-búsqueda)
  - [4. Exportar Reglas Guardadas](#4-exportar-reglas-guardadas)
  - [5. Ayuda Integrada](#5-ayuda-integrada)
- [🔍 Sintaxis de Búsqueda](#-sintaxis-de-búsqueda)
- [📝 Registro de Actividad (Logging)](#-registro-de-actividad-logging)
- [🖼️ Capturas de Pantalla (Sugerencia)](#️-capturas-de-pantalla-sugerencia)
- [📜 Licencia](#-licencia)
- [🤝 Cómo Contribuir](#-cómo-contribuir)

---

## ✨ Características Destacadas

-   **Interfaz Gráfica con Tkinter**: Interfaz de usuario intuitiva para facilitar la interacción, carga de archivos y visualización de datos.
-   **Búsqueda en Dos Etapas o Directa**:
    -   Utiliza un archivo "Diccionario" para identificar filas coincidentes (FCDs), extrae términos clave de estas y luego busca dichos términos en un archivo "Descripciones".
    -   Permite una búsqueda directa en el archivo "Descripciones" si no se encuentran FCDs o si el usuario opta por esta vía.
-   **Operadores de Búsqueda Avanzados**:
    -   Lógicos: `+` (AND), `|` o `/` (OR).
    -   Comparaciones Numéricas: `>`, `<`, `>=`, `<=` (ej. `>1000`, `<50V`).
    -   Rangos Numéricos: `num1-num2` (ej. `10-20KM`).
    -   Negación de términos individuales o expresiones: `#termino_a_excluir` (ej. `switch + #gestionable`).
    -   Soporte para Unidades: Reconoce y filtra por unidades junto a valores numéricos (ej: `>1000W`, `<=10.5A`).
-   **Extractor de Magnitudes**: Capacidad para identificar y normalizar unidades de medida predefinidas para comparaciones numéricas precisas.
-   **Gestión de Reglas de Búsqueda**:
    -   Permite guardar reglas (término original, términos analizados, operador principal, origen de los datos y los propios datos resultantes).
    -   Opción de guardar las coincidencias del diccionario (FCDs) y/o los resultados finales de las descripciones (RFDs) para búsquedas vía diccionario.
    -   Exporta todas las reglas guardadas a un único archivo Excel, con una hoja de índice y hojas detalladas para la definición y los datos de cada regla.
-   **Configuración Persistente**: Guarda las rutas de los últimos archivos utilizados y la configuración de las columnas de búsqueda del diccionario en `config_buscador.json`.
-   **Visualización de Datos Mejorada**:
    -   Muestra vistas previas del "Archivo Diccionario" (limitado a 100 filas para rendimiento) y los resultados completos en tablas (`Treeview`).
    -   Permite el ordenamiento dinámico de las columnas en ambas tablas haciendo clic en sus cabeceras.
-   **Manejo de Errores y Logging**:
    -   Sistema de mensajes informativos y de error al usuario a través de diálogos.
    -   Registro detallado de operaciones, advertencias y errores en el archivo `buscador_app.log` para facilitar la depuración.
-   **Validación Dinámica de Operadores**: La interfaz habilita/deshabilita los botones de inserción de operadores de búsqueda (`+`, `|`, `#`, `>`, etc.) en tiempo real, basándose en la validez y contexto del término de búsqueda que se está escribiendo.
-   **Ayuda Integrada**: Proporciona una guía de sintaxis accesible mediante un botón de ayuda (`?`).

---

## 🏗️ Arquitectura del Software

El script está estructurado modularmente para una mejor organización y mantenibilidad:

### Clases Principales
-   **`InterfazGrafica(tk.Tk)`**: Gestiona todos los aspectos de la interfaz de usuario (ventanas, widgets, eventos), coordina las interacciones y presenta los datos. Es el punto de entrada de la aplicación.
-   **`MotorBusqueda`**: Contiene toda la lógica central para procesar las búsquedas. Esto incluye:
    -   Carga y validación de los DataFrames de Pandas.
    -   Parseo y análisis de los términos de búsqueda introducidos por el usuario.
    -   Generación de máscaras de filtrado para los DataFrames basadas en los términos analizados.
    -   Extracción de términos clave del "Diccionario" para la búsqueda en dos etapas.
    -   Orquestación del flujo de búsqueda (vía diccionario o directa).
-   **`ManejadorExcel`**: Clase de utilidad estática responsable de la carga de archivos Excel (`.xls`, `.xlsx`) utilizando `pandas`, devolviendo el DataFrame y posibles mensajes de error.
-   **`ExtractorMagnitud`**: Utilidad para la normalización y reconocimiento de unidades de medida (magnitudes) presentes en los textos, facilitando las comparaciones numéricas.

### Enumeraciones
-   **`OrigenResultados(Enum)`**: Define y categoriza los diferentes flujos o estados por los cuales se pueden obtener los resultados de una búsqueda (ej., `VIA_DICCIONARIO_CON_RESULTADOS_DESC`, `DIRECTO_DESCRIPCION_VACIA`, `ERROR_CARGA_DICCIONARIO`). Esto es crucial para la lógica interna, el guardado de reglas y la presentación de información al usuario.

---

## 🛠️ Requisitos Previos

-   **Python 3.7 o superior.**
-   **pip** (gestor de paquetes de Python).
-   **Librerías de Python** (se instalarán con `requirements.txt`):
    -   `pandas`
    -   `openpyxl` (necesario para que pandas lea y escriba archivos `.xlsx`)
-   **Tkinter**:
    -   Normalmente incluido con las instalaciones estándar de Python en Windows y macOS.
    -   En algunas distribuciones de Linux, puede requerir una instalación separada. Ejemplo para Debian/Ubuntu:
        ```bash
        sudo apt-get update
        sudo apt-get install python3-tk
        ```

---

## ⚙️ Instalación

1.  **Descarga o Clona el Script**:
    Obtén el archivo Python principal (ej. `buscador_excel_fusionado.py`).

2.  **Crea un Entorno Virtual (Altamente Recomendado)**:
    Navega a la carpeta donde guardaste el script y ejecuta:
    ```bash
    python -m venv venv_buscador
    ```
    Activa el entorno:
    -   Windows: `venv_buscador\Scripts\activate`
    -   macOS/Linux: `source venv_buscador/bin/activate`

3.  **Instala las Dependencias**:
    Crea un archivo llamado `requirements.txt` en el mismo directorio del script con el siguiente contenido:
    ```text
    pandas
    openpyxl
    ```
    Luego, instala las dependencias ejecutando:
    ```bash
    pip install -r requirements.txt
    ```

---

## 🔧 Configuración

### Archivo de Configuración (`config_buscador.json`)
La aplicación crea y utiliza automáticamente un archivo llamado `config_buscador.json` en el mismo directorio donde se ejecuta. Este archivo almacena:
-   `last_dic_path`: Ruta al último "Archivo Diccionario" cargado con éxito.
-   `last_desc_path`: Ruta al último "Archivo de Descripciones" cargado con éxito.
-   `indices_columnas_busqueda_dic`: Lista de índices de las columnas a utilizar del "Archivo Diccionario".

Este archivo se actualiza al cargar archivos o al cerrar la aplicación.

### Columnas del Diccionario
La clave `indices_columnas_busqueda_dic` en el archivo de configuración determina en qué columnas del "Archivo Diccionario" se buscarán los términos y de dónde se extraerán los términos clave para la segunda etapa de búsqueda (en descripciones).
-   Si el valor es una lista vacía `[]` (predeterminado si la clave no existe en el archivo de configuración) o `[-1]`, la aplicación buscará en **todas las columnas del diccionario que sean de tipo texto u objeto**.
-   Puedes especificar una lista de índices basados en cero, por ejemplo: `[0, 2, 5]`, para que la búsqueda en el diccionario se restrinja únicamente a la primera, tercera y sexta columna.

---

## 🚀 Ejecución de la Aplicación

Una vez instaladas las dependencias y con el entorno virtual activado (si creaste uno), ejecuta el script desde tu terminal:

```bash
python buscador_excel_fusionado.py
