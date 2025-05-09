# 📊 Buscador Avanzado Excel (v0.5.8)

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green) ![GUI](https://img.shields.io/badge/GUI-Tkinter-orange)
![Dependencies](https://img.shields.io/badge/dependencies-pandas%20%7C%20openpyxl-brightgreen)
![version](https://img.shields.io/badge/version-0.5.8-informational)

Aplicación de escritorio con interfaz gráfica (Tkinter) diseñada para realizar búsquedas complejas y estructuradas en archivos Excel. Facilita el cruce de información entre un archivo "Diccionario" de referencia y un archivo "Descripciones" de datos, permitiendo guardar y exportar los resultados y las reglas de búsqueda.

---

## 📖 Tabla de Contenidos

- [Características Destacadas](#✨-características-destacadas)
- [Arquitectura del Software](#🏗️-arquitectura-del-software)
  - [Clases Principales](#clases-principales)
  - [Enumeraciones](#enumeraciones)
- [Requisitos Previos](#🛠️-requisitos-previos)
- [Instalación](#⚙️-instalación)
- [Configuración](#🔧-configuración)
  - [Archivo de Configuración](#archivo-de-configuración-config_buscadorjson)
  - [Columnas del Diccionario](#columnas-del-diccionario)
- [Ejecución de la Aplicación](#🚀-ejecución-de-la-aplicación)
- [Guía de Uso](#🗺️-guía-de-uso)
  - [Carga de Archivos](#1-carga-de-archivos)
  - [Realización de Búsquedas](#2-realización-de-búsquedas)
  - [Lógica de Búsqueda Detallada](#lógica-de-búsqueda-detallada)
  - [Guardar Reglas de Búsqueda](#3-guardar-reglas-de-búsqueda)
  - [Exportar Reglas Guardadas](#4-exportar-reglas-guardadas)
  - [Ayuda Integrada](#5-ayuda-integrada)
- [Sintaxis de Búsqueda](#🔍-sintaxis-de-búsqueda)
- [Registro de Actividad (Logging)](#📝-registro-de-actividad-logging)
- [Capturas de Pantalla (Sugerencia)](#🖼️-capturas-de-pantalla-sugerencia)
- [Licencia](#📜-licencia)
- [Cómo Contribuir](#🤝-cómo-contribuir)

---

## ✨ Características Destacadas

-   **Interfaz Gráfica con Tkinter**: Interfaz de usuario intuitiva para facilitar la interacción.
-   **Búsqueda en Dos Etapas**: Utiliza un archivo "Diccionario" para identificar términos clave y luego busca estos términos en un archivo "Descripciones".
-   **Operadores de Búsqueda Avanzados**:
    -   Lógicos: `+` (AND), `|` o `/` (OR).
    -   Comparaciones Numéricas: `>`, `<`, `>=`, `<=`.
    -   Rangos Numéricos: `num1-num2`.
    -   Negación: `#termino_a_excluir`.
    -   Soporte para Unidades: Reconoce unidades junto a valores numéricos (ej: `>1000w`, `<50kg`).
-   **Extractor de Magnitudes**: Capacidad para identificar y extraer cantidades numéricas asociadas a unidades predefinidas (ej: "16GB", "100W").
-   **Gestión de Reglas**: Permite guardar reglas de búsqueda (término, origen y resultados) y exportarlas a un único archivo Excel con múltiples hojas.
-   **Configuración Persistente**: Guarda las últimas rutas de archivos y la configuración de columnas del diccionario en un archivo `config_buscador.json`.
-   **Visualización de Datos**: Muestra vistas previas del diccionario y los resultados en tablas (Treeview) con ordenamiento por columnas.
-   **Manejo de Errores y Logging**: Sistema robusto de mensajes al usuario y registro detallado de operaciones y errores en `buscador_app.log`.
-   **Validación Dinámica de Operadores**: La interfaz habilita/deshabilita botones de operadores según la validez del término de búsqueda actual.

---

## 🏗️ Arquitectura del Software

El script está estructurado en varias clases y enumeraciones para organizar la lógica de la aplicación:

### Clases Principales
    -   `ManejadorExcel`: Encargada de la carga y validación inicial de archivos Excel (`.xls`, `.xlsx`) usando `pandas`.
    -   `MotorBusqueda`: Contiene la lógica central para realizar las búsquedas, parsear términos, aplicar filtros y gestionar los DataFrames.
    -   `ExtractorMagnitud`: Utilidad para identificar y extraer valores numéricos asociados a magnitudes y unidades específicas dentro de cadenas de texto.
    -   `InterfazGrafica`: Construye y gestiona todos los elementos de la interfaz de usuario (ventanas, botones, tablas, etc.) utilizando `tkinter` y `tkinter.ttk`. Coordina las interacciones del usuario con el `MotorBusqueda`.

### Enumeraciones
    -   `OrigenResultados`: Define los diferentes caminos o flujos por los cuales se pueden obtener y clasificar los resultados de una búsqueda (ej: vía diccionario con resultados, búsqueda directa, etc.). Esto es crucial para la lógica de guardado de reglas.

---

## 🛠️ Requisitos Previos

-   Python 3.7 o superior.
-   Librerías de Python:
    -   `pandas`
    -   `openpyxl` (para leer y escribir archivos `.xlsx`)
    -   `tkinter` (generalmente incluido con la instalación estándar de Python)

---

## ⚙️ Instalación

1.  **Descarga o Clona el Script**:
    Obtén el archivo Python principal (ej: `buscador_excel_avanzado.py`).

2.  **Crea un Entorno Virtual (Recomendado)**:
    ```bash
    python -m venv venv
    # En Windows
    venv\Scripts\activate
    # En macOS/Linux
    source venv/bin/activate
    ```

3.  **Instala las Dependencias**:
    ```bash
    pip install pandas openpyxl
    ```

---

## 🔧 Configuración

### Archivo de Configuración (`config_buscador.json`)
    La aplicación crea y utiliza un archivo `config_buscador.json` en el mismo directorio donde se ejecuta. Este archivo almacena:
    -   `last_dic_path`: Ruta al último archivo de Diccionario cargado.
    -   `last_desc_path`: Ruta al último archivo de Descripciones cargado.
    -   `indices_columnas_busqueda_dic`: Lista de índices (basados en 0) de las columnas a utilizar del archivo Diccionario para la búsqueda.

### Columnas del Diccionario
    Por defecto, la aplicación utiliza las columnas en los índices `[0, 3]` del archivo Diccionario para realizar las búsquedas y extraer términos. Puedes cambiar esto modificando el archivo `config_buscador.json` (si ya existe) o se guardará tu selección si la aplicación permite configurarlo vía GUI en el futuro.

---

## 🚀 Ejecución de la Aplicación

Una vez instaladas las dependencias, ejecuta el script desde tu terminal:

```bash
python tu_nombre_de_script.py
