# 📊 Buscador Avanzado (v0.7.0)

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange)
![Dependencies](https://img.shields.io/badge/dependencies-pandas%20%7C%20openpyxl-brightgreen)
![Version](https://img.shields.io/badge/version-0.6.5-informational)

Aplicación de escritorio con interfaz gráfica (Tkinter) diseñada para realizar búsquedas complejas y estructuradas en archivos Excel. Facilita el cruce de información entre un archivo "Diccionario" de referencia y un archivo "Descripciones" de datos, implementando una lógica de parseo jerárquica para los términos de búsqueda. Permite guardar y exportar los resultados y las reglas de búsqueda.

---

## 📖 Tabla de Contenidos

- [✨ Características Destacadas](#-características-destacadas)
- [🏗️ Arquitectura del Software](#️-arquitectura-del-software)
  - [Clases Principales](#clases-principales)
  - [Enumeraciones](#enumeraciones)
  - [Lógica de Parseo de Búsqueda](#lógica-de-parseo-de-búsqueda)
- [🛠️ Requisitos Previos](#️-requisitos-previos)
- [⚙️ Instalación](#️-instalación)
- [🔧 Configuración](#️-configuración)
  - [Archivo de Configuración](#archivo-de-configuración-config_buscador_v0_6_5json)
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
- [🧪 Pruebas Unitarias](#-pruebas-unitarias)
- [📜 Licencia](#-licencia)
- [🤝 Cómo Contribuir](#-cómo-contribuir)

---

## ✨ Características Destacadas

-   **Interfaz Gráfica con Tkinter**: Interfaz de usuario intuitiva para facilitar la interacción, carga de archivos y visualización de datos.
-   **Parseo Jerárquico de Términos de Búsqueda**:
    -   Prioriza operadores `OR` (`|`, `/`) sobre operadores `AND` (`+`).
    -   Permite búsquedas complejas como `terminoA | terminoB + terminoC` interpretada correctamente como `terminoA OR (terminoB AND terminoC)`.
-   **Búsqueda en Dos Etapas o Directa**:
    -   Utiliza un archivo "Diccionario" para identificar filas coincidentes (FCDs) aplicando la lógica de parseo jerárquico.
    -   Extrae términos clave de estos FCDs y luego busca dichos términos (combinados con `OR`) en un archivo "Descripciones".
    -   Permite una búsqueda directa en el archivo "Descripciones" (aplicando también el parseo jerárquico) si no se encuentran FCDs o si el usuario opta por esta vía.
-   **Operadores de Búsqueda Avanzados**:
    -   Lógicos: `+` (AND), `|` o `/` (OR).
    -   Comparaciones Numéricas: `>`, `<`, `>=`, `<=` (ej. `>1000`, `<50V`).
    -   Rangos Numéricos: `num1-num2` (ej. `10-20KM`).
    -   Negación de términos individuales o expresiones: `#termino_a_excluir` (ej. `switch + #gestionable`).
    -   Soporte para Unidades: Reconoce y filtra por unidades junto a valores numéricos (ej: `>1000W`, `<=10.5A`).
-   **Extractor de Magnitudes**: Capacidad para identificar y normalizar unidades de medida predefinidas para comparaciones numéricas precisas.
-   **Gestión de Reglas de Búsqueda**:
    -   Permite guardar reglas (término original, estructura parseada completa, origen de los datos y los propios datos resultantes).
    -   Opción de guardar las coincidencias del diccionario (FCDs) y/o los resultados finales de las descripciones (RFDs) para búsquedas vía diccionario.
    -   Exporta todas las reglas guardadas a un único archivo Excel, con una hoja de índice y hojas detalladas para la definición (incluyendo la estructura parseada) y los datos de cada regla.
-   **Configuración Persistente**: Guarda las rutas de los últimos archivos utilizados y la configuración de las columnas de búsqueda del diccionario en `config_buscador_v0_6_5.json`.
-   **Visualización de Datos Mejorada**:
    -   Muestra vistas previas del "Archivo Diccionario" (limitado a 100 filas para rendimiento) y los resultados completos en tablas (`Treeview`).
    -   Permite el ordenamiento dinámico de las columnas en ambas tablas haciendo clic en sus cabeceras.
-   **Manejo de Errores y Logging**:
    -   Sistema de mensajes informativos y de error al usuario a través de diálogos.
    -   Registro detallado de operaciones, advertencias y errores en el archivo `buscador_app_v0_6_5.log` para facilitar la depuración.
-   **Validación Dinámica de Operadores**: La interfaz habilita/deshabilita los botones de inserción de operadores de búsqueda (`+`, `|`, `#`, `>`, etc.) en tiempo real, basándose en la validez y contexto del término de búsqueda que se está escribiendo.
-   **Ayuda Integrada**: Proporciona una guía de sintaxis accesible mediante un botón de ayuda (`?`).

---

## 🏗️ Arquitectura del Software

El script está estructurado modularmente para una mejor organización y mantenibilidad:

### Clases Principales
-   **`InterfazGrafica(tk.Tk)`**: Gestiona todos los aspectos de la interfaz de usuario (ventanas, widgets, eventos), coordina las interacciones y presenta los datos. Es el punto de entrada de la aplicación.
-   **`MotorBusqueda`**: Contiene toda la lógica central para procesar las búsquedas. Esto incluye:
    -   Carga y validación de los DataFrames de Pandas.
    -   **Parseo jerárquico de dos niveles** de los términos de búsqueda.
    -   Análisis de términos atómicos para identificar negaciones, comparaciones, rangos y texto simple.
    -   Generación de máscaras de filtrado para los DataFrames.
    -   Extracción de términos clave del "Diccionario" para la búsqueda en dos etapas.
    -   Orquestación del flujo de búsqueda (vía diccionario o directa).
-   **`ManejadorExcel`**: Clase de utilidad estática responsable de la carga de archivos Excel (`.xls`, `.xlsx`) utilizando `pandas`.
-   **`ExtractorMagnitud`**: Utilidad para la normalización y reconocimiento de unidades de medida.

### Enumeraciones
-   **`OrigenResultados(Enum)`**: Define y categoriza los diferentes flujos o estados por los cuales se pueden obtener los resultados de una búsqueda (ej., `VIA_DICCIONARIO_CON_RESULTADOS_DESC`, `DIRECTO_DESCRIPCION_VACIA`, `ERROR_CARGA_DICCIONARIO`, `TERMINO_INVALIDO`).

### Lógica de Parseo de Búsqueda
El `MotorBusqueda` implementa un sistema de parseo de dos niveles:
1.  **`_parsear_nivel1_or`**: Divide el término de búsqueda completo en segmentos basados en los operadores `OR` (`|`, `/`), otorgándoles la máxima precedencia.
2.  **`_parsear_nivel2_and`**: Cada segmento del nivel 1 se procesa para dividirlo por el operador `AND` (`+`). Esta función utiliza una máquina de estados para manejar correctamente las expresiones de magnitud y evitar divisiones incorrectas dentro de números o unidades.
3.  **`_analizar_terminos`**: Los términos resultantes del parseo de nivel 2 (que son atómicos o expresiones simples de comparación/rango/negación) se analizan para determinar su tipo y valor.
4.  **Generación de Máscaras**: Se generan máscaras booleanas para los términos analizados y se combinan según la jerarquía de operadores (AND dentro de los segmentos, luego OR entre segmentos).

---

## 🛠️ Requisitos Previos

-   **Python 3.7 o superior.**
-   **pip** (gestor de paquetes de Python).
-   **Librerías de Python**:
    -   `pandas`
    -   `openpyxl` (para leer/escribir archivos `.xlsx`)
-   **Tkinter**:
    -   Normalmente incluido con Python. En Linux, puede necesitar `sudo apt-get install python3-tk`.

---

## ⚙️ Instalación

1.  **Descarga el Script**:
    Obtén el archivo Python `Buscador_v0_6_5.py`.

2.  **Entorno Virtual (Recomendado)**:
    ```bash
    python -m venv venv_buscador
    # Windows:
    venv_buscador\Scripts\activate
    # macOS/Linux:
    source venv_buscador/bin/activate
    ```

3.  **Instala Dependencias**:
    Crea `requirements.txt`:
    ```text
    pandas
    openpyxl
    ```
    Instala:
    ```bash
    pip install -r requirements.txt
    ```

---

## 🔧 Configuración

### Archivo de Configuración (`config_buscador_v0_6_5.json`)
La aplicación crea y utiliza `config_buscador_v0_6_5.json` para almacenar:
-   `last_dic_path`: Ruta al último "Archivo Diccionario".
-   `last_desc_path`: Ruta al último "Archivo de Descripciones".
-   `indices_columnas_busqueda_dic`: Lista de índices de columnas a usar del "Diccionario".

### Columnas del Diccionario
En `config_buscador_v0_6_5.json`, la clave `indices_columnas_busqueda_dic`:
-   `[]` o `[-1]`: Busca en todas las columnas de texto/objeto del diccionario.
-   `[0, 2, 5]`: Busca solo en la primera, tercera y sexta columna.

---

## 🚀 Ejecución de la Aplicación

Con el entorno virtual activado (si se usó):
```bash
python Buscador_v0_6_5.py
