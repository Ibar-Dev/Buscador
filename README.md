# 🔍 Buscador Avanzado (v0.5.8)

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![version](https://img.shields.io/badge/version-0.5.8-informational)

Una potente aplicación de escritorio para realizar búsquedas avanzadas en archivos Excel. Ideal para analizar diccionarios de datos y descripciones extensas, con soporte para operadores lógicos, comparaciones numéricas y exportación de resultados.

---

## 📖 Tabla de Contenidos

- [Características Principales](#✨-características-principales)
- [Requisitos Previos](#🛠️-requisitos-previos)
- [Instalación](#⚙️-instalación)
- [Modo de Uso](#🚀-modo-de-uso)
  - [Cargar Archivos](#1-cargar-archivos)
  - [Realizar Búsquedas](#2-realizar-búsquedas)
  - [Exportar Resultados](#3-exportar-resultados)
- [Ejemplos de Búsqueda](#💡-ejemplos-de-búsqueda)
- [Configuración Automática](#🔧-configuración-automática)
- [Capturas de Pantalla](#🖼️-capturas-de-pantalla)
- [Licencia](#📜-licencia)
- [Notas de la Versión (v0.5.8)](#📄-notas-de-la-versión-v058)
- [Cómo Contribuir](#🤝-cómo-contribuir)
- [Ayuda Adicional](#❓-ayuda-adicional)

---

## ✨ Características Principales

-   **Búsqueda Especializada**: Carga archivos Excel diferenciados para términos de referencia (diccionario) y los datos donde se realizará la búsqueda (descripciones).
-   **Operadores de Búsqueda Avanzados**:
    -   Lógicos: `+` (AND), `|` o `/` (OR).
    -   Comparaciones Numéricas: `>`, `<`, `>=`, `<=`.
    -   Rangos Numéricos: `num1-num2` (ej: `10-20`).
    -   Negación de Términos: `#término_a_excluir`.
-   **Exportación Flexible**: Permite guardar tanto las reglas de búsqueda individuales como un consolidado de todos los resultados en formato Excel.
-   **Interfaz Gráfica Intuitiva**: Desarrollada con Tkinter para una experiencia de usuario sencilla y directa.

---

## 🛠️ Requisitos Previos

-   Python 3.7 o superior.
-   Las siguientes librerías de Python:
    -   `pandas`
    -   `openpyxl` (necesaria para la manipulación de archivos `.xlsx`)
    -   `tkinter` (generalmente incluido en la instalación estándar de Python)

---

## ⚙️ Instalación

1.  **Clona el repositorio (si aplica) o descarga los archivos fuente.**
    ```bash
    # Ejemplo si estuviera en un repositorio Git
    # git clone https://tu_repositorio/buscador_avanzado.git
    # cd buscador_avanzado
    ```

2.  **Se recomienda crear y activar un entorno virtual:**
    ```bash
    python -m venv venv
    # En Windows
    venv\Scripts\activate
    # En macOS/Linux
    source venv/bin/activate
    ```

3.  **Instala las dependencias necesarias:**
    ```bash
    pip install pandas openpyxl
    ```

---

## 🚀 Modo de Uso

1.  **Ejecuta la aplicación:**
    ```bash
    python tu_script_principal.py
    ```
    *(Reemplaza `tu_script_principal.py` con el nombre real de tu archivo de entrada)*

### 1. Cargar Archivos
    -   **Diccionario**: Selecciona el archivo Excel que contiene los términos de referencia.
    -   **Descripciones**: Selecciona el archivo Excel con los datos sobre los cuales deseas realizar la búsqueda.

### 2. Realizar Búsquedas
    -   Introduce tus términos de búsqueda en el campo designado.
    -   Utiliza los operadores avanzados para refinar tus consultas. Consulta la sección de "Ayuda" (botón `?`) dentro de la aplicación para una guía detallada de los operadores.

### 3. Exportar Resultados
    -   **Guardar Regla**: Guarda la regla de búsqueda actual y sus resultados en un archivo Excel.
    -   **Exportar Todo**: Exporta todas las reglas de búsqueda que hayas guardado durante la sesión a un único archivo Excel.

---

## 💡 Ejemplos de Búsqueda

-   `router + cisco`: Encuentra descripciones que contengan **ambos** términos "router" Y "cisco".
-   `switch | ap / ubiquiti`: Busca descripciones que contengan "switch" O "ap" O "ubiquiti".
-   `>1000w`: Localiza valores numéricos estrictamente mayores a 1000, asociados a la unidad "w" (vatios).
-   `10-20 puertos`: Identifica rangos numéricos entre 10 y 20 (inclusivo) seguidos del término "puertos".
-   `switch + #gestionable`: Busca el término "switch" pero excluye aquellas entradas que también contengan "gestionable".

---

## 🔧 Configuración Automática

La aplicación gestiona y guarda automáticamente la siguiente información para mejorar tu experiencia en sesiones futuras:
-   Las rutas de los últimos archivos de Diccionario y Descripciones cargados.
-   Los índices de las columnas seleccionadas para la búsqueda en el diccionario.

---

## 🖼️ Capturas de Pantalla

*(¡Es el momento ideal para mostrar tu aplicación en acción!)*

*Ejemplo de cómo podrías añadir una imagen:*
*Sugerencia: Coloca tus imágenes en una carpeta (ej. `docs/images/`) dentro de tu proyecto y enlaza a ellas.*

---

## 📜 Licencia

Este proyecto se distribuye bajo la Licencia MIT. Consulta el archivo [LICENSE](LICENSE) para más detalles.

---

## 📄 Notas de la Versión (v0.5.8)

-   Mejoras significativas en la validación interna de los operadores de búsqueda.
-   Implementado el soporte para unidades en comparaciones numéricas (ej: `>1000w`, `<50kg`).
-   Optimización del rendimiento general, especialmente notable en búsquedas sobre grandes volúmenes de datos.

---

## 🤝 Cómo Contribuir

¡Las contribuciones son siempre bienvenidas! Si tienes ideas, sugerencias o encuentras algún error:
1.  Abre un "Issue" en el repositorio para discutir cambios o reportar bugs.
2.  Si deseas aportar código, por favor, haz un "Fork" del repositorio y envía un "Pull Request" con tus mejoras.

---

## ❓ Ayuda Adicional

> **Nota Importante**: Para una guía detallada sobre el uso de los operadores de búsqueda y otras funcionalidades, por favor, utiliza el botón de ayuda (`?`) integrado en la aplicación.
