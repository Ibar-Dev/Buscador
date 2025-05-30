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

Buscador_Modulado/
├── main.py                     # Punto de entrada principal de la aplicación.
├── README.md                   # Este archivo de documentación.
├── requirements.txt            # (Recomendado) Archivo con las dependencias del proyecto.
├── config_buscador_avanzado_ui.json # (Generado por la app) Guarda la configuración de la UI.
├── Buscador_Avanzado_App_v1.10.3_Mod.log # (Generado por la app) Archivo de logs.
│
└── buscador_app/               # Paquete principal de la aplicación.
├── init.py             # Hace de 'buscador_app' un paquete Python.
├── enums.py                # Define enumeraciones (ej. OrigenResultados).
├── utils.py                # Módulo para clases y funciones de utilidad.
│
├── core/                   # Subpaquete para la lógica central (motor de búsqueda).
│   ├── init.py         # Hace de 'core' un subpaquete.
│   └── motor_busqueda.py   # Contiene la clase MotorBusqueda.
│
└── gui/                    # Subpaquete para la interfaz gráfica de usuario.
├── init.py         # Hace de 'gui' un subpaquete.
└── interfaz_grafica.py # Contiene la clase InterfazGrafica.


### Descripción de Carpetas y Archivos:

* **`Buscador_Modulado/` (Directorio Raíz)**
    * **`main.py`**:
        * **Función**: Es el script que se ejecuta para iniciar la aplicación.
        * **Responsabilidades**:
            * Configura el sistema de logging.
            * Verifica la presencia de dependencias críticas (pandas, numpy, openpyxl).
            * Crea una instancia de `InterfazGrafica` del módulo `buscador_app.gui.interfaz_grafica`.
            * Inicia el bucle principal de Tkinter (`app.mainloop()`).
    * **`README.md`**: Este mismo archivo, proporcionando documentación completa sobre el proyecto.
    * **`requirements.txt`** (Recomendado): Aunque no se ha generado automáticamente, se recomienda crear este archivo para listar todas las dependencias externas del proyecto, facilitando la instalación en otros entornos (ej. `pip freeze > requirements.txt`).
    * **`config_buscador_avanzado_ui.json`**:
        * **Función**: Almacena configuraciones de la interfaz de usuario, como las rutas de los últimos archivos de diccionario y descripciones cargados, y la configuración de columnas para la vista previa del diccionario.
        * **Generación**: Se crea o actualiza automáticamente por la aplicación al cerrar o al cambiar ciertas configuraciones.
    * **`Buscador_Avanzado_App_v1.10.3_Mod.log`**:
        * **Función**: Archivo de texto donde se registran eventos importantes, advertencias y errores durante la ejecución de la aplicación. Muy útil para depuración.
        * **Generación**: Se crea o sobrescribe en cada inicio de la aplicación.

* **`buscador_app/` (Paquete Principal)**
    * **`__init__.py`**:
        * **Función**: Un archivo (generalmente vacío) que indica a Python que el directorio `buscador_app` debe ser tratado como un paquete, permitiendo importaciones modulares.
    * **`enums.py`**:
        * **Función**: Define enumeraciones personalizadas utilizadas en la aplicación.
        * **Contenido Principal**:
            * `OrigenResultados`: Una enumeración (`Enum`) que define los diferentes caminos o estados por los cuales se pueden obtener o no obtener resultados de búsqueda. Esto ayuda a la lógica interna y a la UI a reaccionar de manera diferente según cómo se haya desarrollado la búsqueda (ej., `VIA_DICCIONARIO_CON_RESULTADOS_DESC`, `DICCIONARIO_SIN_COINCIDENCIAS`, `ERROR_CARGA_DICCIONARIO`, etc.).
    * **`utils.py`**:
        * **Función**: Contiene clases y funciones de utilidad reutilizables por otros módulos de la aplicación, promoviendo el principio DRY (Don't Repeat Yourself).
        * **Contenido Principal**:
            * `ExtractorMagnitud`: Clase responsable de la normalización de texto para unidades y la gestión de un mapeo de sinónimos de unidades a sus formas canónicas. Es crucial para interpretar correctamente las unidades en las consultas numéricas y en los datos del diccionario. Se inicializa con un mapeo predefinido (actualmente vacío) pero se actualiza dinámicamente al cargar un archivo de diccionario.
            * `ManejadorExcel`: Clase con métodos estáticos (`@staticmethod`) para manejar la carga de archivos Excel. Encapsula la lógica de lectura de archivos `.xlsx` (usando `openpyxl`) y `.xls` (intentando con `xlrd`), incluyendo el manejo de errores comunes como `ImportError` (si falta la librería) o archivos no encontrados.

* **`buscador_app/core/` (Subpaquete del Núcleo)**
    * **`__init__.py`**: Hace de `core` un subpaquete de `buscador_app`.
    * **`motor_busqueda.py`**:
        * **Función**: Este es el cerebro de la aplicación. Contiene la clase `MotorBusqueda`, que encapsula toda la lógica de procesamiento de consultas, filtrado de datos y aplicación de las reglas de búsqueda.
        * **Clase `MotorBusqueda`**:
            * **Inicialización**: Configura patrones de expresiones regulares para parsear la sintaxis de búsqueda (comparaciones, rangos, negaciones, etc.) e inicializa una instancia de `ExtractorMagnitud`.
            * **Carga de Datos**: Métodos `cargar_excel_diccionario` y `cargar_excel_descripcion` que utilizan `ManejadorExcel` para cargar los DataFrames de Pandas y, en el caso del diccionario, actualizan dinámicamente el `ExtractorMagnitud`.
            * **Normalización y Parseo**: Métodos internos para normalizar texto (`_normalizar_para_busqueda`), extraer términos negados y positivos de la consulta (`_aplicar_negaciones_y_extraer_positivos`), y descomponer la consulta en niveles lógicos de OR y AND (`_descomponer_nivel1_or`, `_descomponer_nivel2_and`).
            * **Análisis de Términos**: El método `_analizar_terminos` clasifica cada parte de la consulta (después del parseo OR/AND) en tipo "string", "comparación numérica", o "rango numérico", identificando también la unidad asociada si la hay (usando `ExtractorMagnitud`).
            * **Parseo Numérico**: El método `_parse_numero` es una utilidad robusta para convertir strings que representan números (con varios formatos de separadores decimales o de miles) a `float`.
            * **Generación de Máscaras**: Métodos como `_generar_mascara_para_un_termino` y `_aplicar_mascara_combinada_para_segmento_and` crean máscaras booleanas de Pandas para filtrar los DataFrames según los criterios de búsqueda. Estos manejan la lógica de comparación de texto, numérica y de unidades.
            * **Procesamiento de Búsqueda**: El método central `_procesar_busqueda_en_df_objetivo` orquesta la aplicación de negaciones, el parseo de la consulta positiva y la aplicación de las máscaras OR/AND sobre un DataFrame objetivo.
            * **Método Principal `buscar`**: Es el método público que la interfaz gráfica llama. Determina el flujo de búsqueda (vía diccionario o directo), maneja la lógica de AND global, el flujo alternativo por unidad, y devuelve los resultados finales junto con un `OrigenResultados` y cualquier FCD relevante.

* **`buscador_app/gui/` (Subpaquete de la Interfaz Gráfica)**
    * **`__init__.py`**: Hace de `gui` un subpaquete de `buscador_app`.
    * **`interfaz_grafica.py`**:
        * **Función**: Define la clase `InterfazGrafica` que construye y gestiona la ventana principal de la aplicación y todos sus componentes visuales.
        * **Clase `InterfazGrafica` (hereda de `tk.Tk`)**:
            * **Inicialización**: Crea la ventana principal, inicializa una instancia de `MotorBusqueda`, carga la configuración de la aplicación y configura los widgets y el layout.
            * **Gestión de Configuración**: Métodos `_cargar_configuracion_app` y `_guardar_configuracion_app` para leer/escribir el archivo `config_buscador_avanzado_ui.json`.
            * **Creación de Widgets**: El método `_crear_widgets_app` define todos los elementos de la UI: botones para cargar archivos, campo de entrada para la búsqueda, botones de operadores, tablas (Treeviews) para mostrar el diccionario y los resultados, etiquetas de estado, etc.
            * **Layout**: El método `_configurar_grid_layout_app` organiza los widgets en la ventana usando el gestor de geometría `grid` de Tkinter.
            * **Manejo de Eventos**: Métodos de callback (ej., `_cargar_diccionario_ui`, `_ejecutar_busqueda_ui`, `_exportar_resultados_ui`, `_on_texto_busqueda_change`) que responden a las acciones del usuario.
            * **Actualización de la UI**: Métodos como `_actualizar_tabla_treeview_ui` para rellenar las tablas con datos, `_actualizar_mensaje_barra_estado` para mostrar mensajes al usuario, y `_actualizar_estado_general_botones_y_controles` para habilitar/deshabilitar controles según el estado de la aplicación.
            * **Interacción con el Motor**: Llama al método `buscar` del `MotorBusqueda` y procesa los resultados para mostrarlos en la UI, incluyendo el manejo de diferentes `OrigenResultados` para ofrecer búsquedas alternativas o mostrar mensajes de error.
            * **Funcionalidades Adicionales**: Implementa la ordenación de tablas al hacer clic en cabeceras, la exportación de resultados, la visualización de ayuda y la (actualmente en memoria) funcionalidad de "Salvar Regla".

### Interacción entre Módulos:

* `main.py` instancia `InterfazGrafica`.
* `InterfazGrafica` instancia `MotorBusqueda`.
* `InterfazGrafica` usa `OrigenResultados` de `enums.py` para interpretar los resultados del motor.
* `MotorBusqueda` usa `ExtractorMagnitud` y `ManejadorExcel` de `utils.py`, y también `OrigenResultados` de `enums.py`.
* Todos los módulos utilizan `logging` para registrar información.

## Tecnologías Utilizadas

* **Python 3**: Versión 3.7 o superior recomendada.
* **Tkinter**: Para la interfaz gráfica de usuario (incluido en la biblioteca estándar de Python). Se utilizan `tkinter.ttk` para widgets temáticos.
* **Pandas**: Para la manipulación eficiente de datos y la lectura de archivos Excel.
* **NumPy**: Utilizado por Pandas y directamente para algunas comparaciones numéricas (`np.isclose`).
* **Openpyxl**: Para leer y escribir archivos Excel en formato `.xlsx`.
* **Xlrd**: (Opcional) Para leer archivos Excel en el formato antiguo `.xls`. La aplicación intentará usarlo si `openpyxl` no puede manejar un archivo `.xls`.

## Requisitos Previos

* Python 3.7 o superior.
* `pip` (el gestor de paquetes de Python) para instalar las dependencias.

## Instalación y Configuración

1.  **Clonar o Descargar el Proyecto**:
    Si el proyecto está en un repositorio Git (ej. GitHub):
    ```bash
    git clone https://URL_DEL_REPOSITORIO/Buscador_Modulado.git
    cd Buscador_Modulado
    ```
    Si has descargado un archivo ZIP, extráelo y navega hasta el directorio raíz del proyecto (`Buscador_Modulado`).

2.  **(Recomendado) Crear un Entorno Virtual**:
    Es una buena práctica aislar las dependencias del proyecto:
    ```bash
    python -m venv venv
    ```
    Activa el entorno virtual:
    * En Windows:
        ```bash
        venv\Scripts\activate
        ```
    * En macOS/Linux:
        ```bash
        source venv/bin/activate
        ```

3.  **Instalar Dependencias**:
    Si se proporciona un archivo `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```
    De lo contrario, instala las bibliotecas necesarias manualmente:
    ```bash
    pip install pandas numpy openpyxl xlrd
    ```
    * `pandas`: Para la manipulación de datos.
    * `numpy`: Dependencia de Pandas y usado para comparaciones numéricas.
    * `openpyxl`: Necesario para leer y escribir archivos Excel `.xlsx`.
    * `xlrd`: Necesario para leer archivos Excel más antiguos `.xls`.

4.  **Configuración Adicional**:
    No se requiere ninguna configuración manual adicional antes del primer uso. La aplicación creará automáticamente:
    * `config_buscador_avanzado_ui.json`: Al cerrar la aplicación o al cambiar ciertas configuraciones (como cargar un archivo).
    * `Buscador_Avanzado_App_v1.10.3_Mod.log`: Al iniciar la aplicación.

## Uso

1.  **Ejecutar la Aplicación**:
    Asegúrate de estar en el directorio raíz del proyecto (`Buscador_Modulado/`) y que tu entorno virtual (si usas uno) esté activado. Luego, ejecuta:
    ```bash
    python main.py
    ```
    Se abrirá la ventana principal de la aplicación.

2.  **Carga de Archivos**:
    * **Cargar Diccionario**: Haz clic en el botón "Cargar Diccionario" y selecciona el archivo Excel que contiene tus formas canónicas, sinónimos e información de unidades. La primera columna se usa para las formas canónicas principales, y las columnas desde la cuarta en adelante se consideran sinónimos.
    * **Cargar Descripciones**: Haz clic en el botón "Cargar Descripciones" y selecciona el archivo Excel donde se realizará la búsqueda de los términos.

3.  **Realizar Búsquedas**:
    * **Campo de Búsqueda**: Introduce tu consulta en el campo de texto grande.
    * **Operadores**: Puedes usar los botones de operadores (`+`, `|`, `#`, `>`, `<`, etc.) para ayudarte a construir la consulta, o escribirlos directamente. Consulta la sección "Ayuda" (`?`) en la aplicación para una sintaxis detallada.
    * **Ejecutar Búsqueda**: Presiona el botón "Buscar" o la tecla `Enter` en el campo de búsqueda.

4.  **Visualización de Resultados**:
    * **Tabla "Vista Previa Diccionario"**: Muestra el contenido del archivo de diccionario cargado. Si tu búsqueda utiliza el diccionario, las filas FCD que coincidan con tu consulta (o partes de ella) pueden aparecer resaltadas en azul.
    * **Tabla "Resultados / Descripciones"**: Muestra los resultados finales de tu búsqueda obtenidos del archivo de descripciones. Si no se ha realizado una búsqueda o se ha limpiado, puede mostrar una vista previa del archivo de descripciones completo.

5.  **Funciones Adicionales**:
    * **Ordenar Tablas**: Haz clic en la cabecera de cualquier columna en las tablas para ordenar los datos por esa columna (alternando entre ascendente y descendente).
    * **Exportar**: Si hay resultados en la tabla "Resultados / Descripciones", puedes hacer clic en "Exportar" para guardarlos en un nuevo archivo Excel (`.xlsx`) o CSV (`.csv`).
    * **Salvar Regla**: Esta función (actualmente) guarda metadatos sobre la última búsqueda realizada (término, origen, número de filas) en la memoria de la aplicación. No guarda los datos de los resultados en sí.
    * **Ayuda (`?`)**: El botón con un signo de interrogación abre una ventana con información detallada sobre la sintaxis de búsqueda y el flujo de trabajo de la aplicación.
    * **Barra de Estado**: En la parte inferior de la ventana, muestra mensajes sobre el estado actual de la aplicación (ej., archivos cargados, búsqueda en progreso, errores).

## Logging

La aplicación genera un archivo de log llamado `Buscador_Avanzado_App_v1.10.3_Mod.log` en el mismo directorio desde donde se ejecuta `main.py`. Este archivo contiene:
* Mensajes informativos sobre el inicio y fin de la aplicación.
* Detalles sobre la carga de archivos.
* Información sobre el proceso de búsqueda, incluyendo las consultas parseadas y los resultados intermedios (a nivel DEBUG).
* Advertencias sobre situaciones no críticas (ej., no se encontraron mapeos de unidades).
* Errores críticos y tracebacks completos en caso de fallos.

Revisa este archivo si encuentras comportamientos inesperados o para entender mejor el flujo interno de la aplicación.
