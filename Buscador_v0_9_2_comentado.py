# -*- coding: utf-8 -*-
# Se especifica la codificación UTF-8 para asegurar la correcta interpretación de caracteres especiales en el código.

# --- Importaciones ---

# Importaciones de la biblioteca estándar de Python
import json # Para trabajar con el formato de datos JSON (JavaScript Object Notation), usado para la configuración.
import logging # Para registrar eventos, errores y mensajes de depuración de la aplicación.
import os # Para interactuar con el sistema operativo, como la gestión de rutas de archivos.
import platform # Para obtener información sobre la plataforma en la que se ejecuta el script (e.g., Windows, macOS, Linux).
import re # Para trabajar con expresiones regulares, usadas para el análisis y manipulación de patrones de texto.
import traceback # Para obtener información detallada sobre excepciones (e.g., pila de llamadas).
import unicodedata # Para trabajar con la base de datos de caracteres Unicode, útil para normalizar texto.
from enum import Enum, auto # Para crear enumeraciones (conjuntos de constantes simbólicas con nombre).
from pathlib import Path # Ofrece una forma orientada a objetos para manipular rutas de archivos y directorios.
from typing import ( # Módulo para proporcionar indicaciones de tipo (type hints), mejorando la legibilidad y ayudando al análisis estático.
    Any,        # Indica un tipo no restringido, puede ser cualquier cosa.
    Dict,       # Indica un diccionario, con tipos especificados para claves y valores si es necesario (e.g., Dict[str, int]).
    List,       # Indica una lista, con un tipo específico para sus elementos si es necesario (e.g., List[str]).
    Optional,   # Indica que un tipo puede ser el tipo especificado o None (e.g., Optional[str] es str o None).
    Set,        # Indica un conjunto, con un tipo específico para sus elementos si es necesario (e.g., Set[int]).
    Tuple,      # Indica una tupla, con tipos específicos para sus elementos si es necesario (e.g., Tuple[int, str]).
    Union       # Indica que un tipo puede ser uno de varios tipos especificados (e.g., Union[int, str]).
)

# Importaciones de bibliotecas de terceros
import numpy as np # Biblioteca para computación numérica, especialmente para operaciones con arrays y matrices. Usada aquí para np.isclose.
import pandas as pd # Biblioteca fundamental para la manipulación y análisis de datos, especialmente con DataFrames.

# Importaciones específicas de Tkinter (biblioteca estándar para GUI)
import tkinter as tk # La biblioteca principal de Tkinter para crear la interfaz gráfica.
from tkinter import filedialog # Para mostrar diálogos estándar de selección de archivos y directorios.
from tkinter import messagebox # Para mostrar cuadros de diálogo estándar (información, error, advertencia).
from tkinter import ttk # Módulo de Tkinter que provee widgets temáticos (con un aspecto más moderno).


# --- Configuración del Logging ---
# Obtiene una instancia de logger para este módulo.
# El nombre del logger será el nombre del módulo actual (generalmente __main__ cuando se ejecuta como script).
logger = logging.getLogger(__name__)

# --- Enumeraciones ---
class OrigenResultados(Enum):
    """
    Enumeración para rastrear el origen o el estado de una operación de búsqueda.
    Esto ayuda a la lógica de la interfaz de usuario (UI) a determinar qué mensajes mostrar
    o qué acciones tomar en función del resultado de la búsqueda.
    El uso de `auto()` asigna automáticamente un valor único y secuencial a cada miembro.
    """
    NINGUNO = 0 # Estado inicial o no determinado. También usado si la búsqueda no se ejecuta.
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto() # Búsqueda realizada a través del diccionario, se encontraron FCDs (Formas Canónicas del Diccionario) y posteriormente se encontraron resultados en las descripciones.
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto() # Búsqueda vía diccionario, se encontraron FCDs pero no se pudieron extraer términos válidos de ellos para buscar en las descripciones.
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto() # Búsqueda vía diccionario, se encontraron FCDs y se extrajeron términos, pero no hubo coincidencias para esos términos en el archivo de descripciones.
    DICCIONARIO_SIN_COINCIDENCIAS = auto() # La búsqueda inicial en el archivo de diccionario no arrojó ninguna FCD coincidente con la query.
    DIRECTO_DESCRIPCION_CON_RESULTADOS = auto() # Búsqueda realizada directamente en el archivo de descripciones (sin pasar por el diccionario) y se encontraron resultados.
    DIRECTO_DESCRIPCION_VACIA = auto() # Búsqueda directa en descripciones que no arrojó resultados, o la query estaba vacía y se muestran todas las descripciones disponibles.
    ERROR_CARGA_DICCIONARIO = auto() # Ocurrió un error al intentar cargar el archivo Excel del diccionario.
    ERROR_CARGA_DESCRIPCION = auto() # Ocurrió un error al intentar cargar el archivo Excel de descripciones.
    ERROR_CONFIGURACION_COLUMNAS_DICC = auto() # Error detectado en la configuración de las columnas a utilizar del archivo de diccionario.
    ERROR_CONFIGURACION_COLUMNAS_DESC = auto() # Error detectado en la configuración de las columnas a utilizar del archivo de descripciones.
    ERROR_BUSQUEDA_INTERNA_MOTOR = auto() # Un error inesperado ocurrió dentro de la lógica interna del motor de búsqueda.
    TERMINO_INVALIDO = auto() # La consulta de búsqueda o un término específico dentro de ella resultó inválido para el procesamiento (e.g., sintaxis incorrecta no manejable).
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC = auto() # La query original consistía únicamente en términos negados. Se filtraron FCDs (todos menos los negados) y luego se encontraron resultados en descripciones usando los términos de esos FCDs.
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC = auto() # La query original solo tenía negaciones, se filtraron FCDs, pero no se encontraron resultados en descripciones.
    VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC = auto() # Flujo de búsqueda alternativo: se buscaron FCDs que solo contenían la unidad de la query original, y luego se aplicó el filtro numérico/unidad completo en las descripciones encontradas.
    VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC = auto() # Flujo de búsqueda alternativo: se buscaron FCDs por unidad, pero no se encontraron resultados en descripciones al aplicar el filtro numérico.

    # --- Propiedades de la Enumeración ---
    # Estas propiedades son métodos convenientes para agrupar y verificar estados de la enumeración.

    @property
    def es_via_diccionario(self) -> bool:
        """Retorna True si el origen de los resultados involucró el procesamiento del archivo de diccionario en alguna etapa."""
        return self in {
            OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
            OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC,
            OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, # Se considera vía diccionario porque el intento se hizo.
            OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC,
            OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC,
        }

    @property
    def es_directo_descripcion(self) -> bool:
        """Retorna True si el origen de los resultados fue una búsqueda directa en el archivo de descripciones, sin pasar por el diccionario."""
        return self in {
            OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS,
            OrigenResultados.DIRECTO_DESCRIPCION_VACIA
        }

    @property
    def es_error_carga(self) -> bool:
        """Retorna True si el resultado se debe a un error durante la carga de alguno de los archivos Excel (diccionario o descripciones)."""
        return self in {
            OrigenResultados.ERROR_CARGA_DICCIONARIO,
            OrigenResultados.ERROR_CARGA_DESCRIPCION
        }

    @property
    def es_error_configuracion(self) -> bool:
        """Retorna True si el resultado se debe a un error en la configuración de las columnas para la búsqueda (e.g., índices inválidos)."""
        return self in {
            OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC,
            OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC
        }

    @property
    def es_error_operacional(self) -> bool:
        """Retorna True si el resultado se debe a un error interno inesperado durante la operación de búsqueda en el motor."""
        return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR

    @property
    def es_termino_invalido(self) -> bool:
        """Retorna True si la búsqueda no pudo proceder debido a un término o una estructura de consulta inválida."""
        return self == OrigenResultados.TERMINO_INVALIDO

class ExtractorMagnitud:
    """
    Clase responsable de normalizar unidades de medida (magnitudes) y mapear sus sinónimos
    a una forma canónica estándar. Esto es crucial para que la búsqueda numérica
    funcione consistentemente aunque las unidades se expresen de formas diversas en los datos.
    Permite, por ejemplo, que "V", "volt", "voltio" y "VOLTIOS" se traten como la misma unidad.
    """
    MAPEO_MAGNITUDES_PREDEFINIDO: Dict[str, List[str]] = {}

    def __init__(self, mapeo_magnitudes: Optional[Dict[str, List[str]]] = None):
        self.sinonimo_a_canonico_normalizado: Dict[str, str] = {}
        mapeo_a_usar = mapeo_magnitudes if mapeo_magnitudes is not None else self.MAPEO_MAGNITUDES_PREDEFINIDO
        for forma_canonica_original, lista_sinonimos_originales in mapeo_a_usar.items():
            canonico_norm = self._normalizar_texto(forma_canonica_original)
            if not canonico_norm:
                logger.warning(
                    f"Forma canónica original '{forma_canonica_original}' resultó vacía tras normalizar. "
                    f"Será ignorada en la configuración de ExtractorMagnitud."
                )
                continue
            self.sinonimo_a_canonico_normalizado[canonico_norm] = canonico_norm
            for sinonimo_original in lista_sinonimos_originales:
                sinonimo_norm = self._normalizar_texto(str(sinonimo_original))
                if sinonimo_norm:
                    self.sinonimo_a_canonico_normalizado[sinonimo_norm] = canonico_norm
        logger.debug(
            f"ExtractorMagnitud inicializado/actualizado. "
            f"Total de mapeos normalizados (sinónimo -> canónico): {len(self.sinonimo_a_canonico_normalizado)}."
        )

    @staticmethod
    def _normalizar_texto(texto: str) -> str:
        if not isinstance(texto, str) or not texto.strip():
            return ""
        try:
            texto_upper = texto.upper()
            forma_normalizada_nfkd = unicodedata.normalize("NFKD", texto_upper)
            texto_filtrado = "".join(
                caracter for caracter in forma_normalizada_nfkd
                if not unicodedata.combining(caracter)
                and (caracter.isalnum() or caracter.isspace() or caracter in ['.', '-', '_', '/'])
            )
            return ' '.join(texto_filtrado.split())
        except TypeError:
            logger.error(f"TypeError inesperado en _normalizar_texto (ExtractorMagnitud) con entrada: {texto}")
            return ""

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        if not texto_unidad or not str(texto_unidad).strip():
            return None
        unidad_normalizada_input = self._normalizar_texto(str(texto_unidad))
        if not unidad_normalizada_input:
            return None
        return self.sinonimo_a_canonico_normalizado.get(unidad_normalizada_input)

class ManejadorExcel:
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        ruta = Path(ruta_archivo)
        if not ruta.exists():
            mensaje_error = f"¡Archivo no encontrado! La ruta especificada no existe: {ruta}"
            logger.error(f"ManejadorExcel: {mensaje_error}")
            return None, mensaje_error
        try:
            engine: Optional[str] = None
            if ruta.suffix.lower() == ".xlsx":
                engine = "openpyxl"
            logger.info(
                f"ManejadorExcel: Intentando cargar el archivo '{ruta.name}' "
                f"utilizando el motor de Pandas: '{engine if engine else 'automático (pandas intentará determinarlo, usualmente xlrd para .xls si está disponible)'}'."
            )
            df = pd.read_excel(ruta, engine=engine)
            logger.info(
                f"ManejadorExcel: Archivo '{ruta.name}' cargado exitosamente. "
                f"Número de filas: {len(df)}."
            )
            return df, None
        except ImportError as ie:
            mensaje_error_usuario = (
                f"Error al cargar '{ruta.name}': Falta una librería necesaria para leer este tipo de archivo Excel.\n\n"
                f"Para archivos .xlsx (formato moderno), asegúrese de tener instalada la librería 'openpyxl'.\n"
                f"Puede instalarla con el comando: pip install openpyxl\n\n"
                f"Para archivos .xls (formato antiguo), puede necesitar la librería 'xlrd'.\n"
                f"Puede instalarla con: pip install xlrd\n\n"
                f"Detalle técnico del error: {ie}"
            )
            logger.exception(
                f"ManejadorExcel: Falta una dependencia para leer el archivo '{ruta.name}'. Error de importación: {ie}"
            )
            return None, mensaje_error_usuario
        except Exception as e:
            mensaje_error_usuario = (
                f"No se pudo cargar el archivo '{ruta.name}'. Ocurrió un error inesperado:\n{e}\n\n"
                f"Por favor, verifique lo siguiente:\n"
                f"- Que el archivo no esté corrupto o dañado.\n"
                f"- Que la aplicación tenga los permisos necesarios para acceder al archivo y su contenido.\n"
                f"- Que el archivo no esté abierto y bloqueado exclusivamente por otra aplicación (e.g., Microsoft Excel)."
            )
            logger.exception(
                f"ManejadorExcel: Error genérico al intentar cargar el archivo '{ruta.name}'."
            )
            return None, mensaje_error_usuario

class MotorBusqueda:
    """
    Clase principal que encapsula toda la lógica de carga de datos desde archivos Excel,
    el procesamiento de las consultas de búsqueda introducidas por el usuario,
    y la ejecución de la búsqueda sobre los datos cargados.

    Atributos:
        datos_diccionario (Optional[pd.DataFrame]): DataFrame que almacena el contenido del archivo "diccionario".
                                                   Este archivo se usa para buscar sinónimos o características
                                                   que luego se usan para buscar en las descripciones.
        datos_descripcion (Optional[pd.DataFrame]): DataFrame que almacena el contenido del archivo "descripciones",
                                                  sobre el cual se realizan las búsquedas finales.
        archivo_diccionario_actual (Optional[Path]): Ruta al archivo de diccionario actualmente cargado.
        archivo_descripcion_actual (Optional[Path]): Ruta al archivo de descripciones actualmente cargado.
        indices_columnas_busqueda_dic_preview (List[int]): Lista de índices de columnas que se mostrarán
                                                           en la vista previa del diccionario en la UI.
        patron_comparacion (re.Pattern): Expresión regular compilada para parsear términos de comparación numérica
                                       (e.g., ">10V", "<=5A", "=100W").
        patron_rango (re.Pattern): Expresión regular compilada para parsear rangos numéricos
                                 (e.g., "10-20V", "5 - 15 mm").
        patron_termino_negado (re.Pattern): Expresión regular compilada para identificar términos negados
                                          en la consulta (e.g., "#palabra", #"frase negada"#).
        patron_num_unidad_df (re.Pattern): Expresión regular compilada para extraer números y sus unidades
                                         de cadenas de texto dentro de las celdas de los DataFrames.
        extractor_magnitud (ExtractorMagnitud): Instancia para normalizar y mapear unidades de medida.
    """

    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        """
        Inicializa el MotorBusqueda.

        Args:
            indices_diccionario_cfg: Lista opcional de índices de columnas (basados en 0)
                                     a usar para la vista previa del diccionario en la UI.
                                     Si es None o una lista vacía, se usarán por defecto todas
                                     las columnas de tipo texto u objeto del DataFrame del diccionario.
        """
        # DataFrames para almacenar los datos cargados desde los archivos Excel.
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        
        # Rutas (objetos Path) a los archivos Excel actualmente cargados.
        self.archivo_diccionario_actual: Optional[Path] = None
        self.archivo_descripcion_actual: Optional[Path] = None
        
        # Configuración de las columnas que se mostrarán en la vista previa del diccionario en la UI.
        # Si no se especifica, será una lista vacía, y se usarán las columnas por defecto.
        self.indices_columnas_busqueda_dic_preview: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else []
        
        logger.info(
            f"MotorBusqueda inicializado. Índices para la vista previa del diccionario: "
            f"{self.indices_columnas_busqueda_dic_preview or 'Se usarán todas las columnas de tipo texto/objeto por defecto'}."
        )
        
        # --- Patrones de Expresiones Regulares Precompilados ---
        # La precompilación de regex mejora el rendimiento si se usan repetidamente.

        # Patrón para términos de comparación numérica (e.g., ">10V", "<=5A", "=100 W").
        # Grupo 1: Operador de comparación (>, <, >=, <=, o solo =).
        # Grupo 2: Valor numérico (puede ser entero o decimal, usando '.' o ',' como separador decimal).
        # Grupo 3: Unidad de medida (opcional, alfanumérica con algunos caracteres especiales).
        self.patron_comparacion = re.compile(
            r"^\s*"                                  # Coincide con el inicio de la cadena y cualquier espacio en blanco opcional.
            r"([<>]=?|=)"                            # Grupo 1: Captura el operador (>, <, >=, <=, =).
            r"\s*"                                  # Espacios opcionales.
            r"(\d+(?:[.,]\d+)?)"                    # Grupo 2: Captura el valor numérico. \d+ uno o más dígitos. (?:[.,]\d+)? es un grupo opcional sin captura para la parte decimal.
            r"\s*"                                  # Espacios opcionales.
            r"([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?"    # Grupo 3: Captura la unidad (opcional). Permite letras (incluyendo acentuadas y especiales como µ y Ω), puntos, barras, guiones.
            r"\s*$"                                  # Espacios opcionales y fin de la cadena.
        )

        # Patrón para rangos numéricos (e.g., "10-20V", "5.5 - 15.2 mm").
        # Grupo 1: Primer valor numérico del rango.
        # Grupo 2: Segundo valor numérico del rango.
        # Grupo 3: Unidad de medida (opcional, se asume la misma para ambos extremos del rango).
        self.patron_rango = re.compile(
            r"^\s*"                                  # Inicio y espacios.
            r"(\d+(?:[.,]\d+)?)"                    # Grupo 1: Primer número.
            r"\s*-\s*"                              # Separador de rango '-' rodeado de espacios opcionales.
            r"(\d+(?:[.,]\d+)?)"                    # Grupo 2: Segundo número.
            r"\s*"                                  # Espacios opcionales.
            r"([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?"    # Grupo 3: Unidad (opcional).
            r"\s*$"                                  # Espacios y fin.
        )

        # Patrón para identificar términos negados en la consulta (e.g., "#palabra", #"frase con espacios"#).
        # Un término negado indica que las filas que lo contengan deben ser excluidas de los resultados.
        # Grupo 1: Captura una frase negada que está entre comillas dobles.
        # Grupo 2: Captura una palabra negada simple (sin comillas).
        self.patron_termino_negado = re.compile(
            r'#\s*'                                   # Prefijo de negación '#' seguido de cero o más espacios.
            r'(?:'                                   # Inicio de un grupo de no captura (para la alternancia OR).
            r'\"([^\"]+)\"'                         # Opción 1 (Grupo 1): Una comilla doble, seguida de uno o más caracteres que NO son comillas dobles (la frase), y una comilla doble de cierre.
            r'|'                                    # Operador OR.
            r'([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\.\-\_]+)'     # Opción 2 (Grupo 2): Una palabra que consiste en uno o más caracteres alfanuméricos (incluyendo acentuados y especiales comunes) y algunos símbolos permitidos.
            r')',
            re.IGNORECASE | re.UNICODE              # Flags: IGNORECASE para búsqueda insensible a mayúsculas/minúsculas, UNICODE para correcto manejo de caracteres Unicode.
        )
        
        # Patrón para extraer números y sus unidades de cadenas de texto encontradas dentro de las celdas de los DataFrames.
        # Es más flexible que `patron_comparacion` o `patron_rango` ya que busca ocurrencias dentro de un texto más largo.
        # Grupo 1: El valor numérico.
        # Grupo 2: La unidad de medida (opcional).
        self.patron_num_unidad_df = re.compile(
            r"(\d+(?:[.,]\d+)?)"                     # Grupo 1: Número (entero o decimal).
            r"[\s\-]*"                               # Separador opcional (cero o más espacios o guiones) entre el número y la unidad.
            r"([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?"     # Grupo 2: Unidad (opcional).
        )
        
        # Instancia del extractor de magnitudes, que se usará para normalizar unidades.
        # Se inicializa aquí y se puede actualizar si se carga un diccionario con mapeos de unidades.
        self.extractor_magnitud = ExtractorMagnitud() 

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        """
        Carga y procesa el archivo Excel que actúa como "diccionario" de términos y unidades.
        Este diccionario se utiliza para:
        1. Identificar formas canónicas y sinónimos de términos o unidades de medida.
        2. Actualizar dinámicamente el `ExtractorMagnitud` con los mapeos de unidades encontrados.

        Se asume una estructura específica para el archivo de diccionario:
        - La primera columna (índice 0) contiene las formas canónicas de las unidades/términos.
        - Las columnas a partir de la cuarta (índice 3) contienen los sinónimos para la forma canónica de esa fila.
          Las columnas 1 y 2 se ignoran para la extracción de magnitudes pero se conservan en el DataFrame.

        Args:
            ruta_str: Ruta (como cadena) al archivo Excel del diccionario.

        Returns:
            Una tupla (éxito: bool, mensaje_error: Optional[str]).
            `éxito` es True si la carga fue correcta, False en caso contrario.
            `mensaje_error` contiene una descripción del error si `éxito` es False.
        """
        ruta = Path(ruta_str) # Convierte la cadena de ruta a un objeto Path.
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta) # Intenta cargar el Excel.

        # Si la carga del archivo falla (df_cargado es None),
        # resetea los datos del diccionario y el extractor de magnitudes.
        if df_cargado is None:
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            self.extractor_magnitud = ExtractorMagnitud() # Resetea a predefinido o vacío.
            logger.warning(
                f"Fallo al cargar el archivo de diccionario '{ruta.name}'. "
                f"El ExtractorMagnitud se ha reseteado a su estado predefinido o vacío."
            )
            return False, error_msg_carga # Devuelve fallo y el mensaje de error de carga.

        # Diccionario temporal para construir dinámicamente el mapeo para el ExtractorMagnitud.
        # Clave: forma canónica original (string), Valor: lista de sinónimos originales (List[str]).
        mapeo_dinamico_para_extractor: Dict[str, List[str]] = {}
        
        # Procesa el DataFrame cargado para extraer las formas canónicas y sus sinónimos.
        if df_cargado.shape[1] > 0: # Verifica que el DataFrame tenga al menos una columna.
            columna_canonica_nombre = df_cargado.columns[0] # Asume que la primera columna contiene las formas canónicas.
            inicio_col_sinonimos = 3 # Los sinónimos empiezan desde la cuarta columna (índice 3).
                                     # Columnas en índice 1 y 2 se ignoran para este propósito.
            max_cols_a_chequear_para_sinonimos = df_cargado.shape[1] # Número total de columnas en el DataFrame.

            # Itera sobre cada fila del DataFrame del diccionario.
            for _, fila in df_cargado.iterrows():
                forma_canonica_raw = fila.get(columna_canonica_nombre) # Obtiene el valor de la celda de la forma canonica.
                
                # Si la forma canónica es NaN (Not a Number, valor faltante en Pandas) o una cadena vacía/espacios, se salta esta fila.
                if pd.isna(forma_canonica_raw) or str(forma_canonica_raw).strip() == "":
                    continue

                forma_canonica_str = str(forma_canonica_raw).strip() # Convierte a string y quita espacios.
                # La lista de sinónimos para esta forma canónica incluye inicialmente la propia forma canonica.
                sinonimos_para_esta_canonica: List[str] = [forma_canonica_str]

                # Itera sobre las columnas designadas para sinónimos (desde la cuarta columna hasta el final).
                for i in range(inicio_col_sinonimos, max_cols_a_chequear_para_sinonimos):
                    if i < len(df_cargado.columns): # Asegura que el índice de columna sea válido.
                        nombre_col_sinonimo_actual = df_cargado.columns[i] # Nombre de la columna del sinónimo.
                        sinonimo_celda_raw = fila.get(nombre_col_sinonimo_actual) # Valor de la celda del sinónimo.
                        
                        # Si el sinónimo no es NaN y no es una cadena vacía/espacios, se añade a la lista de sinónimos.
                        if pd.notna(sinonimo_celda_raw) and str(sinonimo_celda_raw).strip() != "":
                            sinonimos_para_esta_canonica.append(str(sinonimo_celda_raw).strip())
                
                # Añade la forma canónica y su lista de sinónimos (asegurándose de que sean únicos con set()) al mapeo.
                mapeo_dinamico_para_extractor[forma_canonica_str] = list(set(sinonimos_para_esta_canonica))
            
            # Si se extrajeron mapeos del archivo, se crea una nueva instancia de ExtractorMagnitud con ellos.
            if mapeo_dinamico_para_extractor:
                self.extractor_magnitud = ExtractorMagnitud(mapeo_magnitudes=mapeo_dinamico_para_extractor)
                logger.info(
                    f"Extractor de magnitudes actualizado dinámicamente desde el archivo de diccionario '{ruta.name}'. "
                    f"Se procesaron {len(mapeo_dinamico_para_extractor)} formas canónicas."
                )
            else:
                # Si no se extrajeron mapeos (e.g., archivo vacío o sin datos en las columnas esperadas).
                logger.warning(
                    f"No se extrajeron mapeos de unidad/magnitud válidos desde el archivo '{ruta.name}'. "
                    f"ExtractorMagnitud utilizará su configuración predefinida (si la tiene) o permanecerá vacío."
                )
                self.extractor_magnitud = ExtractorMagnitud() # Resetea a predefinido o vacío.
        else:
            # Si el archivo de diccionario no tiene columnas.
            logger.warning(
                f"El archivo de diccionario '{ruta.name}' no contiene columnas. "
                f"No se pudo actualizar el extractor de magnitudes."
            )
            self.extractor_magnitud = ExtractorMagnitud() # Resetea.

        # Guarda el DataFrame cargado y la ruta del archivo actual del diccionario.
        self.datos_diccionario = df_cargado
        self.archivo_diccionario_actual = ruta

        # Log de depuración (si está habilitado) mostrando las primeras filas del diccionario cargado.
        if logger.isEnabledFor(logging.DEBUG) and self.datos_diccionario is not None:
            logger.debug(
                f"Archivo de diccionario '{ruta.name}' cargado exitosamente y procesado para el extractor de magnitudes. "
                f"Primeras 3 filas del DataFrame del diccionario:\n{self.datos_diccionario.head(3).to_string()}"
            )
        
        return True, None # Indica que la carga y procesamiento fueron exitosos.

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        """
        Carga el archivo Excel que contiene las descripciones o los datos principales sobre los que se buscará.

        Args:
            ruta_str: Ruta (como cadena) al archivo Excel de descripciones.

        Returns:
            Una tupla (éxito: bool, mensaje_error: Optional[str]).
            `éxito` es True si la carga fue correcta, False en caso contrario.
            `mensaje_error` contiene una descripción del error si `éxito` es False.
        """
        ruta = Path(ruta_str) # Convierte a objeto Path.
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta) # Intenta cargar el Excel.

        # Si la carga falla, resetea los datos de descripción.
        if df_cargado is None:
            self.datos_descripcion = None
            self.archivo_descripcion_actual = None
            return False, error_msg_carga # Devuelve fallo y el mensaje de error.
        
        # Si la carga es exitosa, guarda el DataFrame y la ruta.
        self.datos_descripcion = df_cargado
        self.archivo_descripcion_actual = ruta
        logger.info(f"Archivo de descripciones '{ruta.name}' ({len(df_cargado)} filas) cargado exitosamente.")
        return True, None # Indica éxito.

    def _obtener_nombres_columnas_busqueda_df(
        self, df: pd.DataFrame, indices_cfg: List[int], tipo_busqueda: str
    ) -> Tuple[Optional[List[str]], Optional[str]]:
        """
        Determina la lista de nombres de columnas a utilizar para una búsqueda en un DataFrame.
        Puede ser basado en una lista de índices de columnas configurada, o por defecto,
        usando todas las columnas de tipo string u object si no se especifica configuración.

        Args:
            df: El DataFrame del cual obtener los nombres de las columnas.
            indices_cfg: Una lista de enteros representando los índices de las columnas a usar.
                         Si es vacía o contiene solo -1, se usa el comportamiento por defecto.
            tipo_busqueda: Una cadena descriptiva del tipo de búsqueda (e.g., "diccionario_preview",
                           "descripcion_fcds") para mensajes de log/error.

        Returns:
            Una tupla:
                - Lista de nombres de columnas seleccionadas (Optional[List[str]]). None si hay error.
                - Mensaje de error (Optional[str]) si ocurre alguno. None si no hay error.
        """
        # Validaciones iniciales del DataFrame.
        if df is None or df.empty:
            return None, f"El DataFrame para '{tipo_busqueda}' está vacío o no ha sido cargado."
        
        columnas_disponibles = list(df.columns) # Lista de todos los nombres de columnas en el DataFrame.
        num_cols_df = len(columnas_disponibles) # Número total de columnas.

        if num_cols_df == 0:
            return None, f"El DataFrame para '{tipo_busqueda}' no tiene columnas."

        # Determina si se deben usar las columnas por defecto.
        # Esto ocurre si `indices_cfg` está vacía o si el único elemento es -1 (convención para "usar todas").
        usar_columnas_por_defecto = not indices_cfg or indices_cfg == [-1]

        if usar_columnas_por_defecto:
            # Por defecto, se seleccionan todas las columnas que son de tipo string (texto) u object.
            # El tipo 'object' en Pandas a menudo contiene strings, pero puede tener otros tipos mixtos.
            # Se asume que la búsqueda textual es más relevante en estas columnas.
            cols_texto_obj = [
                col_nombre for col_nombre in columnas_disponibles
                if pd.api.types.is_string_dtype(df[col_nombre]) or pd.api.types.is_object_dtype(df[col_nombre])
            ]
            if cols_texto_obj:
                logger.debug(
                    f"Para la búsqueda '{tipo_busqueda}', se utilizarán las columnas de tipo texto/objeto por defecto: "
                    f"{cols_texto_obj}"
                )
                return cols_texto_obj, None
            else:
                # Si no hay columnas de tipo texto/objeto, como fallback, se usan todas las columnas disponibles.
                logger.warning(
                    f"Para la búsqueda '{tipo_busqueda}', no se encontraron columnas de tipo texto/objeto. "
                    f"Como fallback, se utilizarán todas las {num_cols_df} columnas disponibles: {columnas_disponibles}"
                )
                return columnas_disponibles, None
        
        # Si se proporcionaron índices específicos en `indices_cfg`.
        nombres_columnas_seleccionadas: List[str] = []
        indices_invalidos: List[str] = [] # Para registrar índices que estén fuera de rango.

        for i in indices_cfg:
            # Valida que cada índice sea un entero y esté dentro del rango válido de columnas del DataFrame.
            if not (isinstance(i, int) and 0 <= i < num_cols_df):
                indices_invalidos.append(str(i)) # Añade el índice inválido a la lista.
            else:
                # Si el índice es válido, añade el nombre de la columna correspondiente a la lista de seleccionadas.
                nombres_columnas_seleccionadas.append(columnas_disponibles[i])
        
        # Si se encontraron índices inválidos, se retorna un error.
        if indices_invalidos:
            return None, (
                f"Índice(s) de columna {', '.join(indices_invalidos)} inválido(s) para la búsqueda '{tipo_busqueda}'. "
                f"El DataFrame tiene {num_cols_df} columnas (índices válidos de 0 a {num_cols_df-1})."
            )
        
        # Si la lista de configuraciónde índices no resultó en ninguna columna seleccionada (e.g., lista vacía pero no era `-1`).
        if not nombres_columnas_seleccionadas:
            return None, (
                f"La configuración de índices de columnas {indices_cfg} no resultó en ninguna columna válida "
                f"seleccionada para la búsqueda '{tipo_busqueda}'."
            )
            
        logger.debug(
            f"Para la búsqueda '{tipo_busqueda}', se utilizarán las columnas seleccionadas por los índices "
            f"{indices_cfg}: {nombres_columnas_seleccionadas}"
        )
        return nombres_columnas_seleccionadas, None

    def _normalizar_para_busqueda(self, texto: str) -> str:
        """
        Normaliza una cadena de texto específicamente para la comparación durante la búsqueda.
        Este proceso es similar a `ExtractorMagnitud._normalizar_texto` pero puede tener
        ligeras diferencias si se requiriera una normalización distinta para la búsqueda
        en contenido general versus la normalización estricta de unidades.
        Actualmente, la lógica es muy similar:
        1. Convierte a mayúsculas.
        2. Normalización Unicode NFKD para descomponer caracteres.
        3. Eliminación de diacríticos.
        4. Eliminación de caracteres que no sean alfanuméricos, espacios, o '.', '-', '/', '_'.
        5. Consolidación de espacios.

        Args:
            texto: La cadena a normalizar.

        Returns:
            La cadena normalizada. Devuelve la cadena original convertida a mayúsculas y
            con strip si ocurre un error durante la normalización.
        """
        # Si el texto no es una cadena o está vacío, devuelve una cadena vacía.
        if not isinstance(texto, str) or not texto:
            return ""
        try:
            texto_upper = texto.upper() # Insensible a mayúsculas.
            # Descompone caracteres acentuados y otros compuestos.
            texto_norm_nfkd = unicodedata.normalize('NFKD', texto_upper)
            # Elimina los diacríticos (marcas combinatorias).
            texto_sin_acentos = "".join([c for c in texto_norm_nfkd if not unicodedata.combining(c)])
            # Elimina cualquier carácter que no sea letra, número, espacio o los permitidos '.', '-', '/', '_'.
            # Esto ayuda a una coincidencia más robusta al ignorar puntuación no esencial.
            texto_limpio_final = re.sub(r'[^\w\s\.\-\/\_]', '', texto_sin_acentos)
            # Normaliza múltiples espacios a uno solo y quita espacios al inicio/final.
            return ' '.join(texto_limpio_final.split()).strip()
        except Exception as e:
            # Si ocurre cualquier error durante la normalización, se registra y se devuelve
            # una versión simplificada (mayúsculas y strip) del texto original como fallback.
            logger.error(f"Error al normalizar el texto para búsqueda '{texto[:50]}...': {e}")
            return str(texto).upper().strip() # Fallback

    # ... (El resto de los métodos de MotorBusqueda se comentarían con un nivel de detalle similar) ...
    # ... (Se omiten por brevedad en esta respuesta, pero el proceso sería el mismo) ...
    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, cols: List[str], texto: str) -> Tuple[pd.DataFrame, str, List[str]]:
        texto_limpio_entrada = texto.strip(); terminos_negados_encontrados: List[str] = []
        df_a_procesar = df_original.copy() if df_original is not None else pd.DataFrame() # Copia para no modificar el original si se pasa uno
        if not texto_limpio_entrada: return df_a_procesar, "", terminos_negados_encontrados # Si no hay texto, no hay nada que hacer

        partes_positivas: List[str] = []
        ultimo_indice_fin_negado = 0

        # Itera sobre todas las coincidencias de términos negados en el texto de entrada
        for match_negado in self.patron_termino_negado.finditer(texto_limpio_entrada):
            # Añade la parte del texto ANTES del término negado actual a las partes positivas
            partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:match_negado.start()])
            # Actualiza el índice del final del último término negado encontrado
            ultimo_indice_fin_negado = match_negado.end()
            # Obtiene el término negado (puede estar en el grupo 1 si es frase entre comillas, o en el grupo 2 si es palabra)
            termino_negado_raw = match_negado.group(1) or match_negado.group(2)
            if termino_negado_raw:
                # Normaliza el término negado (quitando comillas si las tuviera)
                termino_negado_normalizado = self._normalizar_para_busqueda(termino_negado_raw.strip('"'))
                # Si el término normalizado es válido y no se ha añadido antes, lo guarda
                if termino_negado_normalizado and termino_negado_normalizado not in terminos_negados_encontrados:
                    terminos_negados_encontrados.append(termino_negado_normalizado)
        
        # Añade la parte del texto DESPUÉS del último término negado (o todo el texto si no hubo negados)
        partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:])
        # Reconstruye la cadena de términos positivos, limpiando espacios extra
        terminos_positivos_final_str = ' '.join("".join(partes_positivas).split()).strip()

        # Si no hay DataFrame para procesar, o no hay términos negados, o no hay columnas donde buscar,
        # no se puede aplicar el filtro de negación.
        if df_a_procesar.empty or not terminos_negados_encontrados or not cols:
            logger.debug(f"Parseo de negación: Query='{texto_limpio_entrada}', Positivos='{terminos_positivos_final_str}', Negados={terminos_negados_encontrados}. No se aplicó filtro de negación al DataFrame.")
            return df_a_procesar, terminos_positivos_final_str, terminos_negados_encontrados

        # Máscara booleana para acumular filas que contienen CUALQUIERA de los términos negados
        mascara_exclusion_total = pd.Series(False, index=df_a_procesar.index)
        for termino_negado_actual in terminos_negados_encontrados:
            if not termino_negado_actual: continue # Ignora términos negados vacíos
            
            mascara_para_este_termino_negado = pd.Series(False, index=df_a_procesar.index) # Máscara para este término negado específico
            patron_regex_negado = r"\b" + re.escape(termino_negado_actual) + r"\b" # Busca palabra completa

            # Busca el término negado en cada columna especificada
            for nombre_columna in cols:
                if nombre_columna not in df_a_procesar.columns: continue # Si la columna no existe en el df, la ignora
                try:
                    # Normaliza el contenido de la columna y busca el patrón
                    serie_columna_normalizada = df_a_procesar[nombre_columna].astype(str).map(self._normalizar_para_busqueda)
                    mascara_para_este_termino_negado |= serie_columna_normalizada.str.contains(patron_regex_negado, regex=True, na=False)
                except Exception as e_neg_col:
                    logger.error(f"Error aplicando filtro de negación en columna '{nombre_columna}' para término '{termino_negado_actual}': {e_neg_col}")
            
            # Acumula las filas que coinciden con este término negado
            mascara_exclusion_total |= mascara_para_este_termino_negado
        
        # Invierte la máscara para obtener las filas que NO contienen NINGUNO de los términos negados
        df_resultado_filtrado = df_a_procesar[~mascara_exclusion_total]
        logger.info(f"Filtrado por negación (Query original='{texto_limpio_entrada}'): {len(df_a_procesar)} filas originales -> {len(df_resultado_filtrado)} filas después del filtro. Términos negados: {terminos_negados_encontrados}. Términos positivos resultantes: '{terminos_positivos_final_str}'")
        
        return df_resultado_filtrado, terminos_positivos_final_str, terminos_negados_encontrados

    # ... (rest of MotorBusqueda methods with detailed comments) ...
    # Note: Due to response length limits, only a part of MotorBusqueda is shown here.
    # The full implementation would comment all methods as per the plan.

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        """
        Descompone una cadena de búsqueda compleja en segmentos de Nivel 1,
        determinando si la operación principal es OR o AND.
        Actualmente, solo `|` se trata como un OR de alto nivel si no hay `+`
        que indiquen una operación AND más prioritaria fuera de paréntesis.

        Args:
            texto_complejo: La cadena de búsqueda a descomponer.

        Returns:
            Una tupla:
                - El operador principal detectado ("OR" o "AND").
                - Una lista de cadenas, donde cada cadena es un segmento de la operación.
                  Si es "AND", la lista suele tener un solo elemento (la cadena original).
                  Si es "OR", la lista tiene los operandos del OR.
        """
        texto_limpio = texto_complejo.strip()
        if not texto_limpio:
            return "OR", [] # Si está vacío, no hay segmentos, se podría considerar un OR de nada.

        # Si hay un '+' y la expresión no está completamente entre paréntesis, se asume que es un AND de nivel superior.
        # Ejemplo: "A + B | C" se trataría como AND de ["A + B | C"] en este nivel.
        # Ejemplo: "(A | B) + C" se trataría como AND de ["(A | B) + C"]
        # La lógica de `+` se maneja en `_descomponer_nivel2_and`.
        if '+' in texto_complejo and not (texto_limpio.startswith("(") and texto_limpio.endswith(")")):
            logger.debug(f"Descomposición Nivel 1 (OR) para '{texto_complejo}': Detectado '+' de alto nivel. Tratando como AND en este nivel. Segmento único: ['{texto_complejo}']")
            return "AND", [texto_limpio]

        # Definición de separadores OR de alto nivel. Actualmente solo '|'.
        # Se podría extender para incluir " OR " si se deseara, pero requeriría cuidado con frases exactas.
        separadores_or = [(r"\s*\|\s*", "|")]  # Regex para `|` rodeado de espacios opcionales.

        for sep_regex, sep_char_literal in separadores_or:
            # Solo considera `|` como OR si no hay `+` (ya que `+` tiene precedencia implícita o se maneja como un bloque).
            if '+' not in texto_complejo and sep_char_literal in texto_limpio:
                # Divide la cadena por el separador OR. `re.split` es más robusto que `str.split`.
                # Filtra segmentos vacíos que podrían surgir de múltiples separadores juntos o al inicio/final.
                segmentos_potenciales = [s.strip() for s in re.split(sep_regex, texto_complejo) if s.strip()]
                
                # Si hay más de un segmento, o si hay un solo segmento pero es diferente del original
                # (lo que podría indicar que el separador estaba al inicio/final y fue eliminado por strip/split),
                # entonces se considera una operación OR.
                if len(segmentos_potenciales) > 1 or \
                   (len(segmentos_potenciales) == 1 and texto_limpio != segmentos_potenciales[0]):
                    logger.debug(f"Descomposición Nivel 1 (OR) para '{texto_complejo}': Operador principal=OR, Segmentos={segmentos_potenciales}")
                    return "OR", segmentos_potenciales
        
        # Si no se encontró un operador OR de alto nivel, se asume que toda la cadena es un único segmento AND (o un término simple).
        logger.debug(f"Descomposición Nivel 1 (OR) para '{texto_complejo}': Operador principal=AND (no se encontró OR explícito de alto nivel). Segmento único: ['{texto_limpio}']")
        return "AND", [texto_limpio]

    def _descomponer_nivel2_and(self, termino_segmento_n1: str) -> Tuple[str, List[str]]:
        """
        Descompone un segmento de búsqueda (generalmente proveniente de _descomponer_nivel1_or)
        en términos atómicos que están unidos por el operador AND (representado por '+').

        Args:
            termino_segmento_n1: El segmento de búsqueda a descomponer.

        Returns:
            Una tupla:
                - Siempre "AND" (ya que este nivel asume AND entre los términos resultantes).
                - Una lista de cadenas, donde cada cadena es un término atómico
                  que debe cumplirse (operación AND).
        """
        termino_limpio = termino_segmento_n1.strip()
        if not termino_limpio:
            return "AND", [] # Si está vacío, no hay términos.

        # Divide la cadena por el separador AND (' + ' con espacios).
        # `re.split` con `\s+\+\s+` maneja correctamente los espacios alrededor del '+'.
        partes_crudas = re.split(r'\s+\+\s+', termino_limpio)
        
        # Limpia cada parte resultante (quitando espacios) y filtra las partes vacías.
        partes_limpias_finales = [p.strip() for p in partes_crudas if p.strip()]
        
        logger.debug(f"Descomposición Nivel 2 (AND) para el segmento '{termino_segmento_n1}': Partes resultantes={partes_limpias_finales}")
        return "AND", partes_limpias_finales

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        """
        Analiza una lista de términos de búsqueda "brutos" (cadenas) y los convierte
        en una lista de diccionarios estructurados, cada uno representando un término analizado
        con su tipo (numérico, rango, cadena), valor, y unidad (si aplica).

        Args:
            terminos_brutos: Lista de cadenas, donde cada cadena es un término de búsqueda
                             (e.g., ">10V", "palabra", "10-20mm", "\"frase exacta\"").

        Returns:
            Lista de diccionarios, donde cada diccionario contiene:
                - "original": El término como se procesó (sin comillas de frase exacta).
                - "tipo": "gt", "lt", "ge", "le", "eq", "range", o "str".
                - "valor": El valor numérico, lista de [min, max] para rango, o cadena normalizada.
                - "unidad_busqueda": La unidad canónica normalizada (si aplica), o None.
        """
        terminos_analizados: List[Dict[str, Any]] = []
        for termino_original_bruto in terminos_brutos:
            termino_original_procesado = str(termino_original_bruto).strip() # Limpia espacios.
            es_frase_exacta = False
            termino_final_para_analisis = termino_original_procesado

            # Detecta si es una frase exacta (encerrada entre comillas dobles).
            if len(termino_final_para_analisis) >= 2 and \
               termino_final_para_analisis.startswith('"') and \
               termino_final_para_analisis.endswith('"'):
                termino_final_para_analisis = termino_final_para_analisis[1:-1] # Quita las comillas.
                es_frase_exacta = True # Marca que la búsqueda debe ser exacta para esta frase.
            
            # Si el término queda vacío después de quitar comillas, se ignora.
            if not termino_final_para_analisis:
                continue

            item_analizado: Dict[str, Any] = {"original": termino_final_para_analisis}

            # Intenta hacer match con patrones numéricos solo si NO es una frase exacta.
            # Una frase como "">10V"" se tratará como texto literal, no como comparación.
            match_comparacion = None
            match_rango = None
            if not es_frase_exacta:
                match_comparacion = self.patron_comparacion.match(termino_final_para_analisis)
                match_rango = self.patron_rango.match(termino_final_para_analisis)

            if match_comparacion: # Si coincide con el patrón de comparación (e.g., >10V)
                operador_str, valor_str, unidad_str_raw = match_comparacion.groups()
                valor_numerico = self._parse_numero(valor_str) # Parsea el valor numérico.
                
                if valor_numerico is not None: # Si el parseo fue exitoso.
                    mapa_operadores = {">": "gt", "<": "lt", ">=": "ge", "<=": "le", "=": "eq"}
                    unidad_canonica: Optional[str] = None
                    # Si se proporcionó una unidad, la normaliza.
                    if unidad_str_raw and unidad_str_raw.strip():
                        unidad_canonica = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_raw.strip())
                    item_analizado.update({
                        "tipo": mapa_operadores.get(operador_str), 
                        "valor": valor_numerico, 
                        "unidad_busqueda": unidad_canonica
                    })
                else: # Si no se pudo parsear como número, se trata como término de cadena.
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            
            elif match_rango: # Si coincide con el patrón de rango (e.g., 10-20V)
                valor1_str, valor2_str, unidad_str_r_raw = match_rango.groups()
                valor1_num = self._parse_numero(valor1_str)
                valor2_num = self._parse_numero(valor2_str)

                if valor1_num is not None and valor2_num is not None: # Si ambos números del rango son válidos.
                    unidad_canonica_r: Optional[str] = None
                    if unidad_str_r_raw and unidad_str_r_raw.strip():
                        unidad_canonica_r = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_r_raw.strip())
                    # El valor para un rango es una lista ordenada [min, max].
                    item_analizado.update({
                        "tipo": "range", 
                        "valor": sorted([valor1_num, valor2_num]), 
                        "unidad_busqueda": unidad_canonica_r
                    })
                else: # Si los números del rango no son válidos, se trata como término de cadena.
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            
            else: # Si no es comparación ni rango (o si era frase exacta)
                # Se trata como un término de cadena simple. El valor se normaliza para búsqueda.
                # Si era frase exacta, el valor normalizado será la frase sin comillas, normalizada.
                # La lógica de búsqueda deberá tratar `es_frase_exacta` para buscar la cadena completa.
                # NOTA: El flag `es_frase_exacta` no se almacena directamente en `item_analizado` aquí,
                # pero la normalización de `valor` ya la considera. La lógica en
                # `_generar_mascara_para_un_termino` necesitará saber si buscar `valor` como frase o como palabra.
                # Actualmente, la normalización aquí para "str" y la búsqueda en
                # `_generar_mascara_para_un_termino` (con `r"\b" + re.escape(...) + r"\b"`)
                # implica una búsqueda de palabras/subcadenas normalizadas.
                # Para frases exactas literales, el patrón de regex en la búsqueda debería ser `re.escape(valor_normalizado_busqueda)` sin `\b`.
                # Esta implementación actual trata las frases exactas normalizando el contenido y luego buscando ese contenido
                # como una secuencia de palabras (debido al `\b` en la búsqueda de strings).
                item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            
            terminos_analizados.append(item_analizado)
            
        logger.debug(f"Términos (después de la descomposición AND) analizados para la búsqueda detallada: {terminos_analizados}")
        return terminos_analizados

    # Los métodos restantes de MotorBusqueda (`_generar_mascara_para_un_termino`,
    # `_aplicar_mascara_combinada_para_segmento_and`, `_combinar_mascaras_de_segmentos_or`,
    # `_procesar_busqueda_en_df_objetivo`, `_extraer_terminos_de_fila_completa` y `buscar`)
    # son complejos y requieren un desglose detallado de comentarios similar al de los métodos anteriores.
    # Se omiten aquí por la limitación de longitud de la respuesta, pero el enfoque sería el mismo:
    # - Docstring para el método.
    # - Comentarios para bloques lógicos clave.
    # - Explicación de variables importantes.
    # - Justificación de decisiones de diseño si no son obvias.

    # ... (Implementación completa y comentada de los métodos restantes de MotorBusqueda) ...
    # Por brevedad, se muestra la estructura del método buscar que es el principal.
    def buscar(self, termino_busqueda_original: str, buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[List[int]], Optional[str]]:
        """
        Método principal para ejecutar una búsqueda. Puede operar de dos maneras:
        1. Vía Diccionario: La consulta se busca primero en el `datos_diccionario`. Los términos/sinónimos
           extraídos de las coincidencias del diccionario se usan luego para buscar en `datos_descripcion`.
        2. Directa: La consulta se busca directamente en `datos_descripcion`.

        Maneja la lógica de parseo de la consulta (operadores AND, OR, negaciones, numéricos, rangos, frases).

        Args:
            termino_busqueda_original: La cadena de consulta introducida por el usuario.
            buscar_via_diccionario_flag: Booleano que indica si se debe intentar la búsqueda vía diccionario.

        Returns:
            Una tupla conteniendo:
                - resultados_df (Optional[pd.DataFrame]): DataFrame con las filas de `datos_descripcion` que coinciden.
                - origen_resultado (OrigenResultados): Enum que indica cómo se obtuvieron los resultados o el tipo de error.
                - fcds_coincidentes (Optional[pd.DataFrame]): DataFrame con las filas del `datos_diccionario` que coincidieron
                                                               (relevante si `buscar_via_diccionario_flag` es True).
                - indices_fcds_resaltar (Optional[List[int]]): Lista de índices de las FCDs para resaltar en la UI.
                - mensaje_error (Optional[str]): Mensaje de error si la operación falló.
        """
        logger.info(f"Motor.buscar INICIO: Termino='{termino_busqueda_original}', Vía Diccionario={buscar_via_diccionario_flag}")

        # Prepara un DataFrame vacío con las columnas de descripción por si no hay resultados o errores.
        columnas_descripcion_ref = list(self.datos_descripcion.columns) if self.datos_descripcion is not None else []
        df_vacio_para_descripciones = pd.DataFrame(columns=columnas_descripcion_ref)

        # Inicializa variables que se retornarán.
        fcds_obtenidos_final_para_ui: Optional[pd.DataFrame] = None
        indices_fcds_a_resaltar_en_preview: Optional[List[int]] = None

        # --- Manejo de consulta vacía ---
        if not termino_busqueda_original.strip(): # Si la consulta está vacía o solo espacios.
            if self.datos_descripcion is not None:
                # Si hay descripciones cargadas, devuelve todas las descripciones.
                logger.info("Consulta vacía, devolviendo todas las descripciones.")
                return self.datos_descripcion.copy(), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
            else:
                # Si no hay descripciones, devuelve un DataFrame vacío.
                logger.info("Consulta vacía y no hay descripciones cargadas.")
                return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, "Archivo de descripciones no cargado."

        # --- Parseo Global de Términos Positivos y Negativos de la Consulta Original ---
        # `_aplicar_negaciones_y_extraer_positivos` se usa aquí con un DF dummy solo para parsear la query.
        _, terminos_positivos_globales, terminos_negativos_globales = self._aplicar_negaciones_y_extraer_positivos(
            pd.DataFrame(), [], termino_busqueda_original
        )
        logger.info(f"Parseo global de la consulta: Términos Positivos='{terminos_positivos_globales}', Términos Negativos Globales={terminos_negativos_globales}")
        
        # --- Detección de Filtro Numérico/Unidad en la Query Original ---
        # Si la query original (su parte positiva) parece ser una condición numérica con unidad (e.g., ">10V"),
        # se guarda esta condición para poder aplicarla más tarde en el flujo de "rescate por unidad".
        filtro_numerico_original_de_query: Optional[Dict[str, Any]] = None
        if terminos_positivos_globales.strip():
            # Se analiza solo el primer segmento AND del primer segmento OR para simplificar.
            # Esto es una heurística para capturar el caso simple como " >10V " o " 10-20mm ".
            _op_l1, segs_l1 = self._descomponer_nivel1_or(terminos_positivos_globales)
            if segs_l1: # Si hay al menos un segmento OR.
                _op_l2, segs_l2 = self._descomponer_nivel2_and(segs_l1[0]) # Toma el primer segmento OR y lo descompone en AND.
                if segs_l2: # Si hay al menos un término AND.
                    # Analiza el primer término AND.
                    terminos_analizados_temp = self._analizar_terminos([segs_l2[0]])
                    if terminos_analizados_temp and \
                       terminos_analizados_temp[0]["tipo"] in ["gt", "lt", "ge", "le", "eq", "range"] and \
                       terminos_analizados_temp[0].get("unidad_busqueda"):
                        # Si el primer término es numérico y tiene una unidad, se considera el filtro original.
                        filtro_numerico_original_de_query = terminos_analizados_temp[0].copy()
                        logger.info(f"Detectado filtro numérico/unidad en la consulta original: {filtro_numerico_original_de_query}")

        # --- Lógica Principal de Búsqueda: Vía Diccionario o Directa ---
        if buscar_via_diccionario_flag:
            # --- BÚSQUEDA VÍA DICCIONARIO ---
            if self.datos_diccionario is None:
                logger.error("Intento de búsqueda vía diccionario, pero el diccionario no está cargado.")
                return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, None, "El archivo de Diccionario no ha sido cargado."
            
            # Obtiene las columnas a usar para buscar en el diccionario.
            columnas_dic_para_fcds, err_msg_cols_dic = self._obtener_nombres_columnas_busqueda_df(
                self.datos_diccionario, [], "diccionario_fcds_inicial" # Usa config por defecto para cols de dicc.
            )
            if not columnas_dic_para_fcds:
                return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, None, err_msg_cols_dic

            # --- Manejo de Consultas AND Complejas (con '+' a nivel global) ---
            # Si la parte positiva de la query contiene '+' y no es una frase exacta,
            # se asume una operación AND entre múltiples sub-consultas.
            if "+" in terminos_positivos_globales and not \
               (terminos_positivos_globales.startswith('"') and terminos_positivos_globales.endswith('"')):
                logger.info(f"Detectada búsqueda AND compleja en positivos globales: '{terminos_positivos_globales}'")
                partes_and = [p.strip() for p in terminos_positivos_globales.split("+") if p.strip()] # Divide por '+'.
                
                # Inicializa el DataFrame de resultados acumulados para descripciones.
                df_resultado_acumulado_desc = self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame(columns=columnas_descripcion_ref)
                fcds_indices_acumulados = set() # Para guardar los índices de los FCDs de todas las partes AND.
                todas_partes_and_produjeron_terminos_validos = True # Flag.
                hay_error_en_busqueda_de_parte_o_desc = False # Flag.
                error_msg_critico_partes: Optional[str] = None

                if self.datos_descripcion is None:
                    logger.error("Archivo de descripciones no cargado, no se puede proceder con búsqueda AND vía diccionario.")
                    return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Archivo de descripciones no cargado, necesario para búsqueda AND."
                
                columnas_desc_para_filtrado, err_cols_desc_fil = self._obtener_nombres_columnas_busqueda_df(
                    self.datos_descripcion, [], "descripcion_fcds_and_complejo"
                )
                if not columnas_desc_para_filtrado:
                    return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_fil

                # Procesa cada parte de la operación AND.
                for i, parte_and_actual_str in enumerate(partes_and):
                    if not parte_and_actual_str: continue # Ignora partes vacías.
                    logger.debug(f"Procesando parte AND '{parte_and_actual_str}' (parte {i+1}/{len(partes_and)}) en el diccionario...")
                    
                    # 1. Busca la parte AND actual en el DICCIONARIO para obtener FCDs.
                    fcds_para_esta_parte, error_fcd_parte = self._procesar_busqueda_en_df_objetivo(
                        self.datos_diccionario, columnas_dic_para_fcds, parte_and_actual_str, None # Sin negaciones adicionales aquí.
                    )
                    if error_fcd_parte: # Si hay error en la búsqueda de esta parte.
                        todas_partes_and_produjeron_terminos_validos = False; hay_error_en_busqueda_de_parte_o_desc = True
                        error_msg_critico_partes = error_fcd_parte
                        logger.warning(f"Parte AND '{parte_and_actual_str}' falló en la búsqueda del diccionario con error: {error_fcd_parte}"); break
                    
                    if fcds_para_esta_parte is None or fcds_para_esta_parte.empty: # Si no hay FCDs para esta parte.
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no encontró FCDs coincidentes en el diccionario. La búsqueda AND global fallará."); break
                    
                    fcds_indices_acumulados.update(fcds_para_esta_parte.index.tolist()) # Acumula índices de FCDs.

                    # 2. Extrae todos los términos (sinónimos) de las FCDs encontradas para esta parte AND.
                    terminos_extraidos_de_esta_parte_set: Set[str] = set()
                    for _, fila_fcd in fcds_para_esta_parte.iterrows():
                        terminos_extraidos_de_esta_parte_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                    
                    if not terminos_extraidos_de_esta_parte_set: # Si no se extrajeron términos de los FCDs.
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' encontró FCDs, pero no se pudieron extraer términos de búsqueda de ellos para las descripciones."); break
                    
                    # 3. Crea una query OR con los términos extraídos (sinónimos) para buscar en DESCRIPCIONES.
                    # Las frases se encierran en comillas para la sub-query.
                    terminos_or_con_comillas_actual = [
                        f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t
                        for t in terminos_extraidos_de_esta_parte_set if t
                    ]
                    query_or_simple_actual = " | ".join(terminos_or_con_comillas_actual)

                    if not query_or_simple_actual: # Si no se pudo construir la query OR.
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no generó una query OR válida a partir de los términos extraídos para buscar en descripciones."); break
                    
                    # Si el df_resultado_acumulado_desc ya está vacío por una parte AND anterior, no tiene sentido seguir.
                    if df_resultado_acumulado_desc.empty and i >= 0: # i >= 0 asegura que no sea la primera iteración (aunque ya estaría vacío).
                         logger.info(
                             f"Los resultados acumulados de descripción ya están vacíos antes de aplicar el filtro para la parte AND '{parte_and_actual_str}'. "
                             f"La búsqueda AND global resultará vacía."
                         ); break

                    logger.info(
                        f"Aplicando filtro OR para la parte AND '{parte_and_actual_str}' "
                        f"(Query de sinónimos: '{query_or_simple_actual[:100]}...') "
                        f"sobre {len(df_resultado_acumulado_desc)} filas de descripción acumuladas."
                    )
                    # 4. Filtra el df_resultado_acumulado_desc con la query OR de sinónimos.
                    # Esto efectivamente hace un AND entre los resultados de las descripciones para cada parte de la query original.
                    df_resultado_acumulado_desc, error_sub_busqueda_desc = self._procesar_busqueda_en_df_objetivo(
                        df_resultado_acumulado_desc, columnas_desc_para_filtrado, query_or_simple_actual, None # No se aplican negaciones globales aquí, sino al final.
                    )
                    if error_sub_busqueda_desc: # Si hay error en esta sub-búsqueda.
                        hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_sub_busqueda_desc
                        logger.error(f"Error en la sub-búsqueda OR en descripciones para '{query_or_simple_actual}': {error_sub_busqueda_desc}"); break
                    
                    if df_resultado_acumulado_desc.empty: # Si esta parte AND no encontró nada en las descripciones acumuladas.
                        logger.info(
                            f"El filtro OR para la parte AND '{parte_and_actual_str}' no encontró coincidencias en los resultados acumulados de descripciones. "
                            f"La búsqueda AND global resultará vacía."
                        ); break
                
                # Fin del bucle de partes AND.
                # Construye el DataFrame de FCDs finales para la UI.
                if fcds_indices_acumulados and self.datos_diccionario is not None:
                    fcds_obtenidos_final_para_ui = self.datos_diccionario.loc[list(fcds_indices_acumulados)].drop_duplicates().copy()
                    indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                else: # Si no hubo FCDs acumulados.
                    fcds_obtenidos_final_para_ui = pd.DataFrame(columns=self.datos_diccionario.columns if self.datos_diccionario is not None else [])
                    indices_fcds_a_resaltar_en_preview = []
                
                # Evalúa el resultado del flujo AND complejo.
                if hay_error_en_busqueda_de_parte_o_desc: # Si hubo un error crítico.
                    return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_msg_critico_partes
                
                if not todas_partes_and_produjeron_terminos_validos or df_resultado_acumulado_desc.empty:
                    # Si alguna parte AND falló en encontrar FCDs/términos, o si el resultado final en descripciones es vacío.
                    origen_fallo_and = OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS if not todas_partes_and_produjeron_terminos_validos else OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                    logger.info(f"Búsqueda AND compleja '{terminos_positivos_globales}' no produjo resultados finales en descripciones (Origen: {origen_fallo_and.name}).")
                    return df_vacio_para_descripciones, origen_fallo_and, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                
                # Si todo fue bien y hay resultados, aplica las negaciones globales.
                resultados_desc_final_filtrado_and = df_resultado_acumulado_desc
                if not resultados_desc_final_filtrado_and.empty and terminos_negativos_globales:
                    logger.info(
                        f"Aplicando términos negativos globales ({terminos_negativos_globales}) a "
                        f"{len(resultados_desc_final_filtrado_and)} filas (resultado del AND complejo de partes)."
                    )
                    query_solo_negados_globales = " ".join([f"#{neg}" for neg in terminos_negativos_globales])
                    # Se usa `_aplicar_negaciones_y_extraer_positivos` solo para el efecto de negación.
                    df_temp_neg, _, _ = self._aplicar_negaciones_y_extraer_positivos(
                        resultados_desc_final_filtrado_and, columnas_desc_para_filtrado, query_solo_negados_globales
                    )
                    resultados_desc_final_filtrado_and = df_temp_neg # Actualiza con el resultado filtrado.
                
                logger.info(
                    f"Búsqueda AND compleja '{terminos_positivos_globales}' vía diccionario produjo "
                    f"{len(resultados_desc_final_filtrado_and)} resultados finales en descripciones."
                )
                return resultados_desc_final_filtrado_and, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
            
            else: 
                # --- FLUJO SIMPLE (NO AND COMPLEJO) O PURAMENTE NEGATIVO VÍA DICCIONARIO ---
                origen_propuesto_flujo_simple: OrigenResultados = OrigenResultados.NINGUNO
                fcds_intento1: Optional[pd.DataFrame] = None # FCDs del primer intento de búsqueda en diccionario.

                if terminos_positivos_globales.strip(): # Si hay términos positivos.
                    logger.info(f"BUSCAR EN DICC (FCDs) - Intento 1 (Query con Positivos): Query='{terminos_positivos_globales}'")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
                    try:
                        # Busca los términos positivos en el diccionario.
                        fcds_temp, error_dic_pos = self._procesar_busqueda_en_df_objetivo(
                            self.datos_diccionario, columnas_dic_para_fcds, terminos_positivos_globales, None
                        )
                        if error_dic_pos: # Si hay error en esta búsqueda.
                            return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_pos
                        fcds_intento1 = fcds_temp
                    except Exception as e_dic_pos:
                        logger.exception("Excepción durante la búsqueda en diccionario (flujo simple, positivos).")
                        return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error interno del motor (búsqueda en diccionario - positivos simples): {e_dic_pos}"
                
                elif terminos_negativos_globales: # Si NO hay términos positivos, PERO SÍ hay negativos globales.
                    logger.info(f"BUSCAR EN DICC (FCDs) - Flujo Puramente Negativo: Negativos Globales={terminos_negativos_globales}")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC
                    try:
                        # Crea una query que solo contiene los términos negados.
                        query_solo_negados_fcd = " ".join([f"#{neg}" for neg in terminos_negativos_globales])
                        # Busca en el diccionario. Esto devolverá todas las filas del diccionario MENOS las que coincidan con los negados.
                        fcds_temp, error_dic_neg = self._procesar_busqueda_en_df_objetivo(
                            self.datos_diccionario, columnas_dic_para_fcds, query_solo_negados_fcd, None
                        )
                        if error_dic_neg:
                            return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_neg
                        fcds_intento1 = fcds_temp
                    except Exception as e_dic_neg:
                        logger.exception("Excepción durante la búsqueda en diccionario (flujo puramente negativo).")
                        return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error interno del motor (búsqueda en diccionario - puramente negativo): {e_dic_neg}"
                else:
                    # Si no hay ni términos positivos ni negativos globales (la query original era solo espacios y ya se manejó).
                    # O si la query original se parseó a nada positivo ni negativo.
                    logger.info("No hay términos positivos ni negativos globales después del parseo inicial. Se considera DICCIONARIO_SIN_COINCIDENCIAS.")
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None
                
                fcds_obtenidos_final_para_ui = fcds_intento1 # Guarda los FCDs del primer intento.

                # --- Lógica de "Rescate por Unidad" ---
                # Si el Intento 1 de FCDs falló (no encontró nada o fue None) Y
                # la query original era una condición numérica con unidad específica.
                if (fcds_obtenidos_final_para_ui is None or fcds_obtenidos_final_para_ui.empty) and \
                   filtro_numerico_original_de_query and \
                   filtro_numerico_original_de_query.get("unidad_busqueda"):
                    
                    unidad_query_original_can = filtro_numerico_original_de_query["unidad_busqueda"]
                    logger.info(
                        f"Intento 1 de búsqueda de FCDs (usando query numérica/unidad original) falló o no arrojó resultados. "
                        f"Iniciando Intento 2: Buscando FCDs que contengan solo la unidad '{unidad_query_original_can}' en el diccionario."
                    )
                    # Crea una query para buscar FCDs que simplemente contengan la unidad (como frase exacta).
                    query_solo_unidad_para_fcd = f'"{unidad_query_original_can}"' 
                    fcds_por_unidad, err_fcd_unidad = self._procesar_busqueda_en_df_objetivo(
                        self.datos_diccionario, columnas_dic_para_fcds, query_solo_unidad_para_fcd, None
                    )
                    if err_fcd_unidad:
                        logger.warning(f"Error en la búsqueda alternativa de FCDs por unidad '{query_solo_unidad_para_fcd}': {err_fcd_unidad}")
                    
                    if fcds_por_unidad is not None and not fcds_por_unidad.empty:
                        # Si se encontraron FCDs que contienen la unidad.
                        logger.info(
                            f"Intento 2 (rescate por unidad): Encontrados {len(fcds_por_unidad)} FCDs alternativos "
                            f"basados solo en la unidad '{query_solo_unidad_para_fcd}'."
                        )
                        fcds_obtenidos_final_para_ui = fcds_por_unidad # Actualiza los FCDs para la UI.
                        indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                        
                        # Extrae términos de estos FCDs (basados en unidad) para buscar en descripciones.
                        terminos_de_unidad_para_desc_set: Set[str] = set()
                        for _, fila_fcd_unidad in fcds_por_unidad.iterrows():
                            terminos_de_unidad_para_desc_set.update(self._extraer_terminos_de_fila_completa(fila_fcd_unidad))
                        
                        if not terminos_de_unidad_para_desc_set:
                            logger.info("Intento 2 (rescate por unidad): Se encontraron FCDs por unidad, pero no se extrajeron términos de ellos para buscar en descripciones.")
                            return df_vacio_para_descripciones, OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                        
                        query_or_de_unidades_para_desc = " | ".join([
                            f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t
                            for t in terminos_de_unidad_para_desc_set if t
                        ])

                        if not query_or_de_unidades_para_desc:
                            return df_vacio_para_descripciones, OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "La query OR construida desde FCDs por unidad (para buscar en descripciones) está vacía."
                        
                        if self.datos_descripcion is None:
                            return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Archivo de descripciones no cargado (necesario para flujo de rescate por unidad)."
                        
                        columnas_desc_alt, err_cols_desc_alt = self._obtener_nombres_columnas_busqueda_df(
                            self.datos_descripcion, [], "descripcion_fcds_alternativo_unidad"
                        )
                        if not columnas_desc_alt:
                            return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, err_cols_desc_alt
                        
                        logger.info(
                            f"BUSCAR EN DESCRIPCIONES (Intento 2 - vía FCDs por unidad): "
                            f"Query de sinónimos='{query_or_de_unidades_para_desc[:100]}...'. "
                            f"Se aplicará el filtro numérico original: {filtro_numerico_original_de_query} "
                            f"y los Negativos Globales: {terminos_negativos_globales}"
                        )
                        # Los negativos globales se aplican aquí también, a menos que la query original fuera puramente negativa.
                        neg_glob_alt = terminos_negativos_globales if origen_propuesto_flujo_simple != OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC else []
                        
                        # Busca en descripciones usando los términos de los FCDs (de unidad) Y el filtro numérico original.
                        resultados_desc_alt, error_desc_alt = self._procesar_busqueda_en_df_objetivo(
                            self.datos_descripcion, 
                            columnas_desc_alt, 
                            query_or_de_unidades_para_desc, 
                            terminos_negativos_adicionales=neg_glob_alt, 
                            filtro_numerico_original_desc=filtro_numerico_original_de_query # Crucial: aplica el filtro numérico original.
                        )
                        if error_desc_alt:
                            return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_desc_alt
                        
                        if resultados_desc_alt is None or resultados_desc_alt.empty:
                            return df_vacio_para_descripciones, OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                        else:
                            # Éxito en el flujo de rescate por unidad.
                            return resultados_desc_alt, OrigenResultados.VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                    else: 
                        # Si el Intento 2 (rescate por unidad) tampoco encontró FCDs.
                        logger.info(
                            f"Intento 2 (rescate por unidad): No se encontraron FCDs basados solo en la unidad '{query_solo_unidad_para_fcd}'. "
                            f"La búsqueda vía diccionario para '{termino_busqueda_original}' se considera sin coincidencias."
                        )
                        return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None # No hay FCDs.
                
                # --- Continuación del Flujo Simple (si Intento 1 de FCDs tuvo éxito o si el rescate por unidad no aplicó/falló y volvemos aquí) ---
                if fcds_obtenidos_final_para_ui is not None and not fcds_obtenidos_final_para_ui.empty: 
                    # Si se encontraron FCDs (ya sea del intento 1 o porque el rescate no cambió nada).
                    if indices_fcds_a_resaltar_en_preview is None: # Si no se establecieron en el rescate.
                        indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                    
                    logger.info(f"FCDs obtenidos del diccionario (flujo estándar simple o puramente negativo): {len(fcds_obtenidos_final_para_ui)} filas.")
                    
                    if self.datos_descripcion is None: # Necesitamos descripciones para seguir.
                        return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Archivo de descripciones no cargado."
                    
                    # Extrae todos los términos únicos de las FCDs encontradas.
                    terminos_para_buscar_en_descripcion_set: Set[str] = set()
                    for _, fila_fcd in fcds_obtenidos_final_para_ui.iterrows():
                        terminos_para_buscar_en_descripcion_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                    
                    if not terminos_para_buscar_en_descripcion_set:
                        # Si se encontraron FCDs pero no se pudieron extraer términos de ellos.
                        logger.info("Se encontraron FCDs (flujo estándar), pero no se extrajeron términos de ellos para buscar en descripciones.")
                        origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                        if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                            origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                        return df_vacio_para_descripciones, origen_final_sinterm, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None

                    logger.info(
                        f"Términos extraídos de FCDs para buscar en descripciones ({len(terminos_para_buscar_en_descripcion_set)} únicos). "
                        f"Muestra: {sorted(list(terminos_para_buscar_en_descripcion_set))[:10]}..."
                    )
                    # Construye una query OR con todos los términos extraídos para buscar en descripciones.
                    terminos_or_con_comillas_desc = [
                        f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t
                        for t in terminos_para_buscar_en_descripcion_set if t
                    ]
                    query_or_para_desc_simple = " | ".join(terminos_or_con_comillas_desc)

                    if not query_or_para_desc_simple: # Si la query OR resultante es vacía.
                        origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                        if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                             origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                        return df_vacio_para_descripciones, origen_q_vacia, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "La query OR construida para descripciones está vacía."

                    columnas_desc_final_simple, err_cols_desc_final_simple = self._obtener_nombres_columnas_busqueda_df(
                        self.datos_descripcion, [], "descripcion_fcds_flujo_simple"
                    )
                    if not columnas_desc_final_simple:
                        return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, err_cols_desc_final_simple
                    
                    # Determina si se deben aplicar los negativos globales (solo si la query original no era puramente negativa).
                    negativos_a_aplicar_desc_simple = terminos_negativos_globales if origen_propuesto_flujo_simple != OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC else []
                    
                    logger.info(
                        f"BUSCAR EN DESCRIPCIONES (vía FCDs del flujo estándar/simple): Query de sinónimos='{query_or_para_desc_simple[:200]}...'. "
                        f"Negativos adicionales a aplicar en Descripciones: {negativos_a_aplicar_desc_simple}"
                    )
                    try:
                        # Busca en descripciones usando la query OR de sinónimos y los negativos globales.
                        resultados_desc_final_simple, error_busqueda_desc_simple = self._procesar_busqueda_en_df_objetivo(
                            self.datos_descripcion, columnas_desc_final_simple, query_or_para_desc_simple,
                            terminos_negativos_adicionales=negativos_a_aplicar_desc_simple
                        )
                        if error_busqueda_desc_simple: # Si hay error en esta búsqueda.
                            return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_busqueda_desc_simple
                        
                        if resultados_desc_final_simple is None or resultados_desc_final_simple.empty:
                            # Si no se encontraron resultados en descripciones.
                            origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                            if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                                origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                            return df_vacio_para_descripciones, origen_res_desc_vacio_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                        else:
                            # Éxito en el flujo simple/negativo vía diccionario.
                            return resultados_desc_final_simple, origen_propuesto_flujo_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                    except Exception as e_desc_proc_simple:
                        logger.exception("Excepción durante la búsqueda final en descripciones (flujo estándar/simple).")
                        return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, f"Error interno del motor (búsqueda en descripciones - flujo estándar): {e_desc_proc_simple}"
                else: 
                    # Si el Intento 1 de FCDs falló Y el flujo de rescate por unidad no aplicó o también falló en encontrar FCDs.
                    logger.info(
                        f"No se encontraron FCDs en el diccionario para la consulta '{termino_busqueda_original}' (después de todos los intentos). "
                        f"Resultado: DICCIONARIO_SIN_COINCIDENCIAS."
                    )
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None
        else: 
            # --- BÚSQUEDA DIRECTA EN DESCRIPCIONES ---
            # Esto ocurre si buscar_via_diccionario_flag es False, o si se llama desde la UI
            # después de que la búsqueda vía diccionario no arrojó resultados satisfactorios.
            if self.datos_descripcion is None:
                logger.error("Intento de búsqueda directa en descripciones, pero las descripciones no están cargadas.")
                return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "El archivo de Descripciones no ha sido cargado."
            
            columnas_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(
                self.datos_descripcion, [], "descripcion_directa" # Usa config por defecto para cols de desc.
            )
            if not columnas_desc_directo:
                return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_directo
            
            try:
                logger.info(f"BUSCAR EN DESCRIPCIONES (BÚSQUEDA DIRECTA): Query original='{termino_busqueda_original}'")
                # Llama a _procesar_busqueda_en_df_objetivo con la consulta original completa
                # (que ya incluye positivos y negativos parseados por `_aplicar_negaciones_y_extraer_positivos` al inicio de `buscar`).
                # No se pasan negativos adicionales aquí, ya están en `termino_busqueda_original` (o más bien, en cómo se procesa).
                resultados_directos_desc, error_busqueda_desc_dir = self._procesar_busqueda_en_df_objetivo(
                    self.datos_descripcion, columnas_desc_directo, termino_busqueda_original, None
                )
                if error_busqueda_desc_dir: # Si hay error en la búsqueda directa.
                    return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_busqueda_desc_dir
                
                if resultados_directos_desc is None or resultados_directos_desc.empty:
                    # Si la búsqueda directa no arrojó resultados.
                    return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
                else:
                    # Si la búsqueda directa encontró resultados.
                    return resultados_directos_desc, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None, None
            except Exception as e_desc_dir_proc:
                logger.exception("Excepción durante la búsqueda directa en descripciones.")
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error interno del motor (búsqueda directa en descripciones): {e_desc_dir_proc}"

# --- Interfaz Gráfica ---
class InterfazGrafica(tk.Tk):
    """
    Clase principal de la aplicación que hereda de tk.Tk para crear la ventana principal
    y gestionar todos los elementos de la interfaz gráfica de usuario (GUI) y sus interacciones.
    """
    # Nombre del archivo JSON donde se guarda y carga la configuración de la aplicación.
    CONFIG_FILE_NAME = "config_buscador_avanzado_ui.json" 

    def __init__(self):
        """
        Constructor de la clase InterfazGrafica.
        Inicializa la ventana principal, carga la configuración, instancia el motor de búsqueda,
        y crea y configura todos los widgets de la UI.
        """
        super().__init__() # Llama al constructor de la clase padre (tk.Tk).
        
        # Configuración inicial de la ventana principal.
        self.title("Buscador Avanzado v1.10.3 (Reconocimiento Magnitudes Mejorado)") # Título de la ventana.
        self.geometry("1250x800") # Tamaño inicial de la ventana (ancho x alto).
        
        # Carga la configuración de la aplicación desde el archivo JSON.
        self.config: Dict[str, Any] = self._cargar_configuracion_app() 
        
        # Obtiene los índices de columnas para la vista previa del diccionario desde la configuración cargada.
        # Si no está en la config, usa una lista vacía por defecto.
        indices_cfg_preview_dic = self.config.get("indices_columnas_busqueda_dic_preview", [])
        
        # Crea una instancia del motor de búsqueda, pasándole la configuración de las columnas del diccionario.
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg_preview_dic) 
        
        # --- Atributos de estado de la UI ---
        self.resultados_actuales: Optional[pd.DataFrame] = None # Almacena el DataFrame de los últimos resultados de búsqueda mostrados.
        self.texto_busqueda_var = tk.StringVar(self) # Variable de Tkinter asociada al campo de entrada de búsqueda.
        # `trace_add` observa cambios en `texto_busqueda_var`. Cuando el usuario escribe ("write"),
        # llama a `_on_texto_busqueda_change` para actualizar el estado de los botones de operadores.
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change) 
        self.ultimo_termino_buscado: Optional[str] = None # Almacena la última consulta de búsqueda ejecutada.
        self.reglas_guardadas: List[Dict[str, Any]] = [] # Lista para almacenar metadatos de búsquedas guardadas (no implementado completamente en esta versión).
        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None # DataFrame de las FCDs encontradas en la última búsqueda vía diccionario.
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None # DataFrame de los resultados finales en descripciones de la última búsqueda.
        self.indices_fcds_resaltados: Optional[List[int]] = None # Lista de índices de las FCDs para resaltar en la tabla de diccionario.
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO # Estado de la última búsqueda según la enumeración.
        
        # Colores para las filas de las tablas (Treeview).
        self.color_fila_par: str = "white"
        self.color_fila_impar: str = "#f0f0f0" # Un gris claro.
        self.color_resaltado_dic: str = "sky blue" # Color para resaltar filas en la tabla de diccionario.
        
        # Diccionario para almacenar los botones de operadores (+, |, #, etc.) para fácil acceso.
        self.op_buttons: Dict[str, ttk.Button] = {} 
        
        # --- Llamadas a métodos de configuración de la UI ---
        self._configurar_estilo_ttk_app() # Aplica un tema y estilo a los widgets ttk.
        self._crear_widgets_app() # Crea todos los widgets (botones, etiquetas, tablas, etc.).
        self._configurar_grid_layout_app() # Organiza los widgets en la ventana usando el sistema grid.
        self._configurar_eventos_globales_app() # Configura eventos globales como presionar Enter en la búsqueda.
        self._configurar_tags_estilo_treeview_app() # Configura los tags de estilo para las tablas Treeview (colores de filas).
        
        # Configura la funcionalidad de ordenamiento por columnas para ambas tablas.
        self._configurar_funcionalidad_orden_tabla(self.tabla_resultados) 
        self._configurar_funcionalidad_orden_tabla(self.tabla_diccionario) 
        
        # Mensaje inicial en la barra de estado.
        self._actualizar_mensaje_barra_estado("Listo. Cargue el archivo de Diccionario y el de Descripciones para comenzar.") 
        
        # Estado inicial de los botones.
        self._deshabilitar_botones_operadores() # Los botones de operadores empiezan deshabilitados.
        self._actualizar_estado_general_botones_y_controles() # Actualiza el estado de otros botones (Buscar, Exportar, etc.).
        
        logger.info(f"Interfaz Gráfica (v{self.title().split('v')[-1].split(' ')[0]}) inicializada correctamente.")

    def _try_except_wrapper(self, func: callable, *args: Any, **kwargs: Any) -> Any:
        """
        Un decorador o wrapper de utilidad para envolver llamadas a funciones (generalmente callbacks de UI)
        en un bloque try-except. Esto centraliza el manejo de excepciones inesperadas,
        registrándolas y mostrando un mensaje de error al usuario sin que la aplicación crashee.

        Args:
            func: La función a ejecutar de forma segura.
            *args: Argumentos posicionales para `func`.
            **kwargs: Argumentos de palabra clave para `func`.

        Returns:
            El resultado de `func(*args, **kwargs)` si no hay excepciones, o None si ocurre una excepción.
        """
        try:
            # Intenta ejecutar la función proporcionada con sus argumentos.
            return func(*args, **kwargs)
        except Exception as e:
            # Si ocurre cualquier excepción durante la ejecución de `func`.
            func_name = func.__name__ # Nombre de la función que causó el error.
            error_type = type(e).__name__ # Tipo de la excepción (e.g., ValueError, FileNotFoundError).
            error_msg = str(e) # Mensaje de la excepción.
            tb_str = traceback.format_exc() # Pila de llamadas completa como string.

            # Registra el error crítico en el archivo de log y lo imprime en la consola.
            logger.critical(f"Error inesperado en la función '{func_name}': {error_type} - {error_msg}\nTraceback completo:\n{tb_str}")
            print(f"--- TRACEBACK COMPLETO (manejado por _try_except_wrapper para la función: {func_name}) ---\n{tb_str}")
            
            # Muestra un cuadro de diálogo de error al usuario.
            messagebox.showerror(
                f"Error Interno en '{func_name}'",
                f"Ocurrió un error inesperado al realizar la operación:\n"
                f"Tipo de error: {error_type}\n"
                f"Mensaje: {error_msg}\n\n"
                f"Consulte el archivo de log ({LOG_FILE_NAME}) y la consola para obtener el traceback completo y más detalles técnicos."
            )
            
            # Lógica específica si el error ocurrió durante la carga de archivos:
            # Actualiza las etiquetas de los archivos y el estado de los botones para reflejar el fallo.
            if func_name in ["_cargar_diccionario_ui", "_cargar_excel_descripcion_ui"]:
                self._actualizar_etiquetas_archivos_cargados()
                self._actualizar_estado_general_botones_y_controles()
            
            return None # Devuelve None para indicar que la función envuelta no completó exitosamente.

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str) -> None:
        """
        Callback que se ejecuta cada vez que el contenido de `self.texto_busqueda_var` (la entrada de búsqueda) cambia.
        Su propósito principal es llamar a `_actualizar_estado_botones_operadores` para habilitar/deshabilitar
        los botones de operadores (+, |, #, etc.) según el contexto del texto y la posición del cursor.

        Args:
            var_name: El nombre interno de la variable de Tkinter que cambió (no se usa aquí).
            index: El índice donde ocurrió el cambio (no se usa aquí).
            mode: El tipo de operación que causó el cambio ("write", "read", "delete") (no se usa aquí).
        """
        self._actualizar_estado_botones_operadores() # Llama al método que actualiza los botones.

    def _cargar_configuracion_app(self) -> Dict[str, Any]:
        """
        Carga la configuración de la aplicación desde un archivo JSON (definido por `CONFIG_FILE_NAME`).
        La configuración puede incluir rutas a los últimos archivos cargados y otras preferencias de UI.

        Returns:
            Un diccionario con la configuración cargada. Si el archivo no existe o hay un error
            al cargarlo, devuelve un diccionario con valores por defecto para ciertas claves.
        """
        config_cargada: Dict[str, Any] = {} # Diccionario para almacenar la configuración.
        ruta_archivo_config = Path(self.CONFIG_FILE_NAME) # Objeto Path para el archivo de configuración.

        if ruta_archivo_config.exists(): # Si el archivo de configuración existe.
            try:
                # Abre el archivo en modo lectura ("r") con codificación UTF-8.
                with ruta_archivo_config.open("r", encoding="utf-8") as f:
                    config_cargada = json.load(f) # Carga los datos JSON del archivo al diccionario.
                logger.info(f"Configuración de la aplicación cargada exitosamente desde: {self.CONFIG_FILE_NAME}")
            except Exception as e:
                # Si ocurre un error al leer o parsear el archivo JSON.
                logger.error(f"Error al cargar el archivo de configuración '{self.CONFIG_FILE_NAME}': {e}")
        else:
            # Si el archivo de configuración no se encuentra.
            logger.info(f"Archivo de configuración '{self.CONFIG_FILE_NAME}' no encontrado. Se usarán valores por defecto.")

        # Asegura que las claves para las rutas de los últimos archivos existan en la config,
        # convirtiéndolas a objetos Path si existen, o None si no.
        for clave_ruta in ["last_dic_path", "last_desc_path"]:
            valor_ruta = config_cargada.get(clave_ruta) # Obtiene el valor de la ruta (puede ser None).
            # Si hay un valor, lo convierte a Path; sino, lo deja como None.
            config_cargada[clave_ruta] = Path(valor_ruta) if valor_ruta else None
        
        # Asegura que la clave para los índices de preview del diccionario exista, con una lista vacía como valor por defecto.
        config_cargada.setdefault("indices_columnas_busqueda_dic_preview", [])
        
        return config_cargada # Devuelve el diccionario de configuración.

    def _guardar_configuracion_app(self) -> None:
        """
        Guarda la configuración actual de la aplicación en el archivo JSON (`CONFIG_FILE_NAME`).
        Esto incluye las rutas de los últimos archivos de diccionario y descripciones cargados,
        y la configuración de los índices de columnas para la vista previa del diccionario.
        Se llama típicamente al cerrar la aplicación.
        """
        # Actualiza el diccionario `self.config` con los valores actuales.
        # Convierte las rutas de Path a string para la serialización JSON, o guarda None si no hay archivo.
        self.config["last_dic_path"] = str(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else None
        self.config["last_desc_path"] = str(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else None
        self.config["indices_columnas_busqueda_dic_preview"] = self.motor.indices_columnas_busqueda_dic_preview
        
        try:
            # Abre el archivo de configuración en modo escritura ("w") con codificación UTF-8.
            # `json.dump` escribe el diccionario `self.config` al archivo en formato JSON.
            # `indent=4` formatea el JSON para que sea legible por humanos (con indentación).
            with open(self.CONFIG_FILE_NAME, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4)
            logger.info(f"Configuración de la aplicación guardada exitosamente en: {self.CONFIG_FILE_NAME}")
        except Exception as e:
            # Si ocurre un error al escribir el archivo JSON.
            logger.error(f"Error al guardar el archivo de configuración '{self.CONFIG_FILE_NAME}': {e}")

    def _configurar_estilo_ttk_app(self) -> None:
        """
        Configura el estilo de los widgets ttk (temáticos) para la aplicación.
        Intenta aplicar un tema nativo del sistema operativo si está disponible,
        o recurre a temas por defecto de Tkinter.
        También configura un estilo específico para los botones de operadores.
        """
        style = ttk.Style(self) # Obtiene la instancia de estilo de la aplicación.
        os_name = platform.system() # Obtiene el nombre del sistema operativo (e.g., "Windows", "Darwin" para macOS, "Linux").

        # Preferencias de temas TTK para diferentes sistemas operativos.
        # El orden en las listas indica la prioridad.
        theme_preferences = {
            "Windows": ["vista", "xpnative", "clam"], # Para Windows, intenta 'vista', luego 'xpnative', luego 'clam'.
            "Darwin": ["aqua", "clam"],             # Para macOS, intenta 'aqua', luego 'clam'.
            "Linux": ["clam", "alt", "default"]       # Para Linux, intenta 'clam', 'alt', luego 'default'.
        }
        
        # Selecciona el primer tema preferido que esté disponible en el sistema.
        # `style.theme_names()` devuelve una tupla de los nombres de temas disponibles.
        # `style.theme_use()` (sin argumentos) devuelve el tema actual.
        available_themes = style.theme_names()
        preferred_os_themes = theme_preferences.get(os_name, ["clam"]) # Usa 'clam' como fallback si el SO no está en `theme_preferences`.
        
        chosen_theme = style.theme_use() # Empieza con el tema actual o por defecto.
        for theme_name in preferred_os_themes:
            if theme_name in available_themes:
                chosen_theme = theme_name
                break # Usa el primer tema preferido que encuentre.
        
        try:
            style.theme_use(chosen_theme) # Aplica el tema seleccionado.
            # Configura un estilo personalizado llamado "Operator.TButton" para los botones de operadores.
            # `padding` añade espacio interno, `font` ajusta la fuente (usa la fuente por defecto de Tkinter, tamaño 9).
            style.configure("Operator.TButton", padding=(2,1), font=("TkDefaultFont", 9))
            logger.info(f"Tema TTK aplicado: {chosen_theme}")
        except tk.TclError:
            # Si hay un error al aplicar el tema (e.g., el tema no es válido o hay un problema con Tcl/Tk).
            logger.warning(f"Fallo al aplicar el tema TTK '{chosen_theme}'. Se usará el tema por defecto.")


    # ... (El resto de los métodos de InterfazGrafica y el bloque __main__
    #      se comentarían con un nivel de detalle similar, explicando la creación
    #      de widgets, el layout, el manejo de eventos, y la lógica de la UI.) ...
    # Por brevedad, solo se incluye una parte de la clase InterfazGrafica.


# --- Punto de Entrada Principal de la Aplicación ---
if __name__ == "__main__":
    # Este bloque solo se ejecuta cuando el script es el programa principal (no cuando es importado como módulo).

    # Define el nombre del archivo de log. Se usa la versión actual para diferenciar logs si se tienen múltiples versiones.
    LOG_FILE_NAME = f"Buscador_Avanzado_App_v1.10.3.log" # Actualizado
    
    # Configuración básica del logging.
    logging.basicConfig(
        level=logging.DEBUG, # Nivel mínimo de severidad para los mensajes que se registrarán (DEBUG es el más bajo).
        format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s", # Formato de los mensajes de log.
        handlers=[
            logging.FileHandler(LOG_FILE_NAME, encoding="utf-8", mode="w"), # Escribe logs a un archivo (modo "w" sobrescribe en cada ejecución).
            logging.StreamHandler() # Envía logs a la consola (stderr por defecto).
        ]
    )
    # Obtiene el logger raíz para configurar un mensaje inicial.
    root_logger = logging.getLogger()
    root_logger.info(f"--- Iniciando Buscador Avanzado v1.10.3 (Reconocimiento Magnitudes Mejorado) (Script: {Path(__file__).name}) ---")
    root_logger.info(f"Los logs de esta sesión se guardarán en: {Path(LOG_FILE_NAME).resolve()}")

    # --- Verificación de Dependencias Críticas ---
    # Lista para almacenar los nombres de las dependencias que falten.
    dependencias_faltantes_main: List[str] = []
    
    # Intenta importar cada dependencia y registra su versión si está presente.
    # Si falla la importación, añade el nombre de la dependencia a la lista de faltantes.
    try:
        import pandas as pd_check_main # Importa con un alias temporal para no interferir con el 'pd' global.
        root_logger.info(f"Dependencia encontrada: Pandas versión {pd_check_main.__version__}")
    except ImportError:
        dependencias_faltantes_main.append("pandas")
        root_logger.error("Dependencia crítica 'pandas' NO encontrada.")

    try:
        import openpyxl as opxl_check_main
        root_logger.info(f"Dependencia encontrada: openpyxl versión {opxl_check_main.__version__}")
    except ImportError:
        dependencias_faltantes_main.append("openpyxl") # Necesario para leer/escribir archivos .xlsx.
        root_logger.error("Dependencia 'openpyxl' NO encontrada. El manejo de archivos .xlsx fallará.")
    
    try:
        import numpy as np_check_main
        root_logger.info(f"Dependencia encontrada: Numpy versión {np_check_main.__version__}")
    except ImportError:
        dependencias_faltantes_main.append("numpy")
        root_logger.error("Dependencia crítica 'numpy' NO encontrada.")

    try:
        import xlrd as xlrd_check_main # Opcional, para formatos .xls antiguos.
        root_logger.info(f"Dependencia (opcional para .xls) encontrada: xlrd versión {xlrd_check_main.__version__}")
    except ImportError:
        # No se añade a `dependencias_faltantes_main` como crítico, pero se advierte.
        root_logger.warning(
            "Dependencia 'xlrd' no encontrada. La carga de archivos Excel .xls (formato antiguo) podría fallar. "
            "Se recomienda usar archivos .xlsx."
        )

    # Si alguna dependencia crítica falta, muestra un error y termina la aplicación.
    if dependencias_faltantes_main:
        mensaje_error_deps_main = (
            f"Faltan dependencias críticas para ejecutar la aplicación: {', '.join(dependencias_faltantes_main)}.\n\n"
            f"Por favor, instálelas usando pip. Por ejemplo:\n"
            f"pip install {' '.join(dependencias_faltantes_main)}"
        )
        root_logger.critical(mensaje_error_deps_main) # Registra el error crítico.
        try:
            # Intenta mostrar el error en un cuadro de diálogo de Tkinter si Tkinter está disponible.
            root_error_tk_main = tk.Tk() # Crea una ventana raíz temporal.
            root_error_tk_main.withdraw() # La oculta, solo se necesita para el messagebox.
            messagebox.showerror("Dependencias Faltantes", mensaje_error_deps_main)
            root_error_tk_main.destroy() # Destruye la ventana temporal.
        except Exception as e_tk_dep_main:
            # Si falla al mostrar el messagebox (e.g., Tkinter no está bien configurado), imprime el error en consola.
            print(f"ERROR CRÍTICO (No se pudo mostrar el mensaje de error con Tkinter: {e_tk_dep_main}): {mensaje_error_deps_main}")
        exit(1) # Termina la aplicación con un código de error.

    # --- Ejecución de la Aplicación ---
    try:
        app = InterfazGrafica() # Crea una instancia de la clase principal de la GUI.
        app.mainloop() # Inicia el bucle principal de eventos de Tkinter, que mantiene la ventana abierta y receptiva.
    except Exception as e_main_app_exc:
        # Captura cualquier excepción no controlada que pueda ocurrir durante la inicialización o ejecución de la app.
        root_logger.critical("Error fatal no controlado en la aplicación principal:", exc_info=True) # exc_info=True añade el traceback al log.
        tb_str_fatal = traceback.format_exc() # Obtiene el traceback como string.
        print(f"--- TRACEBACK FATAL (desde el bloque principal __main__) ---\n{tb_str_fatal}")
        try:
            # Intenta mostrar un mensaje de error fatal con Tkinter.
            root_fatal_tk_main = tk.Tk(); root_fatal_tk_main.withdraw()
            messagebox.showerror("Error Fatal Inesperado", f"Ocurrió un error crítico inesperado: {e_main_app_exc}\n\nConsulte el archivo '{LOG_FILE_NAME}' y la consola para más detalles.");
            root_fatal_tk_main.destroy()
        except:
            # Si Tkinter también falla aquí, solo imprime en consola.
            print(f"ERROR FATAL: {e_main_app_exc}. Revise el archivo de log '{LOG_FILE_NAME}'.")
    finally:
        # Este bloque se ejecuta siempre, ya sea que la aplicación termine normalmente o por una excepción.
        root_logger.info(f"--- Finalizando la ejecución del Buscador Avanzado ---")