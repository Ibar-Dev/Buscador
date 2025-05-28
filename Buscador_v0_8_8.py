# -*- coding: utf-8 -*-
# Se especifica la codificación UTF-8 para asegurar la correcta interpretación de caracteres especiales.

# Importaciones de la biblioteca estándar y de terceros
import re # Módulo para trabajar con expresiones regulares, fundamental para el análisis de texto.
import tkinter as tk # Biblioteca para la creación de interfaces gráficas de usuario (GUI).
from tkinter import ttk # Módulo de Tkinter que provee widgets temáticos (mejorados).
from tkinter import messagebox # Para mostrar cuadros de diálogo estándar (información, error, advertencia).
from tkinter import filedialog # Para mostrar diálogos de selección de archivos y directorios.
import pandas as pd # Biblioteca para la manipulación y análisis de datos, especialmente con DataFrames.
from typing import ( # Módulo para proporcionar indicaciones de tipo (type hints), mejorando la legibilidad y ayudando al análisis estático.
    Optional, # Indica que un tipo puede ser el tipo especificado o None.
    List, # Indica una lista de un tipo específico.
    Tuple, # Indica una tupla de tipos específicos.
    Union, # Indica que un tipo puede ser uno de varios tipos especificados.
    Set, # Indica un conjunto de un tipo específico.
    Dict, # Indica un diccionario con tipos específicos para claves y valores.
    Any, # Indica un tipo no restringido.
)
from enum import Enum, auto # Módulo para crear enumeraciones, que son conjuntos de constantes simbólicas.

import platform # Módulo para acceder a datos de identificación de la plataforma subyacente (sistema operativo).
import unicodedata # Módulo para acceder a la Base de Datos de Caracteres Unicode (UCD).
import logging # Módulo para emitir mensajes de registro desde bibliotecas y aplicaciones.
import json # Módulo para trabajar con el formato de datos JSON.
import os # Módulo que proporciona una forma de usar funcionalidades dependientes del sistema operativo.
from pathlib import Path # Módulo que ofrece clases para representar rutas de sistema de archivos con semántica para diferentes SO.
import traceback # Módulo para obtener y formatear tracebacks de excepciones.

import numpy as np # Biblioteca para computación numérica, fundamental para operaciones con arrays.

# --- Configuración del Logging ---
logger = logging.getLogger(__name__)

# --- Enumeraciones ---
class OrigenResultados(Enum):
    NINGUNO = 0 # Estado inicial o sin resultados definidos.
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto() # Búsqueda vía diccionario encontró FCDs y estos produjeron resultados en descripciones.
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto() # Búsqueda vía diccionario encontró FCDs, pero no se extrajeron términos válidos de ellos para buscar en descripciones.
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto() # Búsqueda vía diccionario encontró FCDs y generó términos, pero estos no dieron resultados en descripciones.
    DICCIONARIO_SIN_COINCIDENCIAS = auto() # El término de búsqueda no encontró ninguna FCD en el diccionario.
    DIRECTO_DESCRIPCION_CON_RESULTADOS = auto() # Búsqueda directa en descripciones produjo resultados.
    DIRECTO_DESCRIPCION_VACIA = auto() # Búsqueda directa en descripciones no produjo resultados (o término vacío mostrando todas).
    ERROR_CARGA_DICCIONARIO = auto() # Error al intentar cargar el archivo de diccionario.
    ERROR_CARGA_DESCRIPCION = auto() # Error al intentar cargar el archivo de descripciones.
    ERROR_CONFIGURACION_COLUMNAS_DICC = auto() # Error en la configuración de columnas para el diccionario.
    ERROR_CONFIGURACION_COLUMNAS_DESC = auto() # Error en la configuración de columnas para las descripciones.
    ERROR_BUSQUEDA_INTERNA_MOTOR = auto() # Un error genérico o inesperado dentro del motor de búsqueda.
    TERMINO_INVALIDO = auto() # El término de búsqueda fue parseado como inválido o no generó segmentos de búsqueda válidos.
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC = auto() # La búsqueda era puramente negativa, FCDs filtrados por negación produjeron resultados en desc.
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC = auto() # La búsqueda era puramente negativa, FCDs filtrados por negación no produjeron resultados en desc.

    @property
    def es_via_diccionario(self) -> bool:
        # Propiedad para verificar si el origen del resultado fue a través del flujo del diccionario.
        return self in {
            OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
            OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC,
            OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS,
            OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC,
        }
    @property
    def es_directo_descripcion(self) -> bool:
        # Propiedad para verificar si el origen del resultado fue una búsqueda directa en descripciones.
        return self in {OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, OrigenResultados.DIRECTO_DESCRIPCION_VACIA}
    @property
    def es_error_carga(self) -> bool:
        # Propiedad para verificar si el origen fue un error de carga de archivo.
        return self in {OrigenResultados.ERROR_CARGA_DICCIONARIO, OrigenResultados.ERROR_CARGA_DESCRIPCION}
    @property
    def es_error_configuracion(self) -> bool:
        # Propiedad para verificar si el origen fue un error de configuración de columnas.
        return self in {OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC}
    @property
    def es_error_operacional(self) -> bool: 
        # Propiedad para verificar si el origen fue un error operacional interno.
        return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR
    @property
    def es_termino_invalido(self) -> bool: 
        # Propiedad para verificar si el origen fue un término de búsqueda inválido.
        return self == OrigenResultados.TERMINO_INVALIDO

class ExtractorMagnitud:
    MAPEO_MAGNITUDES_PREDEFINIDO: Dict[str, List[str]] = {} # Mapeo predefinido (puede ser cargado externamente o estar vacío)

    def __init__(self, mapeo_magnitudes: Optional[Dict[str, List[str]]] = None):
        self.sinonimo_a_canonico_normalizado: Dict[str, str] = {} # Diccionario interno: sinónimo normalizado -> canónica normalizada
        # Utiliza el mapeo proporcionado o el predefinido.
        mapeo_a_usar = mapeo_magnitudes if mapeo_magnitudes is not None else self.MAPEO_MAGNITUDES_PREDEFINIDO
        
        for forma_canonica_original, lista_sinonimos_originales in mapeo_a_usar.items(): # Itera sobre el mapeo (canónica -> lista de sinónimos)
            canonico_norm = self._normalizar_texto(forma_canonica_original) # Normaliza la forma canónica
            if not canonico_norm: # Si la forma canónica normalizada es inválida, la ignora.
                logger.warning(f"Forma canónica '{forma_canonica_original}' resultó vacía tras normalizar y fue ignorada en ExtractorMagnitud.")
                continue
            
            self.sinonimo_a_canonico_normalizado[canonico_norm] = canonico_norm # Mapea la forma canónica normalizada a sí misma
            
            for sinonimo_original in lista_sinonimos_originales: # Itera sobre cada sinónimo
                sinonimo_norm = self._normalizar_texto(str(sinonimo_original)) # Normaliza el sinónimo
                if sinonimo_norm: # Si el sinónimo normalizado es válido
                    self.sinonimo_a_canonico_normalizado[sinonimo_norm] = canonico_norm # Mapea el sinónimo normalizado a la forma canónica normalizada
        logger.debug(f"ExtractorMagnitud inicializado/actualizado con {len(self.sinonimo_a_canonico_normalizado)} mapeos normalizados.")


    @staticmethod
    def _normalizar_texto(texto: str) -> str:
        # Normaliza un texto: mayúsculas, sin acentos, solo alfanuméricos y ciertos símbolos, espacios normalizados.
        if not isinstance(texto, str) or not texto: return "" # Retorna vacío si no es string o está vacío
        try:
            texto_upper = texto.upper() # Convierte a mayúsculas
            forma_normalizada = unicodedata.normalize("NFKD", texto_upper) # Normaliza a NFKD (descompone caracteres con acentos)
            # Elimina caracteres diacríticos (acentos) y conserva solo alfanuméricos, espacios y símbolos permitidos
            res = "".join(c for c in forma_normalizada if not unicodedata.combining(c) and (c.isalnum() or c.isspace() or c in ['.', '-', '_', '/']))
            return ' '.join(res.split()) # Reemplaza múltiples espacios por uno solo y quita espacios al inicio/final
        except TypeError: # Captura error si el texto no es procesable (ej. si fuera None a pesar del chequeo inicial)
            logger.error(f"TypeError en _normalizar_texto (ExtractorMagnitud) con entrada: {texto}")
            return ""

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        # Obtiene la forma canónica normalizada para un texto de unidad dado.
        if not texto_unidad: return None # Si la unidad es vacía, retorna None
        normalizada = self._normalizar_texto(texto_unidad) # Normaliza el texto de la unidad
        # Busca la unidad normalizada en el mapeo y retorna su forma canónica
        return self.sinonimo_a_canonico_normalizado.get(normalizada) if normalizada else None

class ManejadorExcel:
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        # Carga un archivo Excel (xls o xlsx) en un DataFrame de pandas.
        ruta = Path(ruta_archivo) # Convierte la ruta a un objeto Path para un manejo de rutas más robusto
        if not ruta.exists(): # Verifica si el archivo existe en la ruta especificada
            mensaje_error = f"¡Archivo no encontrado! Ruta: {ruta}"
            logger.error(f"ManejadorExcel: {mensaje_error}") # Registra el error
            return None, mensaje_error # Retorna None para el DataFrame y el mensaje de error
        try:
            engine: Optional[str] = None # Inicializa el motor de Excel como None
            if ruta.suffix.lower() == ".xlsx": engine = "openpyxl" # Usa 'openpyxl' para archivos .xlsx, pandas lo usa por defecto para xlsx
            logger.info(f"ManejadorExcel: Cargando '{ruta.name}' con engine='{engine or 'auto (pandas intentará xlrd para .xls)'}'...")
            df = pd.read_excel(ruta, engine=engine) # Lee el archivo Excel
            logger.info(f"ManejadorExcel: Archivo '{ruta.name}' ({len(df)} filas) cargado exitosamente.")
            return df, None # Retorna el DataFrame cargado y None para el mensaje de error
        except ImportError as ie: # Captura error si falta la librería necesaria (openpyxl o xlrd)
            mensaje_error_usuario = (f"Error al cargar '{ruta.name}': Falta librería.\nPara .xlsx: pip install openpyxl\nPara .xls: pip install xlrd\nDetalle: {ie}")
            logger.exception(f"ManejadorExcel: Falta dependencia para leer '{ruta.name}'. Error: {ie}") # Registra la excepción
            return None, mensaje_error_usuario
        except Exception as e: # Captura cualquier otra excepción durante la carga del archivo
            mensaje_error_usuario = (f"No se pudo cargar '{ruta.name}': {e}\nVerifique formato, permisos y si está en uso.")
            logger.exception(f"ManejadorExcel: Error genérico al cargar '{ruta.name}'.") # Registra la excepción
            return None, mensaje_error_usuario

class MotorBusqueda:
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None # DataFrame para el archivo de diccionario (FCDs)
        self.datos_descripcion: Optional[pd.DataFrame] = None # DataFrame para el archivo de descripciones de artículos
        self.archivo_diccionario_actual: Optional[Path] = None # Ruta del archivo de diccionario actualmente cargado
        self.archivo_descripcion_actual: Optional[Path] = None # Ruta del archivo de descripciones actualmente cargado
        # Índices de columnas a mostrar en la vista previa del diccionario en la UI. Si está vacío o es [-1], se usan todas las columnas de texto/objeto.
        self.indices_columnas_busqueda_dic_preview: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else []
        logger.info(f"MotorBusqueda inicializado. Índices preview dicc: {self.indices_columnas_busqueda_dic_preview or 'Todas texto/objeto'}")
        # Expresión regular para parsear términos de comparación numérica (ej. ">10V", "<=5A", "=2.5mm")
        self.patron_comparacion = re.compile(r"^\s*([<>]=?)\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$")
        # Expresión regular para parsear términos de rango numérico (ej. "10-20V", "2.5-3.5mm")
        self.patron_rango = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$")
        # Expresión regular para extraer términos negados (ej. #palabra, #"frase negada")
        self.patron_termino_negado = re.compile(r'#\s*(?:\"([^\"]+)\"|([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\.\-\_]+))', re.IGNORECASE | re.UNICODE)
        # Expresión regular para encontrar números y unidades dentro de las celdas de un DataFrame
        self.patron_num_unidad_df = re.compile(r"(\d+(?:[.,]\d+)?)[\s\-]*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?")
        self.extractor_magnitud = ExtractorMagnitud() # Inicializa el extractor de magnitudes (inicialmente vacío o con predefinidos)

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str) # Convierte la ruta a objeto Path
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta) # Intenta cargar el archivo

        if df_cargado is None: # Si la carga falla
            self.datos_diccionario = None # Resetea el DataFrame
            self.archivo_diccionario_actual = None # Resetea la ruta
            self.extractor_magnitud = ExtractorMagnitud() # Resetea el extractor de magnitudes al estado inicial
            return False, error_msg_carga # Devuelve fallo y mensaje de error

        mapeo_dinamico_para_extractor: Dict[str, List[str]] = {} # Diccionario para construir el mapeo de unidades/términos
        
        if df_cargado.shape[1] > 0: # Si el DataFrame cargado tiene al menos una columna
            columna_canonica_nombre = df_cargado.columns[0] # Asume que la primera columna contiene la forma canónica
            inicio_col_sinonimos = 3 # Asume que los sinónimos comienzan desde la cuarta columna (índice 3)
            max_cols_a_chequear_para_sinonimos = df_cargado.shape[1] # Considera todas las columnas restantes para sinónimos

            for _, fila in df_cargado.iterrows(): # Itera sobre cada fila del DataFrame del diccionario
                forma_canonica_raw = fila[columna_canonica_nombre] # Obtiene el valor de la forma canónica
                if pd.isna(forma_canonica_raw) or str(forma_canonica_raw).strip() == "": # Si es NaN o vacío, lo ignora
                    continue 

                forma_canonica_str = str(forma_canonica_raw).strip() # Convierte a string y limpia espacios
                # La forma canónica original es también un sinónimo de sí misma.
                sinonimos_para_esta_canonica: List[str] = [forma_canonica_str] 

                for i in range(inicio_col_sinonimos, max_cols_a_chequear_para_sinonimos): # Itera sobre las columnas de sinónimos
                    if i < df_cargado.shape[1]: # Asegura que el índice de columna sea válido
                        sinonimo_celda_raw = fila[df_cargado.columns[i]] # Obtiene el valor del sinónimo
                        if pd.notna(sinonimo_celda_raw) and str(sinonimo_celda_raw).strip() != "": # Si no es NaN ni vacío
                            sinonimos_para_esta_canonica.append(str(sinonimo_celda_raw).strip()) # Lo añade a la lista de sinónimos
                
                # La forma canónica (no normalizada aquí, ExtractorMagnitud lo hará) se usa como clave para el mapeo
                # que se pasará al constructor de ExtractorMagnitud.
                # El Extractor se encargará de la normalización final y de aplanar la estructura.
                mapeo_dinamico_para_extractor[forma_canonica_str] = list(set(sinonimos_para_esta_canonica)) # Elimina duplicados de sinónimos para esta clave

            if mapeo_dinamico_para_extractor: # Si se construyó algún mapeo
                # Crea una nueva instancia de ExtractorMagnitud con el mapeo dinámico extraído.
                self.extractor_magnitud = ExtractorMagnitud(mapeo_magnitudes=mapeo_dinamico_para_extractor)
                logger.info(f"Extractor de magnitudes actualizado desde '{ruta.name}' usando formas canónicas y sinónimos.")
            else: # Si no se pudieron extraer mapeos
                logger.warning(f"No se extrajeron mapeos de unidad válidos desde '{ruta.name}'. ExtractorMagnitud usará su predefinido (si existe) o estará vacío.")
                self.extractor_magnitud = ExtractorMagnitud() # Resetea al predefinido o vacío
        else: # Si el archivo de diccionario no tiene columnas
            logger.warning(f"El archivo de diccionario '{ruta.name}' no tiene columnas. No se pudo actualizar el extractor de magnitudes.")
            self.extractor_magnitud = ExtractorMagnitud() # Resetea

        self.datos_diccionario = df_cargado # Almacena el DataFrame cargado
        self.archivo_diccionario_actual = ruta # Almacena la ruta del archivo

        if logger.isEnabledFor(logging.DEBUG) and self.datos_diccionario is not None: # Si el logging DEBUG está activo
            logger.debug(f"Archivo de diccionario '{ruta.name}' cargado (primeras 3 filas):\n{self.datos_diccionario.head(3).to_string()}")
        return True, None # Retorna éxito

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str) # Convierte la ruta string a objeto Path
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta) # Intenta cargar el archivo
        if df_cargado is None: # Si la carga falla
            self.datos_descripcion = None; self.archivo_descripcion_actual = None # Resetea
            return False, error_msg_carga # Devuelve fallo
        self.datos_descripcion = df_cargado; self.archivo_descripcion_actual = ruta # Almacena DataFrame y ruta
        logger.info(f"Archivo de descripciones '{ruta.name}' cargado.")
        return True, None # Retorna éxito

    def _obtener_nombres_columnas_busqueda_df(self, df: pd.DataFrame, indices_cfg: List[int], tipo_busqueda: str) -> Tuple[Optional[List[str]], Optional[str]]:
        if df is None or df.empty: return None, f"DF para '{tipo_busqueda}' vacío." # Chequea si el DF está vacío
        columnas_disponibles = list(df.columns); num_cols_df = len(columnas_disponibles) # Obtiene nombres y cantidad de columnas
        if num_cols_df == 0: return None, f"DF '{tipo_busqueda}' sin columnas." # Chequea si hay columnas
        
        usar_columnas_por_defecto = not indices_cfg or indices_cfg == [-1] # Determina si usar las columnas por defecto
        if usar_columnas_por_defecto: # Si es así
            # Selecciona todas las columnas de tipo string u object
            cols_texto_obj = [col for col in columnas_disponibles if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
            if cols_texto_obj: # Si se encontraron
                logger.debug(f"Para '{tipo_busqueda}', usando columnas de tipo texto/objeto (defecto): {cols_texto_obj}")
                return cols_texto_obj, None
            else: # Si no, usa todas las columnas
                logger.warning(f"Para '{tipo_busqueda}', no hay cols texto/objeto. Usando todas las {num_cols_df} columnas: {columnas_disponibles}")
                return columnas_disponibles, None
        
        nombres_columnas_seleccionadas: List[str] = [] # Lista para nombres de columnas válidas
        indices_invalidos: List[str] = [] # Lista para índices inválidos
        for i in indices_cfg: # Itera sobre los índices configurados
            if not (isinstance(i, int) and 0 <= i < num_cols_df): indices_invalidos.append(str(i)) # Valida cada índice
            else: nombres_columnas_seleccionadas.append(columnas_disponibles[i]) # Si es válido, añade el nombre de la columna
        
        if indices_invalidos: return None, f"Índice(s) {', '.join(indices_invalidos)} inválido(s) para '{tipo_busqueda}'. Columnas: {num_cols_df} (0 a {num_cols_df-1})."
        if not nombres_columnas_seleccionadas: return None, f"Config de índices {indices_cfg} no resultó en columnas válidas para '{tipo_busqueda}'."
        
        logger.debug(f"Para '{tipo_busqueda}', usando columnas por índices {indices_cfg}: {nombres_columnas_seleccionadas}")
        return nombres_columnas_seleccionadas, None

    def _normalizar_para_busqueda(self, texto: str) -> str:
        # Normaliza el texto para la búsqueda: mayúsculas, sin acentos, solo alfanuméricos y ciertos símbolos, espacios normalizados.
        if not isinstance(texto, str) or not texto: return "" # Retorna vacío si no es string o está vacío
        try:
            texto_upper = texto.upper() # Convierte a mayúsculas
            texto_norm_nfkd = unicodedata.normalize('NFKD', texto_upper) # Normaliza a NFKD para separar diacríticos
            texto_sin_acentos = "".join([c for c in texto_norm_nfkd if not unicodedata.combining(c)]) # Elimina diacríticos
            # Elimina caracteres no deseados, conservando alfanuméricos, espacios y . - / _
            texto_limpio_final = re.sub(r'[^\w\s\.\-\/\_]', '', texto_sin_acentos) 
            return ' '.join(texto_limpio_final.split()).strip() # Normaliza espacios y limpia extremos
        except Exception as e: # Captura cualquier error durante la normalización
            logger.error(f"Error al normalizar el texto '{texto[:50]}...': {e}")
            return str(texto).upper().strip() # Fallback a conversión simple a mayúsculas y strip

    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, cols: List[str], texto: str) -> Tuple[pd.DataFrame, str, List[str]]:
        texto_limpio_entrada = texto.strip(); terminos_negados_encontrados: List[str] = [] # Inicializa variables
        df_a_procesar = df_original.copy() if df_original is not None else pd.DataFrame() # Copia DataFrame o crea uno vacío
        if not texto_limpio_entrada: return df_a_procesar, "", terminos_negados_encontrados # Si el texto es vacío, retorna
        
        partes_positivas: List[str] = []; ultimo_indice_fin_negado = 0
        for match_negado in self.patron_termino_negado.finditer(texto_limpio_entrada): # Itera sobre términos negados encontrados
            partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:match_negado.start()]) # Añade parte positiva previa
            ultimo_indice_fin_negado = match_negado.end() # Actualiza el índice
            termino_negado_raw = match_negado.group(1) or match_negado.group(2) # Extrae el término negado (con o sin comillas)
            if termino_negado_raw:
                termino_negado_normalizado = self._normalizar_para_busqueda(termino_negado_raw.strip('"')) # Normaliza el término negado
                if termino_negado_normalizado and termino_negado_normalizado not in terminos_negados_encontrados:
                    terminos_negados_encontrados.append(termino_negado_normalizado) # Añade a la lista de negados únicos
        partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:]) # Añade la última parte positiva
        terminos_positivos_final_str = ' '.join("".join(partes_positivas).split()).strip() # Concatena y limpia los términos positivos

        if df_a_procesar.empty or not terminos_negados_encontrados or not cols: # Si no hay que filtrar por negación
            logger.debug(f"Parseo negación: Query='{texto_limpio_entrada}', Positivos='{terminos_positivos_final_str}', Negados={terminos_negados_encontrados}. No se aplicó filtro al DF.")
            return df_a_procesar, terminos_positivos_final_str, terminos_negados_encontrados
        
        mascara_exclusion_total = pd.Series(False, index=df_a_procesar.index) # Máscara para identificar filas a excluir
        for termino_negado_actual in terminos_negados_encontrados: # Itera sobre los términos negados
            if not termino_negado_actual: continue # Salta si el término negado es vacío
            mascara_para_este_termino_negado = pd.Series(False, index=df_a_procesar.index) # Máscara para el término negado actual
            patron_regex_negado = r"\b" + re.escape(termino_negado_actual) + r"\b" # Regex para buscar el término como palabra completa
            for nombre_columna in cols: # Itera sobre las columnas de búsqueda
                if nombre_columna not in df_a_procesar.columns: continue # Salta si la columna no existe
                try:
                    serie_columna_normalizada = df_a_procesar[nombre_columna].astype(str).map(self._normalizar_para_busqueda) # Normaliza la columna
                    mascara_para_este_termino_negado |= serie_columna_normalizada.str.contains(patron_regex_negado, regex=True, na=False) # Aplica OR a la máscara
                except Exception as e_neg_col: logger.error(f"Error aplicando negación en col '{nombre_columna}', term '{termino_negado_actual}': {e_neg_col}")
            mascara_exclusion_total |= mascara_para_este_termino_negado # Acumula la máscara de exclusión
        
        df_resultado_filtrado = df_a_procesar[~mascara_exclusion_total] # Filtra el DataFrame
        logger.info(f"Filtrado por negación (Query='{texto_limpio_entrada}'): {len(df_a_procesar)} -> {len(df_resultado_filtrado)} filas. Negados: {terminos_negados_encontrados}. Positivos: '{terminos_positivos_final_str}'")
        return df_resultado_filtrado, terminos_positivos_final_str, terminos_negados_encontrados

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        texto_limpio = texto_complejo.strip() # Limpia el texto
        if not texto_limpio: return "OR", [] # Si está vacío, retorna OR con lista vacía
        # Si hay '+' y no está entre paréntesis (ej. "A + (B|C)"), trátalo como AND a alto nivel
        if '+' in texto_limpio and not (texto_limpio.startswith("(") and texto_limpio.endswith(")")):
             logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Detectado '+' de alto nivel, tratando como AND. Segmento=['{texto_limpio}']")
             return "AND", [texto_limpio] # El segmento es toda la query

        separadores_or = [(r"\s*\|\s*", "|")] # Separador OR es '|' (con espacios opcionales alrededor)
        for sep_regex, sep_char_literal in separadores_or:
            # Si no hay '+' a alto nivel y el separador OR está presente
            if '+' not in texto_complejo and sep_char_literal in texto_limpio: # Verifica presencia de separador OR literal
                # Divide por el separador OR
                segmentos_potenciales = [s.strip() for s in re.split(sep_regex, texto_limpio) if s.strip()] # Divide y filtra vacíos
                # Asegura que la división fue efectiva
                if len(segmentos_potenciales) > 1 or (len(segmentos_potenciales) == 1 and texto_limpio != segmentos_potenciales[0]):
                    logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Op=OR, Segs={segmentos_potenciales}")
                    return "OR", segmentos_potenciales # Retorna OR y los segmentos
        # Si no se encontró un OR explícito a alto nivel, se asume AND (o un único término)
        logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Op=AND (no OR explícito de alto nivel), Seg=['{texto_limpio}']")
        return "AND", [texto_limpio] # El segmento es toda la query

    def _descomponer_nivel2_and(self, termino_segmento_n1: str) -> Tuple[str, List[str]]:
        termino_limpio = termino_segmento_n1.strip() # Limpia el segmento
        if not termino_limpio: return "AND", [] # Si está vacío, retorna AND con lista vacía
        partes_crudas = re.split(r'\s+\+\s+', termino_limpio) # Divide por ' + ' para obtener términos AND
        partes_limpias_finales = [p.strip() for p in partes_crudas if p.strip()] # Limpia y filtra partes vacías
        logger.debug(f"Descomp. N2 (AND) para '{termino_segmento_n1}': Partes={partes_limpias_finales}")
        return "AND", partes_limpias_finales # Retorna AND y los términos

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        terminos_analizados: List[Dict[str, Any]] = [] # Lista para almacenar los términos analizados
        for termino_original_bruto in terminos_brutos: # Itera sobre cada término bruto
            termino_original_procesado = str(termino_original_bruto).strip() # Convierte a string y limpia espacios
            es_frase_exacta = False
            termino_final_para_analisis = termino_original_procesado
            # Comprueba si el término es una frase exacta (entre comillas dobles)
            if len(termino_final_para_analisis) >= 2 and \
               termino_final_para_analisis.startswith('"') and \
               termino_final_para_analisis.endswith('"'):
                termino_final_para_analisis = termino_final_para_analisis[1:-1] # Quita las comillas
                es_frase_exacta = True
            if not termino_final_para_analisis: continue # Si queda vacío, lo salta
            
            item_analizado: Dict[str, Any] = {"original": termino_final_para_analisis} # Inicializa el diccionario del término
            match_comparacion = self.patron_comparacion.match(termino_final_para_analisis) # Intenta parsear como comparación numérica
            match_rango = self.patron_rango.match(termino_final_para_analisis) # Intenta parsear como rango numérico
            
            if match_comparacion and not es_frase_exacta: # Si es una comparación numérica y no una frase exacta
                operador_str, valor_str, unidad_str_raw = match_comparacion.groups() # Extrae operador, valor y unidad
                valor_numerico = self._parse_numero(valor_str) # Parsea el valor numérico
                if valor_numerico is not None: # Si el parseo fue exitoso
                    mapa_operadores = {">": "gt", "<": "lt", ">=": "ge", "<=": "le", "=": "eq"} # Mapeo de operadores
                    unidad_canonica: Optional[str] = None
                    if unidad_str_raw and unidad_str_raw.strip(): # Si se proporcionó una unidad
                        unidad_canonica = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_raw.strip()) # Normaliza la unidad
                    item_analizado.update({"tipo": mapa_operadores.get(operador_str), "valor": valor_numerico, "unidad_busqueda": unidad_canonica})
                else: # Si el parseo numérico falló, se trata como término de string
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            elif match_rango and not es_frase_exacta: # Si es un rango numérico y no una frase exacta
                valor1_str, valor2_str, unidad_str_r_raw = match_rango.groups() # Extrae valores y unidad del rango
                valor1_num = self._parse_numero(valor1_str); valor2_num = self._parse_numero(valor2_str) # Parsea ambos valores
                if valor1_num is not None and valor2_num is not None: # Si ambos parseos son exitosos
                    unidad_canonica_r: Optional[str] = None
                    if unidad_str_r_raw and unidad_str_r_raw.strip(): # Si se proporcionó una unidad
                        unidad_canonica_r = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_r_raw.strip()) # Normaliza la unidad
                    item_analizado.update({"tipo": "range", "valor": sorted([valor1_num, valor2_num]), "unidad_busqueda": unidad_canonica_r})
                else: # Si el parseo numérico falló, se trata como término de string
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            else: # Si no es ni comparación ni rango, es un término de string (o frase exacta)
                item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            terminos_analizados.append(item_analizado) # Añade el término analizado a la lista
        logger.debug(f"Términos (post-AND) analizados para búsqueda detallada: {terminos_analizados}")
        return terminos_analizados

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        # Mantiene la conversión directa si ya es numérico.
        if isinstance(num_str, (int, float)):
            return float(num_str)
        # Retorna None si no es string o está vacío después de limpiar.
        if not isinstance(num_str, str):
            logger.debug(f"Parseo num: Entrada '{num_str}' no es string.") # Log para identificar entradas no string
            return None
        s_limpio = num_str.strip() # Elimina espacios al inicio y al final del string.
        if not s_limpio:
            logger.debug(f"Parseo num: Entrada '{num_str}' vacía tras limpiar.") # Log para identificar entradas vacías
            return None

        logger.debug(f"Parseo num: Intentando convertir '{s_limpio}' (originado de '{num_str}')") # Log general al inicio de la función

        try:
            # Regla 1: Si no hay comas ni puntos, es un número simple (entero o float sin separadores).
            if ',' not in s_limpio and '.' not in s_limpio:
                logger.debug(f"  '{s_limpio}': Sin separadores. Intento de float directo.") # Log para la rama de sin separadores
                return float(s_limpio) # Intenta convertir directamente a float.

            # Regla 2: Hay comas. La coma puede ser decimal o de miles.
            # Prioridad a la coma como potencial separador decimal principal.
            if ',' in s_limpio:
                logger.debug(f"  '{s_limpio}': Contiene comas. Procesando según reglas de coma.") # Log para la rama de comas detectadas
                partes_coma = s_limpio.split(',') # Divide la cadena por comas.
                
                # Parte entera es todo antes de la primera coma.
                parte_entera_antes_primera_coma_str = partes_coma[0].strip() # Toma la primera parte y limpia espacios.
                
                # Heurística del cero inicial (ej. "09" en "09,10" o "09,100")
                # Se considera relevante si empieza con '0', tiene más de un dígito, y el resto son dígitos (ej. "09" pero no "0" solo o "0x").
                es_cero_inicial_relevante = parte_entera_antes_primera_coma_str.startswith('0') and \
                                           len(parte_entera_antes_primera_coma_str) > 1 and \
                                           parte_entera_antes_primera_coma_str[1:].isdigit()

                # Caso A: Una sola coma (ej. "9,10", "09,100", "9,100", "1.234,56")
                if len(partes_coma) == 2:
                    # Quita puntos de miles de la parte entera (ej. "1.234" de "1.234,56" -> "1234")
                    parte_entera_limpia_str = parte_entera_antes_primera_coma_str.replace('.', '') 
                    parte_decimal_str = partes_coma[1].strip() # Toma la parte decimal y limpia espacios.

                    # Antes de proceder, verifica si las partes son realmente numéricas.
                    if not (parte_entera_limpia_str.isdigit() or (parte_entera_limpia_str.startswith('-') and parte_entera_limpia_str[1:].isdigit())) or \
                       not parte_decimal_str.isdigit():
                        logger.warning(f"    Partes no numéricas alrededor de coma única: '{s_limpio}' -> entera:'{parte_entera_limpia_str}', decimal:'{parte_decimal_str}'")
                        raise ValueError("Partes no numéricas con coma única.") # Lanza error si no son numéricas.

                    if es_cero_inicial_relevante: 
                        # "09,10" o "09,100" -> float (ej. 9.10 o 9.100). Coma es decimal.
                        numero_reconstruido = f"{parte_entera_limpia_str}.{parte_decimal_str}" # Reconstruye como "entera.decimal"
                        logger.debug(f"    Coma única con cero inicial ('{parte_entera_antes_primera_coma_str}'): '{s_limpio}' -> float '{numero_reconstruido}'")
                        return float(numero_reconstruido) # Convierte a float.
                    elif len(parte_decimal_str) == 3 and not es_cero_inicial_relevante : 
                        # "9,100" (pero no "09,100") -> entero (9100). Coma es separador de miles.
                        numero_reconstruido = f"{parte_entera_limpia_str}{parte_decimal_str}" # Reconstruye uniendo las partes.
                        logger.debug(f"    Coma única, 3 dig post-coma, no cero inicial relevante: '{s_limpio}' -> entero '{numero_reconstruido}'")
                        return float(numero_reconstruido) # Convierte a float (que será un entero .0).
                    else: 
                        # Otros casos con una coma: "9,10", "1.234,56" (interpretado como 1234.56). La coma es decimal.
                        # También cubre "0,123" -> 0.123
                        numero_reconstruido = f"{parte_entera_limpia_str}.{parte_decimal_str}" # Reconstruye como "entera.decimal"
                        logger.debug(f"    Coma única, otro (tratado como decimal): '{s_limpio}' -> float '{numero_reconstruido}'")
                        return float(numero_reconstruido) # Convierte a float.
                
                # Caso B: Múltiples comas (ej. "1,234,567" o "1,234,56")
                elif len(partes_coma) > 2:
                    parte_decimal_final_m_comas = partes_coma[-1].strip() # La última parte es la candidata a decimal.
                    # Une todas las partes antes de la última coma, quitando puntos internos (si los hubiera).
                    parte_entera_reconstruida_m_comas = "".join(p.replace('.', '') for p in partes_coma[:-1]) 
                    
                    # Validar que las partes sean dígitos.
                    if not (parte_entera_reconstruida_m_comas.isdigit() or (parte_entera_reconstruida_m_comas.startswith('-') and parte_entera_reconstruida_m_comas[1:].isdigit())) or \
                       not parte_decimal_final_m_comas.isdigit():
                        logger.warning(f"    Partes no numéricas con múltiples comas: '{s_limpio}'")
                        raise ValueError("Partes no numéricas con múltiples comas.")

                    # Si la última parte tiene 3 dígitos Y no hay cero inicial relevante al inicio del número Y es todo dígitos -> se interpreta como parte de un entero grande.
                    if len(parte_decimal_final_m_comas) == 3 and not es_cero_inicial_relevante:
                         numero_reconstruido = f"{parte_entera_reconstruida_m_comas}{parte_decimal_final_m_comas}" # ej. "1,234,567" -> "1234" + "567" = "1234567"
                         logger.debug(f"  Múltiples comas, interpretado como entero grande (3 dig post-última coma): '{s_limpio}' -> '{numero_reconstruido}'")
                         return float(numero_reconstruido)
                    else: # Se asume que la última coma era decimal. Ej "1,234,56" -> 1234.56
                        numero_reconstruido = f"{parte_entera_reconstruida_m_comas}.{parte_decimal_final_m_comas}" # ej. "1,234,56" -> "1234" + "." + "56" = "1234.56"
                        logger.debug(f"  Múltiples comas, última coma tratada como decimal: '{s_limpio}' -> float '{numero_reconstruido}'")
                        return float(numero_reconstruido)

            # Regla 3: No hay comas, pero SÍ hay puntos. El punto puede ser decimal o de miles.
            elif '.' in s_limpio: 
                logger.debug(f"  '{s_limpio}': Contiene puntos, sin comas. Procesando según reglas de punto.")
                partes_punto = s_limpio.split('.') # Divide la cadena por puntos.
                parte_entera_punto_str = partes_punto[0].strip() # Toma la primera parte como entera.
                # Heurística del cero inicial para puntos.
                es_cero_inicial_relevante_punto = parte_entera_punto_str.startswith('0') and \
                                                  len(parte_entera_punto_str) > 1 and \
                                                  parte_entera_punto_str[1:].isdigit()

                if len(partes_punto) == 2: # Un solo punto: "09.10", "9.100", "9.10"
                    parte_decimal_punto_str = partes_punto[1].strip() # Toma la parte decimal.
                    # Validar que ambas partes sean numéricas
                    if not (parte_entera_punto_str.isdigit() or (parte_entera_punto_str.startswith('-') and parte_entera_punto_str[1:].isdigit())) or \
                       not parte_decimal_punto_str.isdigit():
                        logger.warning(f"    Partes no numéricas alrededor de punto único: '{s_limpio}' -> entera:'{parte_entera_punto_str}', decimal:'{parte_decimal_punto_str}'")
                        raise ValueError("Partes no numéricas con punto único.")

                    if es_cero_inicial_relevante_punto: # "09.10", "09.100" -> float
                        logger.debug(f"    Punto único con cero inicial ('{parte_entera_punto_str}'): '{s_limpio}' -> float")
                        return float(s_limpio) # Python maneja "09.100" como 9.1
                    elif len(parte_decimal_punto_str) == 3 and not es_cero_inicial_relevante_punto: # "9.100" -> entero (9100)
                        numero_reconstruido = f"{parte_entera_punto_str}{parte_decimal_punto_str}" # Reconstruye uniendo las partes.
                        logger.debug(f"    Punto único, 3 dec, no cero inicial: '{s_limpio}' -> entero '{numero_reconstruido}'")
                        return float(numero_reconstruido) # Convierte a float (entero.0).
                    else: # "9.10" o "123.45" -> float
                        logger.debug(f"    Punto único, otro (tratado como decimal): '{s_limpio}' -> float")
                        return float(s_limpio) # Python maneja "9.10", "123.45" bien.
                
                elif len(partes_punto) > 2: # Múltiples puntos: "1.234.567" -> separadores de miles
                    # Asegurarse que todas las partes sean dígitos
                    if not all(p.isdigit() for p in partes_punto):
                        logger.warning(f"    Partes no numéricas con múltiples puntos: '{s_limpio}'")
                        raise ValueError("Partes no numéricas con múltiples puntos.")
                    numero_reconstruido_entero_miles = "".join(partes_punto) # Une todas las partes sin puntos.
                    logger.debug(f"    Múltiples puntos, asumido como separador de miles: '{s_limpio}' -> entero '{numero_reconstruido_entero_miles}'")
                    return float(numero_reconstruido_entero_miles)
            
            # Fallback si la lógica anterior no cubrió el caso (ej. un string que ya es un float válido como "123.45" y no activó reglas de "09.xxx" o "9.xxx")
            # Esto podría ser alcanzado si s_limpio tiene un solo punto pero no cumple las sub-condiciones de len(partes_punto)==2.
            logger.debug(f"Parseo num: '{num_str}' -> Fallback o caso simple no cubierto específicamente. Intentando float directo en '{s_limpio}'")
            return float(s_limpio) # Intenta convertir directamente.

        except ValueError: # Captura ValueError de las conversiones float o las que se lanzan manualmente en las validaciones isdigit.
            logger.warning(f"Parseo num: ValueError final al convertir '{s_limpio}' (originado de '{num_str}') a float.")
            return None
        except Exception as e: # Captura cualquier otra excepción inesperada durante el parseo.
            logger.error(f"Parseo num: Excepción inesperada '{type(e).__name__}' para '{s_limpio}' (originado de '{num_str}'): {e}")
            return None

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols: List[str], term_an: Dict[str, Any]) -> pd.Series:
        # No hay cambios en esta función más allá de la llamada a la nueva _parse_numero.
        # Los comentarios permanecen igual que en la versión original proporcionada por el usuario.
        tipo_termino = term_an["tipo"]; valor_termino = term_an["valor"]; unidad_requerida_canonica = term_an.get("unidad_busqueda")
        mascara_total_termino = pd.Series(False, index=df.index)
        for nombre_columna in cols:
            if nombre_columna not in df.columns: continue
            columna_serie = df[nombre_columna]; mascara_columna_actual_numerica = pd.Series(False, index=df.index)
            if tipo_termino in ["gt", "lt", "ge", "le", "range", "eq"]:
                for indice_fila, valor_celda_raw in columna_serie.items():
                    if pd.isna(valor_celda_raw) or str(valor_celda_raw).strip() == "": continue
                    for match_num_unidad_celda in self.patron_num_unidad_df.finditer(str(valor_celda_raw)):
                        try:
                            num_celda_str = match_num_unidad_celda.group(1)
                            num_celda_val = self._parse_numero(num_celda_str) 
                            u_c_raw = match_num_unidad_celda.group(2)
                            if num_celda_val is None: continue
                            u_c_canon = self.extractor_magnitud.obtener_magnitud_normalizada(u_c_raw.strip()) if u_c_raw and u_c_raw.strip() else None
                            u_ok = (unidad_requerida_canonica is None) or \
                                   (u_c_canon is not None and u_c_canon == unidad_requerida_canonica) or \
                                   (u_c_raw and unidad_requerida_canonica and self.extractor_magnitud._normalizar_texto(u_c_raw.strip()) == unidad_requerida_canonica)
                            if not u_ok: continue
                            cond = False
                            if tipo_termino == "eq" and np.isclose(num_celda_val, valor_termino): cond = True
                            elif tipo_termino == "gt" and num_celda_val > valor_termino and not np.isclose(num_celda_val, valor_termino): cond = True
                            elif tipo_termino == "lt" and num_celda_val < valor_termino and not np.isclose(num_celda_val, valor_termino): cond = True
                            elif tipo_termino == "ge" and (num_celda_val >= valor_termino or np.isclose(num_celda_val, valor_termino)): cond = True
                            elif tipo_termino == "le" and (num_celda_val <= valor_termino or np.isclose(num_celda_val, valor_termino)): cond = True
                            elif tipo_termino == "range" and ((valor_termino[0] <= num_celda_val or np.isclose(num_celda_val, valor_termino[0])) and \
                                                               (num_celda_val <= valor_termino[1] or np.isclose(num_celda_val, valor_termino[1]))): cond = True
                            if cond: mascara_columna_actual_numerica.at[indice_fila] = True; break
                        except ValueError: continue
                mascara_total_termino |= mascara_columna_actual_numerica
            elif tipo_termino == "str":
                try:
                    val_norm_busq = str(valor_termino);
                    if not val_norm_busq: continue
                    serie_norm_df_col = columna_serie.astype(str).map(self._normalizar_para_busqueda)
                    pat_regex = r"\b" + re.escape(val_norm_busq) + r"\b"
                    mascara_col_actual = serie_norm_df_col.str.contains(pat_regex, regex=True, na=False)
                    mascara_total_termino |= mascara_col_actual
                except Exception as e: logger.warning(f"Error búsqueda STR col '{nombre_columna}' para '{valor_termino}': {e}")
        return mascara_total_termino

    def _aplicar_mascara_combinada_para_segmento_and(self, df: pd.DataFrame, cols: List[str], term_an_seg: List[Dict[str, Any]]) -> pd.Series:
        if df is None or df.empty or not cols: return pd.Series(False, index=df.index if df is not None else None)
        if not term_an_seg: return pd.Series(False, index=df.index)
        mascara_final = pd.Series(True, index=df.index) # Inicia con todos True para la operación AND
        for term_ind_an in term_an_seg: # Itera sobre cada término del segmento AND
            if term_ind_an["tipo"] == "str" and \
               ("|" in term_ind_an["original"] or "/" in term_ind_an["original"]) and \
               term_ind_an["original"].startswith("(") and term_ind_an["original"].endswith(")"): # Si es una sub-query OR encapsulada
                logger.debug(f"Segmento AND contiene sub-query OR: '{term_ind_an['original']}'. Se procesará por separado.")
                # Llama recursivamente a _procesar_busqueda_en_df_objetivo para obtener la máscara de la sub-query OR
                sub_mascara_or_series, err_sub_or = self._procesar_busqueda_en_df_objetivo(df, cols, term_ind_an["original"], None, return_mask_only=True) 
                if err_sub_or or sub_mascara_or_series is None: # Si hay error o no se devuelve máscara
                    logger.warning(f"Sub-query OR '{term_ind_an['original']}' falló o no devolvió máscara: {err_sub_or}")
                    return pd.Series(False, index=df.index) # El AND completo falla
                mascara_este_term = sub_mascara_or_series.reindex(df.index, fill_value=False) # Alinea la máscara
            else: # Si es un término atómico normal
                mascara_este_term = self._generar_mascara_para_un_termino(df, cols, term_ind_an) # Genera la máscara para este término
            mascara_final &= mascara_este_term # Aplica AND lógico a la máscara final
            if not mascara_final.any(): break # Optimización: si ya no hay True, el resultado del AND será False
        return mascara_final

    def _combinar_mascaras_de_segmentos_or(self, lista_mascaras: List[pd.Series], df_idx_ref: Optional[pd.Index] = None) -> pd.Series:
        if not lista_mascaras: # Si no hay máscaras para combinar
            # Retorna una Serie vacía o con el índice de referencia, toda False
            return pd.Series(False, index=df_idx_ref) if df_idx_ref is not None else pd.Series(dtype=bool)
        
        idx_usar = df_idx_ref # Usa el índice de referencia si se proporciona
        if idx_usar is None or idx_usar.empty: # Si no hay índice de referencia, intenta tomarlo de la primera máscara no vacía
            if lista_mascaras and not lista_mascaras[0].empty:
                idx_usar = lista_mascaras[0].index
        
        if idx_usar is None or idx_usar.empty: # Si aún no hay un índice válido (ej. todas las máscaras están vacías)
            return pd.Series(dtype=bool) # Retorna una Serie booleana vacía

        mascara_final = pd.Series(False, index=idx_usar) # Inicializa la máscara final con todos False
        for masc_seg in lista_mascaras: # Itera sobre cada máscara de segmento
            if masc_seg.empty: continue # Salta máscaras vacías
            mascara_alineada = masc_seg
            if not masc_seg.index.equals(idx_usar): # Si los índices no coinciden
                try: # Intenta reindexar la máscara del segmento al índice de referencia
                    mascara_alineada = masc_seg.reindex(idx_usar, fill_value=False)
                except Exception as e_reidx: 
                    logger.error(f"Fallo reindex máscara OR: {e_reidx}. Máscara ignorada."); continue # Registra error y salta
            mascara_final |= mascara_alineada # Combina con OR lógico
        return mascara_final

    def _procesar_busqueda_en_df_objetivo(self, df_obj: pd.DataFrame, cols_obj: List[str], termino_busqueda_original_para_este_df: str, terminos_negativos_adicionales: Optional[List[str]] = None, return_mask_only: bool = False) -> Union[Tuple[pd.DataFrame, Optional[str]], Tuple[Optional[pd.Series], Optional[str]]]:
        logger.debug(f"Proc. búsqueda DF: Query='{termino_busqueda_original_para_este_df}' en {len(cols_obj)} cols de DF ({len(df_obj if df_obj is not None else [])} filas). Neg. Adic: {terminos_negativos_adicionales}, ReturnMask: {return_mask_only}")
        
        if df_obj is None: # Si el DataFrame objetivo no existe
             df_obj = pd.DataFrame() # Usa un DataFrame vacío para evitar errores
        
        # Aplica negaciones de la query actual y extrae términos positivos
        df_despues_negaciones_query, terminos_positivos_de_query, _ = \
            self._aplicar_negaciones_y_extraer_positivos(df_obj, cols_obj, termino_busqueda_original_para_este_df)
        
        df_actual_procesando = df_despues_negaciones_query # DataFrame a usar para la búsqueda positiva
        
        # Aplica términos negativos adicionales si existen y el DataFrame no está ya vacío
        if terminos_negativos_adicionales and not df_actual_procesando.empty:
            # Construye una query que solo contenga los términos negativos adicionales
            query_solo_negativos_adicionales = " ".join([f"#{neg}" for neg in terminos_negativos_adicionales if neg])
            if query_solo_negativos_adicionales: # Si hay algo que negar
                logger.debug(f"Aplicando neg. ADICIONALES: '{query_solo_negativos_adicionales}' a {len(df_actual_procesando)} filas.")
                # Filtra el DataFrame actual usando solo estos términos negativos
                df_actual_procesando, _, _ = self._aplicar_negaciones_y_extraer_positivos(df_actual_procesando, cols_obj, query_solo_negativos_adicionales)
                logger.info(f"Filtrado por neg. ADICIONALES: {len(df_despues_negaciones_query)} -> {len(df_actual_procesando)} filas.")

        terminos_positivos_final_para_parseo = terminos_positivos_de_query # Términos positivos a usar para la búsqueda principal
        
        # Si el DataFrame está vacío después de las negaciones y no hay términos positivos, retorna vacío
        if df_actual_procesando.empty and not terminos_positivos_final_para_parseo.strip():
            logger.debug("DF vacío post-negaciones y sin términos positivos. Devolviendo DF/Máscara vacía.")
            idx_ref = df_obj.index if df_obj is not None else None # Para mantener el índice original si es posible
            return (pd.Series(False, index=idx_ref) if return_mask_only else df_actual_procesando.copy()), None
        
        # Si no hay términos positivos, el resultado es el DataFrame después de aplicar las negaciones
        if not terminos_positivos_final_para_parseo.strip():
            logger.debug(f"Sin términos positivos ('{terminos_positivos_final_para_parseo}'). Devolviendo DF/Máscara post-negaciones ({len(df_actual_procesando)} filas).")
            return (pd.Series(True, index=df_actual_procesando.index) if return_mask_only and not df_actual_procesando.empty else pd.Series(False, index=df_obj.index if df_obj is not None else None)), \
                   None if return_mask_only else df_actual_procesando.copy(), \
                   None

        operador_nivel1, segmentos_nivel1_or = self._descomponer_nivel1_or(terminos_positivos_final_para_parseo) # Descompone la query en segmentos OR
        
        if not segmentos_nivel1_or: # Si no hay segmentos válidos tras la descomposición OR
            if termino_busqueda_original_para_este_df.strip() or terminos_positivos_final_para_parseo.strip(): # Si había algo que parsear
                msg_error_segmentos = f"Térm. positivo '{terminos_positivos_final_para_parseo}' (de '{termino_busqueda_original_para_este_df}') inválido post-OR."
                logger.warning(msg_error_segmentos)
                return (pd.Series(False, index=df_actual_procesando.index if not df_actual_procesando.empty else None) if return_mask_only else pd.DataFrame(columns=df_actual_procesando.columns)), msg_error_segmentos
            else: # Si todo estaba vacío
                logger.debug("Query original y positiva post-negación vacías. Devolviendo DF/Máscara post-negaciones.")
                return (pd.Series(True, index=df_actual_procesando.index if not df_actual_procesando.empty else None) if return_mask_only else df_actual_procesando.copy()), None

        lista_mascaras_para_or: List[pd.Series] = [] # Lista para almacenar máscaras de cada segmento OR
        for segmento_or_actual in segmentos_nivel1_or: # Itera sobre cada segmento OR
            _operador_nivel2, terminos_brutos_nivel2_and = self._descomponer_nivel2_and(segmento_or_actual) # Descompone el segmento en términos AND
            terminos_atomicos_analizados_and = self._analizar_terminos(terminos_brutos_nivel2_and) # Analiza los términos atómicos
            
            mascara_para_segmento_or_actual: pd.Series
            if not terminos_atomicos_analizados_and: # Si no hay términos atómicos válidos en este segmento
                if operador_nivel1 == "AND": # Si el operador de nivel 1 era AND, esto es un fallo para todo el AND
                    msg_error_and = f"Segmento AND '{segmento_or_actual}' sin términos atómicos válidos. Falla."
                    logger.warning(msg_error_and)
                    return (pd.Series(False, index=df_actual_procesando.index if not df_actual_procesando.empty else None) if return_mask_only else pd.DataFrame(columns=df_actual_procesando.columns)), msg_error_and
                # Si el operador era OR, un segmento sin términos se ignora (máscara False)
                logger.debug(f"Segmento OR '{segmento_or_actual}' sin términos atómicos. Se ignora para OR.")
                mascara_para_segmento_or_actual = pd.Series(False, index=df_actual_procesando.index if not df_actual_procesando.empty else None)
            else: # Si hay términos atómicos, aplica la máscara combinada AND
                mascara_para_segmento_or_actual = self._aplicar_mascara_combinada_para_segmento_and(df_actual_procesando, cols_obj, terminos_atomicos_analizados_and)
            lista_mascaras_para_or.append(mascara_para_segmento_or_actual) # Añade la máscara del segmento a la lista

        idx_ref_or = df_actual_procesando.index if not df_actual_procesando.empty else (df_obj.index if df_obj is not None else None)
        if not lista_mascaras_para_or : # Si no se generaron máscaras (ej. todos los segmentos OR eran inválidos)
             logger.warning("No se generaron máscaras OR válidas.")
             return (pd.Series(False, index=idx_ref_or) if return_mask_only else pd.DataFrame(columns=df_obj.columns if df_obj is not None else [])), "No se generaron máscaras OR válidas."

        # Combina todas las máscaras de segmento OR
        mascara_final_df_objetivo = self._combinar_mascaras_de_segmentos_or(lista_mascaras_para_or, idx_ref_or)
        
        if return_mask_only: # Si solo se debe devolver la máscara
            logger.debug(f"Devolviendo solo máscara para '{termino_busqueda_original_para_este_df}': {mascara_final_df_objetivo.sum()} coincidencias.")
            return mascara_final_df_objetivo, None
        else: # Si se debe devolver el DataFrame filtrado
            df_resultado_final: pd.DataFrame
            if mascara_final_df_objetivo.empty: # Si la máscara final está vacía
                 df_resultado_final = pd.DataFrame(columns=df_obj.columns if df_obj is not None else [])
            elif not mascara_final_df_objetivo.any(): # Si la máscara final es todo False
                df_resultado_final = pd.DataFrame(columns=df_obj.columns if df_obj is not None else [])
            else: # Aplica la máscara final al DataFrame procesado
                 df_resultado_final = df_actual_procesando[mascara_final_df_objetivo].copy() 
            logger.debug(f"Resultado _procesar_busqueda_en_df_objetivo para '{termino_busqueda_original_para_este_df}': {len(df_resultado_final)} filas.")
            return df_resultado_final, None

    def _extraer_terminos_de_fila_completa(self, fila_df: pd.Series) -> Set[str]:
        terminos_extraidos_de_fila: Set[str] = set() # Conjunto para almacenar términos únicos de la fila
        if fila_df is None or fila_df.empty: return terminos_extraidos_de_fila # Si la fila está vacía, retorna
        for valor_celda in fila_df.values: # Itera sobre cada valor en la fila
            if pd.notna(valor_celda): # Si el valor no es NaN
                texto_celda_str = str(valor_celda).strip() # Convierte a string y limpia espacios
                if texto_celda_str: # Si no está vacío
                    texto_celda_norm = self._normalizar_para_busqueda(texto_celda_str) # Normaliza el texto de la celda
                    # Divide en palabras y filtra las que son solo dígitos o muy cortas (longitud 1)
                    palabras_significativas_celda = [palabra for palabra in texto_celda_norm.split() if len(palabra) > 1 and not palabra.isdigit()]
                    if palabras_significativas_celda: terminos_extraidos_de_fila.update(palabras_significativas_celda) # Añade palabras significativas al conjunto
                    # Si no hay palabras significativas pero el texto normalizado es útil (no es número, no es corto)
                    elif texto_celda_norm and len(texto_celda_norm) > 1 and not texto_celda_norm.isdigit() and self._parse_numero(texto_celda_norm) is None:
                        terminos_extraidos_de_fila.add(texto_celda_norm) # Añade el texto normalizado completo como un término
        return terminos_extraidos_de_fila

    def buscar(self, termino_busqueda_original: str, buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[List[int]], Optional[str]]:
        logger.info(f"Motor.buscar INICIO: termino='{termino_busqueda_original}', via_dicc={buscar_via_diccionario_flag}")
        # Define columnas de referencia para DataFrames vacíos, basado en descripciones si están cargadas
        columnas_descripcion_ref = self.datos_descripcion.columns if self.datos_descripcion is not None else []
        df_vacio_para_descripciones = pd.DataFrame(columns=columnas_descripcion_ref) # DataFrame vacío estándar para descripciones
        fcds_obtenidos_final_para_ui: Optional[pd.DataFrame] = None # DataFrame de FCDs para mostrar en UI
        indices_fcds_a_resaltar_en_preview: Optional[List[int]] = None # Índices de FCDs a resaltar

        # Manejo de término de búsqueda vacío
        if not termino_busqueda_original.strip():
            if self.datos_descripcion is not None: # Si hay descripciones cargadas, las devuelve todas
                logger.info("Término vacío. Devolviendo todas las descripciones.")
                return self.datos_descripcion.copy(), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
            else: # Si no hay descripciones cargadas
                logger.warning("Término vacío y descripciones no cargadas.")
                return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, "Descripciones no cargadas."

        # Flujo de búsqueda vía diccionario
        if buscar_via_diccionario_flag:
            if self.datos_diccionario is None: return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, None, "Diccionario no cargado."
            # Obtiene columnas para buscar en el diccionario
            columnas_dic_para_fcds, err_msg_cols_dic = self._obtener_nombres_columnas_busqueda_df(self.datos_diccionario, [], "diccionario_fcds_inicial")
            if not columnas_dic_para_fcds: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, None, err_msg_cols_dic

            # Parsea la query original para separar términos positivos globales de negativos globales
            _df_dummy, terminos_positivos_globales, terminos_negativos_globales = self._aplicar_negaciones_y_extraer_positivos(pd.DataFrame(), [], termino_busqueda_original)
            logger.info(f"Parseo global: Positivos='{terminos_positivos_globales}', Negativos Globales={terminos_negativos_globales}")

            # Lógica para búsqueda AND (cuando hay '+' en los términos positivos que no es parte de una frase exacta)
            if "+" in terminos_positivos_globales and not (terminos_positivos_globales.startswith('"') and terminos_positivos_globales.endswith('"')):
                logger.info(f"Detectada búsqueda AND en positivos globales: '{terminos_positivos_globales}'")
                partes_and = [p.strip() for p in terminos_positivos_globales.split("+") if p.strip()] # Divide la query AND
                df_resultado_acumulado_desc = self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame(columns=columnas_descripcion_ref)
                fcds_indices_acumulados = set() # Conjunto para almacenar índices de FCDs encontrados
                todas_partes_and_produjeron_terminos_validos = True # Flag para rastrear validez
                hay_error_en_busqueda_de_parte_o_desc = False # Flag para errores
                error_msg_critico_partes: Optional[str] = None # Mensaje de error

                if self.datos_descripcion is None: # Verifica que las descripciones estén cargadas
                     logger.error("Archivo de descripciones no cargado, no se puede proceder con búsqueda AND vía diccionario.")
                     return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas para búsqueda AND."
                columnas_desc_para_filtrado, err_cols_desc_fil = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                if not columnas_desc_para_filtrado: # Verifica configuración de columnas de descripción
                    return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_fil

                # Procesa cada parte de la consulta AND secuencialmente
                for i, parte_and_actual_str in enumerate(partes_and):
                    if not parte_and_actual_str: continue # Salta partes vacías
                    logger.debug(f"Procesando parte AND '{parte_and_actual_str}' (parte {i+1}/{len(partes_and)}) en diccionario...")
                    # Busca la parte actual en el diccionario para obtener FCDs
                    fcds_para_esta_parte, error_fcd_parte = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, columnas_dic_para_fcds, parte_and_actual_str, None)
                    if error_fcd_parte: # Si hay error en la búsqueda de esta parte
                        todas_partes_and_produjeron_terminos_validos = False; hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_fcd_parte
                        logger.warning(f"Parte AND '{parte_and_actual_str}' falló en diccionario con error: {error_fcd_parte}"); break
                    if fcds_para_esta_parte is None or fcds_para_esta_parte.empty: # Si no se encuentran FCDs para esta parte
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no encontró FCDs en diccionario."); break
                    
                    fcds_indices_acumulados.update(fcds_para_esta_parte.index.tolist()) # Acumula índices de FCDs
                    terminos_extraidos_de_esta_parte_set: Set[str] = set() # Extrae términos de los FCDs encontrados
                    for _, fila_fcd in fcds_para_esta_parte.iterrows(): terminos_extraidos_de_esta_parte_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                    
                    if not terminos_extraidos_de_esta_parte_set: # Si no se extrajeron términos válidos
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' encontró FCDs, pero no se extrajeron términos de ellas."); break
                    
                    # Construye una query OR con los términos extraídos para buscar en descripciones
                    terminos_or_con_comillas_actual = [f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_extraidos_de_esta_parte_set if t]
                    query_or_simple_actual = " | ".join(terminos_or_con_comillas_actual)
                    if not query_or_simple_actual: # Si la query OR está vacía
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no generó una query OR válida para descripciones."); break
                    
                    if df_resultado_acumulado_desc.empty and i >= 0: # Si ya no hay resultados acumulados en descripciones, el AND falla
                         logger.info(f"Resultados acumulados de descripción vacíos antes de aplicar filtro para '{parte_and_actual_str}'. Búsqueda AND final será vacía."); break
                    
                    logger.info(f"Aplicando filtro OR para '{parte_and_actual_str}' (Query: '{query_or_simple_actual[:100]}...') sobre {len(df_resultado_acumulado_desc)} filas de descripción.")
                    # Filtra los resultados acumulados de descripción con la query OR actual
                    df_resultado_acumulado_desc, error_sub_busqueda_desc = self._procesar_busqueda_en_df_objetivo(df_resultado_acumulado_desc, columnas_desc_para_filtrado, query_or_simple_actual, None) # Negativos globales se aplican al final de todo
                    if error_sub_busqueda_desc: # Si hay error en la sub-búsqueda en descripciones
                        hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_sub_busqueda_desc
                        logger.error(f"Error en sub-búsqueda OR para '{query_or_simple_actual}': {error_sub_busqueda_desc}"); break
                    if df_resultado_acumulado_desc.empty: # Si no hay resultados después del filtro
                        logger.info(f"Filtro OR para '{parte_and_actual_str}' no encontró coincidencias en resultados acumulados. Búsqueda AND final será vacía."); break
                
                # Prepara FCDs para mostrar en la UI
                if fcds_indices_acumulados and self.datos_diccionario is not None:
                    fcds_obtenidos_final_para_ui = self.datos_diccionario.loc[list(fcds_indices_acumulados)].drop_duplicates().copy()
                    indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                else: # Si no hay FCDs acumulados o diccionario no está cargado
                    fcds_obtenidos_final_para_ui = pd.DataFrame(columns=self.datos_diccionario.columns if self.datos_diccionario is not None else [])
                    indices_fcds_a_resaltar_en_preview = []
                
                if hay_error_en_busqueda_de_parte_o_desc: # Si hubo error en el proceso
                    return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_msg_critico_partes
                if not todas_partes_and_produjeron_terminos_validos or df_resultado_acumulado_desc.empty: # Si alguna parte no produjo términos o no hay resultados finales en descripciones
                    origen_fallo_and = OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS if not todas_partes_and_produjeron_terminos_validos else OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                    logger.info(f"Búsqueda AND '{terminos_positivos_globales}' no produjo resultados finales en descripciones (Origen: {origen_fallo_and.name}).")
                    return df_vacio_para_descripciones, origen_fallo_and, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                
                resultados_desc_final_filtrado_and = df_resultado_acumulado_desc # Resultados finales de descripción para el AND
                # Aplica negativos globales si existen y hay resultados
                if not resultados_desc_final_filtrado_and.empty and terminos_negativos_globales:
                    logger.info(f"Aplicando negativos globales {terminos_negativos_globales} a {len(resultados_desc_final_filtrado_and)} filas (resultado del AND de ORs)")
                    df_temp_neg, _, _ = self._aplicar_negaciones_y_extraer_positivos(
                        resultados_desc_final_filtrado_and, 
                        columnas_desc_para_filtrado, 
                        " ".join([f"#{neg}" for neg in terminos_negativos_globales]) # Construye una query solo de negativos
                    )
                    resultados_desc_final_filtrado_and = df_temp_neg # Actualiza los resultados
                
                logger.info(f"Búsqueda AND '{terminos_positivos_globales}' vía diccionario produjo {len(resultados_desc_final_filtrado_and)} resultados en descripciones.")
                return resultados_desc_final_filtrado_and, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
            else: # Flujo simple (no AND de alto nivel en positivos globales) o búsqueda puramente negativa
                origen_propuesto_flujo_simple: OrigenResultados = OrigenResultados.NINGUNO
                fcds_query_simple: Optional[pd.DataFrame] = None
                
                if terminos_positivos_globales.strip(): # Si hay términos positivos (flujo simple)
                    logger.info(f"BUSCAR EN DICC (FCDs) - Positivos (sin AND de alto nivel): Query='{terminos_positivos_globales}'")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
                    try:
                        # Negativos globales se aplican en descripciones, no en esta búsqueda de FCDs
                        fcds_temp, error_dic_pos = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, columnas_dic_para_fcds, terminos_positivos_globales, None)
                        if error_dic_pos: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_pos
                        fcds_query_simple = fcds_temp
                    except Exception as e_dic_pos:
                        logger.exception("Excepción búsqueda en diccionario (positivos simples)."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-positivos simples): {e_dic_pos}"
                elif terminos_negativos_globales: # Si la query original era puramente negativa (ej. "#termino1 #termino2")
                    logger.info(f"BUSCAR EN DICC (FCDs) - Puramente Negativo: Negs Globales={terminos_negativos_globales}")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC
                    try:
                        # Busca en el diccionario aplicando solo los negativos. Los FCDs resultantes NO deben contener los términos negados.
                        query_solo_negados_fcd = " ".join([f"#{neg}" for neg in terminos_negativos_globales])
                        fcds_temp, error_dic_neg = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, columnas_dic_para_fcds, query_solo_negados_fcd, None)
                        if error_dic_neg: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_neg
                        fcds_query_simple = fcds_temp
                    except Exception as e_dic_neg:
                        logger.exception("Excepción búsqueda en diccionario (puramente negativo)."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-negativo): {e_dic_neg}"
                else: # No hay términos positivos ni negativos (ya manejado por el chequeo de termino_busqueda_original vacío)
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None

                fcds_obtenidos_final_para_ui = fcds_query_simple # FCDs para mostrar en la UI
                if fcds_obtenidos_final_para_ui is not None and not fcds_obtenidos_final_para_ui.empty: # Si se encontraron FCDs
                    indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                    logger.info(f"FCDs obtenidas del diccionario (flujo simple/negativo): {len(fcds_obtenidos_final_para_ui)} filas.")
                else: # No se encontraron FCDs
                    logger.info(f"No se encontraron FCDs en diccionario para '{termino_busqueda_original}' (flujo simple/negativo).")
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                
                if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Descripciones no cargadas."
                
                terminos_para_buscar_en_descripcion_set: Set[str] = set() # Extrae términos de los FCDs para buscar en descripciones
                for _, fila_fcd in fcds_obtenidos_final_para_ui.iterrows(): terminos_para_buscar_en_descripcion_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                
                if not terminos_para_buscar_en_descripcion_set: # Si no se extrajeron términos válidos de los FCDs
                    logger.info("FCDs encontrados (flujo simple/negativo), pero no se extrajeron términos para descripciones.")
                    origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                    if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC: 
                        origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                    return df_vacio_para_descripciones, origen_final_sinterm, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None

                logger.info(f"Términos para desc ({len(terminos_para_buscar_en_descripcion_set)} únicos, muestra): {sorted(list(terminos_para_buscar_en_descripcion_set))[:10]}...")
                terminos_or_con_comillas_desc = [f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_para_buscar_en_descripcion_set if t]
                query_or_para_desc_simple = " | ".join(terminos_or_con_comillas_desc) # Construye query OR para descripciones
                
                if not query_or_para_desc_simple: # Si la query OR para descripciones está vacía
                    origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                    if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                        origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                    return df_vacio_para_descripciones, origen_q_vacia, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Query OR para descripciones vacía."

                columnas_desc_final_simple, err_cols_desc_final_simple = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                if not columnas_desc_final_simple: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, err_cols_desc_final_simple
                
                # Negativos globales se aplican aquí SOLO si la búsqueda original de FCDs NO era puramente negativa.
                # Si era puramente negativa, los FCDs ya están filtrados por esos negativos.
                negativos_a_aplicar_en_desc = terminos_negativos_globales if origen_propuesto_flujo_simple != OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC else []

                logger.info(f"BUSCAR EN DESC (vía FCD simple/negativa): Query='{query_or_para_desc_simple[:200]}...'. Neg. Adicionales a aplicar en Desc: {negativos_a_aplicar_en_desc}")
                try:
                    # Busca en descripciones usando la query OR y los negativos globales (si aplica)
                    resultados_desc_final_simple, error_busqueda_desc_simple = self._procesar_busqueda_en_df_objetivo(
                        self.datos_descripcion, 
                        columnas_desc_final_simple, 
                        query_or_para_desc_simple, 
                        terminos_negativos_adicionales=negativos_a_aplicar_en_desc
                    )
                    if error_busqueda_desc_simple: return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_busqueda_desc_simple
                    
                    if resultados_desc_final_simple is None or resultados_desc_final_simple.empty: # Si no hay resultados en descripciones
                        origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                        if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                            origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                        return df_vacio_para_descripciones, origen_res_desc_vacio_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                    else: # Si hay resultados
                        return resultados_desc_final_simple, origen_propuesto_flujo_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                except Exception as e_desc_proc_simple: # Captura cualquier otra excepción
                    logger.exception("Excepción búsqueda final en descripciones (flujo simple/negativo)."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, f"Error motor (desc final simple/negativo): {e_desc_proc_simple}"
        else: # Búsqueda directa en descripciones (no vía diccionario)
            if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas."
            columnas_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion")
            if not columnas_desc_directo: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_directo
            try:
                logger.info(f"BUSCAR EN DESC (DIRECTO): Query '{termino_busqueda_original}'")
                # Procesa la búsqueda directamente en descripciones. Los negativos de la query original se manejan dentro de _procesar_busqueda_en_df_objetivo.
                resultados_directos_desc, error_busqueda_desc_dir = self._procesar_busqueda_en_df_objetivo(self.datos_descripcion, columnas_desc_directo, termino_busqueda_original, None)
                if error_busqueda_desc_dir: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_busqueda_desc_dir
                if resultados_directos_desc is None or resultados_directos_desc.empty: return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
                else: return resultados_directos_desc, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None, None
            except Exception as e_desc_dir_proc: # Captura cualquier otra excepción
                logger.exception("Excepción búsqueda directa en descripciones."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (desc directa): {e_desc_dir_proc}"

# --- Interfaz Gráfica ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE_NAME = "config_buscador_avanzado_ui.json" 

    def __init__(self):
        super().__init__() # Llama al constructor de la clase padre tk.Tk
        self.title("Buscador Avanzado v1.10.1 (Parseo Numérico Mejorado)") # Título de la ventana
        self.geometry("1250x800") # Dimensiones iniciales de la ventana
        self.config: Dict[str, Any] = self._cargar_configuracion_app() # Carga la configuración de la aplicación
        # Obtiene los índices de columnas para la vista previa del diccionario desde la configuración
        indices_cfg_preview_dic = self.config.get("indices_columnas_busqueda_dic_preview", [])
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg_preview_dic) # Inicializa el motor de búsqueda
        self.resultados_actuales: Optional[pd.DataFrame] = None # Almacena los resultados de la búsqueda actual
        self.texto_busqueda_var = tk.StringVar(self) # Variable de Tkinter para el campo de entrada de búsqueda
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change) # Llama a _on_texto_busqueda_change cuando cambia el texto
        self.ultimo_termino_buscado: Optional[str] = None # Almacena el último término buscado
        self.reglas_guardadas: List[Dict[str, Any]] = [] # Lista para guardar reglas/búsquedas (funcionalidad no implementada)
        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None # DataFrame de FCDs de la última búsqueda
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None # DataFrame de descripciones finales de la última búsqueda
        self.indices_fcds_resaltados: Optional[List[int]] = None # Índices de FCDs a resaltar en la tabla
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO # Origen de los resultados actuales
        # Colores para las filas de las tablas y resaltados
        self.color_fila_par: str = "white"; self.color_fila_impar: str = "#f0f0f0"; self.color_resaltado_dic: str = "sky blue"
        self.op_buttons: Dict[str, ttk.Button] = {} # Diccionario para los botones de operadores
        self._configurar_estilo_ttk_app() # Configura el estilo de los widgets ttk
        self._crear_widgets_app() # Crea todos los widgets de la interfaz
        self._configurar_grid_layout_app() # Configura la disposición de los widgets en la cuadrícula
        self._configurar_eventos_globales_app() # Configura eventos globales (ej. Enter en búsqueda)
        self._configurar_tags_estilo_treeview_app() # Configura estilos para las tablas (Treeview)
        self._configurar_funcionalidad_orden_tabla(self.tabla_resultados) # Habilita ordenación en tabla de resultados
        self._configurar_funcionalidad_orden_tabla(self.tabla_diccionario) # Habilita ordenación en tabla de diccionario
        self._actualizar_mensaje_barra_estado("Listo. Cargue Diccionario y Descripciones.") # Mensaje inicial en barra de estado
        self._deshabilitar_botones_operadores() # Deshabilita botones de operadores inicialmente
        self._actualizar_estado_general_botones_y_controles() # Actualiza el estado de todos los botones y controles
        logger.info(f"Interfaz Gráfica (v1.10.1 Parseo Numérico Mejorado) inicializada.")

    def _try_except_wrapper(self, func, *args, **kwargs):
        # Envoltorio para manejar excepciones en funciones de la UI, mostrando un mensaje de error y registrando el traceback.
        try:
            return func(*args, **kwargs) # Ejecuta la función envuelta
        except Exception as e: # Si ocurre una excepción
            func_name = func.__name__; error_type = type(e).__name__; error_msg = str(e); tb_str = traceback.format_exc()
            logger.critical(f"Error en {func_name}: {error_type} - {error_msg}\n{tb_str}") # Registra el error crítico
            print(f"--- TRACEBACK COMPLETO (desde _try_except_wrapper para {func_name}) ---\n{tb_str}") # Imprime traceback en consola
            messagebox.showerror(f"Error Interno en {func_name}", f"Ocurrió un error inesperado:\n{error_type}: {error_msg}\n\nConsulte el log y la consola para el traceback completo.") # Muestra mensaje de error al usuario
            # Si el error ocurrió durante la carga de archivos, actualiza la UI para reflejar el estado de carga fallido
            if func_name in ["_cargar_diccionario_ui", "_cargar_excel_descripcion_ui"]: 
                self._actualizar_etiquetas_archivos_cargados()
                self._actualizar_estado_general_botones_y_controles()
            return None # Retorna None en caso de error

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str): 
        # Se llama cada vez que el texto en el campo de búsqueda cambia.
        self._actualizar_estado_botones_operadores() # Actualiza el estado (habilitado/deshabilitado) de los botones de operador.
    
    def _cargar_configuracion_app(self) -> Dict[str, Any]:
        # Carga la configuración de la aplicación desde un archivo JSON.
        config_cargada: Dict[str, Any] = {} # Diccionario para la configuración
        ruta_archivo_config = Path(self.CONFIG_FILE_NAME) # Ruta al archivo de configuración
        if ruta_archivo_config.exists(): # Si el archivo existe
            try:
                with ruta_archivo_config.open("r", encoding="utf-8") as f: config_cargada = json.load(f) # Carga el JSON
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE_NAME}")
            except Exception as e: logger.error(f"Error al cargar config '{self.CONFIG_FILE_NAME}': {e}") # Maneja error de carga
        else: logger.info(f"Archivo config '{self.CONFIG_FILE_NAME}' no encontrado.") # Si no existe, lo informa
        # Convierte las rutas guardadas como strings a objetos Path
        for clave_ruta in ["last_dic_path", "last_desc_path"]:
            valor_ruta = config_cargada.get(clave_ruta)
            config_cargada[clave_ruta] = Path(valor_ruta) if valor_ruta else None
        # Asegura que la clave para índices de preview del diccionario exista
        config_cargada.setdefault("indices_columnas_busqueda_dic_preview", [])
        return config_cargada

    def _guardar_configuracion_app(self):
        # Guarda la configuración actual de la aplicación en un archivo JSON.
        # Almacena las rutas de los últimos archivos cargados y la configuración de columnas de preview.
        self.config["last_dic_path"] = str(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else None
        self.config["last_desc_path"] = str(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else None
        self.config["indices_columnas_busqueda_dic_preview"] = self.motor.indices_columnas_busqueda_dic_preview
        try:
            with open(self.CONFIG_FILE_NAME, "w", encoding="utf-8") as f: json.dump(self.config, f, indent=4) # Guarda el JSON con indentación
            logger.info(f"Configuración guardada en: {self.CONFIG_FILE_NAME}")
        except Exception as e: logger.error(f"Error al guardar config '{self.CONFIG_FILE_NAME}': {e}") # Maneja error de guardado

    def _configurar_estilo_ttk_app(self):
        # Configura el tema de los widgets ttk para una apariencia más nativa o moderna.
        style = ttk.Style(self); os_name = platform.system() # Obtiene el estilo y el nombre del SO
        # Preferencias de tema según el SO
        prefs = {"Windows":["vista","xpnative"],"Darwin":["aqua"],"Linux":["clam","alt"]}
        # Selecciona el primer tema disponible de la lista de preferencias para el SO actual
        theme = next((t for t in prefs.get(os_name,["clam"]) if t in style.theme_names()), style.theme_use() or "default")
        try: 
            style.theme_use(theme) # Aplica el tema
            style.configure("Operator.TButton",padding=(2,1),font=("TkDefaultFont",9)) # Estilo específico para botones de operador
            logger.info(f"Tema TTK: {theme}")
        except: logger.warning(f"Fallo al aplicar tema {theme}") # Si falla, registra advertencia

    def _crear_widgets_app(self):
        # Crea todos los widgets principales de la interfaz gráfica.
        # Marco para los controles de carga de archivos y búsqueda
        self.marco_controles=ttk.LabelFrame(self,text="Controles")
        # Botón y etiqueta para cargar el archivo de diccionario
        self.btn_cargar_diccionario=ttk.Button(self.marco_controles,text="Cargar Diccionario",command=lambda: self._try_except_wrapper(self._cargar_diccionario_ui))
        self.lbl_dic_cargado=ttk.Label(self.marco_controles,text="Dic: Ninguno",width=25,anchor=tk.W,relief=tk.SUNKEN,borderwidth=1)
        # Botón y etiqueta para cargar el archivo de descripciones
        self.btn_cargar_descripciones=ttk.Button(self.marco_controles,text="Cargar Descripciones",command=lambda: self._try_except_wrapper(self._cargar_excel_descripcion_ui))
        self.lbl_desc_cargado=ttk.Label(self.marco_controles,text="Desc: Ninguno",width=25,anchor=tk.W,relief=tk.SUNKEN,borderwidth=1)
        
        # Frame para los botones de operadores de búsqueda
        self.frame_ops=ttk.Frame(self.marco_controles)
        op_buttons_defs = [("+","+"),("|","|"),("#","#"),("> ",">"),("< ","<"),("≥ ",">="),("≤ ","<="),("-","-")] # Definiciones de botones
        for i, (text, op_val_clean) in enumerate(op_buttons_defs): # Crea cada botón de operador
            btn = ttk.Button(self.frame_ops,text=text,command=lambda op=op_val_clean: self._insertar_operador_validado(op),style="Operator.TButton",width=3)
            btn.grid(row=0,column=i,padx=1,pady=1,sticky="nsew"); self.op_buttons[op_val_clean] = btn
        
        # Campo de entrada para el término de búsqueda
        self.entrada_busqueda=ttk.Entry(self.marco_controles,width=60,textvariable=self.texto_busqueda_var)
        # Botones de acción: Buscar, Salvar Regla (actualmente no funcional), Ayuda, Exportar
        self.btn_buscar=ttk.Button(self.marco_controles,text="Buscar",command=lambda: self._try_except_wrapper(self._ejecutar_busqueda_ui))
        self.btn_salvar_regla=ttk.Button(self.marco_controles,text="Salvar Regla",command=lambda: self._try_except_wrapper(self._salvar_regla_actual_ui),state="disabled")
        self.btn_ayuda=ttk.Button(self.marco_controles,text="?",command=self._mostrar_ayuda_ui,width=3)
        self.btn_exportar=ttk.Button(self.marco_controles,text="Exportar",command=lambda: self._try_except_wrapper(self._exportar_resultados_ui),state="disabled")
        
        # Etiqueta y frame para la tabla de vista previa del diccionario
        self.lbl_tabla_diccionario=ttk.Label(self,text="Vista Previa Diccionario:")
        self.frame_tabla_diccionario=ttk.Frame(self);self.tabla_diccionario=ttk.Treeview(self.frame_tabla_diccionario,show="headings",height=8);self.scrolly_diccionario=ttk.Scrollbar(self.frame_tabla_diccionario,orient="vertical",command=self.tabla_diccionario.yview);self.scrollx_diccionario=ttk.Scrollbar(self.frame_tabla_diccionario,orient="horizontal",command=self.tabla_diccionario.xview);self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set,xscrollcommand=self.scrollx_diccionario.set)
        # Etiqueta y frame para la tabla de resultados/descripciones
        self.lbl_tabla_resultados=ttk.Label(self,text="Resultados / Descripciones:");self.frame_tabla_resultados=ttk.Frame(self);self.tabla_resultados=ttk.Treeview(self.frame_tabla_resultados,show="headings");self.scrolly_resultados=ttk.Scrollbar(self.frame_tabla_resultados,orient="vertical",command=self.tabla_resultados.yview);self.scrollx_resultados=ttk.Scrollbar(self.frame_tabla_resultados,orient="horizontal",command=self.tabla_resultados.xview);self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set,xscrollcommand=self.scrollx_resultados.set)
        # Barra de estado en la parte inferior
        self.barra_estado=ttk.Label(self,text="Listo.",relief=tk.SUNKEN,anchor=tk.W,borderwidth=1);self._actualizar_etiquetas_archivos_cargados()

    def _configurar_grid_layout_app(self):
        # Configura la disposición de los widgets en la ventana principal usando el sistema grid.
        self.grid_rowconfigure(2,weight=1);self.grid_rowconfigure(4,weight=3);self.grid_columnconfigure(0,weight=1) # Configura pesos para expansión
        self.marco_controles.grid(row=0,column=0,sticky="new",padx=10,pady=(10,5)) # Marco de controles
        self.marco_controles.grid_columnconfigure(1,weight=1);self.marco_controles.grid_columnconfigure(3,weight=1) # Pesos dentro del marco
        self.btn_cargar_diccionario.grid(row=0,column=0,padx=(5,0),pady=5,sticky="w")
        self.lbl_dic_cargado.grid(row=0,column=1,padx=(2,10),pady=5,sticky="ew")
        self.btn_cargar_descripciones.grid(row=0,column=2,padx=(5,0),pady=5,sticky="w")
        self.lbl_desc_cargado.grid(row=0,column=3,padx=(2,5),pady=5,sticky="ew")
        self.frame_ops.grid(row=1,column=0,columnspan=6,padx=5,pady=(5,0),sticky="ew");[self.frame_ops.grid_columnconfigure(i,weight=1) for i in range(len(self.op_buttons))]
        self.entrada_busqueda.grid(row=2,column=0,columnspan=2,padx=5,pady=(0,5),sticky="ew")
        self.btn_buscar.grid(row=2,column=2,padx=(2,0),pady=(0,5),sticky="w")
        self.btn_salvar_regla.grid(row=2,column=3,padx=(2,0),pady=(0,5),sticky="w")
        self.btn_ayuda.grid(row=2,column=4,padx=(2,0),pady=(0,5),sticky="w")
        self.btn_exportar.grid(row=2,column=5,padx=(10,5),pady=(0,5),sticky="e")
        self.lbl_tabla_diccionario.grid(row=1,column=0,sticky="sw",padx=10,pady=(10,0)) # Etiqueta tabla diccionario
        self.frame_tabla_diccionario.grid(row=2,column=0,sticky="nsew",padx=10,pady=(0,10)) # Frame tabla diccionario
        self.frame_tabla_diccionario.grid_rowconfigure(0,weight=1);self.frame_tabla_diccionario.grid_columnconfigure(0,weight=1) # Pesos dentro del frame
        self.tabla_diccionario.grid(row=0,column=0,sticky="nsew");self.scrolly_diccionario.grid(row=0,column=1,sticky="ns");self.scrollx_diccionario.grid(row=1,column=0,sticky="ew") # Tabla y scrollbars
        self.lbl_tabla_resultados.grid(row=3,column=0,sticky="sw",padx=10,pady=(0,0)) # Etiqueta tabla resultados
        self.frame_tabla_resultados.grid(row=4,column=0,sticky="nsew",padx=10,pady=(0,10)) # Frame tabla resultados
        self.frame_tabla_resultados.grid_rowconfigure(0,weight=1);self.frame_tabla_resultados.grid_columnconfigure(0,weight=1) # Pesos dentro del frame
        self.tabla_resultados.grid(row=0,column=0,sticky="nsew");self.scrolly_resultados.grid(row=0,column=1,sticky="ns");self.scrollx_resultados.grid(row=1,column=0,sticky="ew") # Tabla y scrollbars
        self.barra_estado.grid(row=5,column=0,sticky="sew",padx=0,pady=(5,0)) # Barra de estado

    def _configurar_eventos_globales_app(self): 
        # Configura eventos globales: Enter en campo de búsqueda y acción al cerrar la ventana.
        self.entrada_busqueda.bind("<Return>",lambda e:self._try_except_wrapper(self._ejecutar_busqueda_ui)) # Enter para buscar
        self.protocol("WM_DELETE_WINDOW",self.on_closing_app) # Manejo del cierre de ventana

    def _actualizar_mensaje_barra_estado(self,m): 
        # Actualiza el mensaje en la barra de estado y lo registra en el log.
        self.barra_estado.config(text=m);logger.info(f"Mensaje UI (BarraEstado): {m}");self.update_idletasks()

    def _mostrar_ayuda_ui(self):
        # Muestra una ventana de ayuda con la sintaxis de búsqueda y el flujo.
        texto_ayuda = ("Sintaxis:\n- Texto: `router cisco`\n- AND: `tarjeta + 16 puertos`\n- OR: `modulo | SFP` (Nota: `/` ya no es OR)\n"
                       "- Numérico: `>1000W`, `<50V`, `>=48A`, `<=10.5W`\n- Rango: `10-20V`\n- Frase: `\"rack 19\"`\n- Negación: `#palabra` o `# \"frase\"`\n\n"
                       "Flujo Vía Diccionario:\n1. Query 'A+B': Parte 'A' y 'B' se buscan individualmente en Diccionario (FCDs).\n"
                       "2. Sinónimos: De las FCDs de 'A' se extraen Sinónimos_A. De las FCDs de 'B' se extraen Sinónimos_B.\n"
                       "3. Búsqueda en Descripciones: Se buscan filas que contengan (ALGÚN Sinónimo_A) Y (ALGÚN Sinónimo_B) mediante filtrado secuencial.\n"
                       "4. Negativos (#global): Se aplican al final sobre los resultados de descripciones.\n"
                       "5. Falla en Diccionario: Si 'A' o 'B' no da FCDs/sinónimos, se ofrece búsqueda directa de 'A+B' en Descripciones.")
        messagebox.showinfo("Ayuda - Sintaxis y Flujo", texto_ayuda) # Muestra el cuadro de diálogo de información

    def _configurar_tags_estilo_treeview_app(self):
        # Configura tags de estilo para las tablas (Treeview) para alternar colores de filas y resaltar.
        for tabla in [self.tabla_diccionario, self.tabla_resultados]: # Aplica a ambas tablas
            tabla.tag_configure("par", background=self.color_fila_par) # Tag para filas pares
            tabla.tag_configure("impar", background=self.color_fila_impar) # Tag para filas impares
        self.tabla_diccionario.tag_configure("resaltado_azul", background=self.color_resaltado_dic, foreground="black") # Tag para resaltar filas en diccionario

    def _configurar_funcionalidad_orden_tabla(self,tabla):
        # Configura la funcionalidad de ordenación de columnas para una tabla Treeview.
        cols = tabla["columns"] # Obtiene las columnas actuales de la tabla
        if cols: # Si hay columnas definidas
            # Asigna un comando a cada cabecera de columna para permitir la ordenación al hacer clic
            [tabla.heading(c,text=str(c),anchor=tk.W,command=lambda col=c,tbl=tabla: self._try_except_wrapper(self._ordenar_columna_tabla_ui,tbl,col,False)) for c in cols]

    def _ordenar_columna_tabla_ui(self,tabla,col,rev):
        # Ordena los datos de una tabla Treeview por la columna especificada.
        df_copia=None;idx_resaltar=None # Inicializa variables
        # Determina qué DataFrame usar según la tabla (diccionario o resultados)
        if tabla==self.tabla_diccionario and self.motor.datos_diccionario is not None:
            df_copia=self.motor.datos_diccionario.copy();idx_resaltar=self.indices_fcds_resaltados
        elif tabla==self.tabla_resultados and self.resultados_actuales is not None:
            df_copia=self.resultados_actuales.copy()
        else: # Si no hay datos o tabla no reconocida, revierte el comando de ordenación y retorna
            tabla.heading(col,command=lambda c=col,t=tabla:self._try_except_wrapper(self._ordenar_columna_tabla_ui,t,c,not rev));return
        
        if df_copia.empty or col not in df_copia.columns: # Si el DF está vacío o la columna no existe, retorna
            tabla.heading(col,command=lambda c=col,t=tabla:self._try_except_wrapper(self._ordenar_columna_tabla_ui,t,c,not rev));return
        
        df_num=pd.to_numeric(df_copia[col],errors='coerce') # Intenta convertir la columna a numérico
        # Ordena el DataFrame: numéricamente si es posible, sino alfabéticamente (ignorando may/min)
        df_ord=df_copia.sort_values(
            by=col,
            ascending=not rev, # Invierte el orden en cada clic
            na_position='last', # Coloca NaN al final
            key=(lambda x:pd.to_numeric(x,errors='coerce')) if not df_num.isna().all() else (lambda x:x.astype(str).str.lower())
        )
        
        columnas_para_diccionario_ordenado = None # Columnas específicas para la vista previa del diccionario
        if tabla==self.tabla_diccionario and self.motor.datos_diccionario is not None:
            # Obtiene los nombres de columna según la configuración de preview para mantener consistencia
            columnas_para_diccionario_ordenado, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                df_ord, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview"
            )
            if not columnas_para_diccionario_ordenado: columnas_para_diccionario_ordenado = list(df_ord.columns) # Fallback a todas las columnas
        
        # Actualiza la tabla correspondiente con los datos ordenados
        if tabla==self.tabla_diccionario:
            self._actualizar_tabla_treeview_ui(tabla,df_ord,limite_filas=None,columnas_a_mostrar=columnas_para_diccionario_ordenado, indices_a_resaltar=idx_resaltar)
        elif tabla==self.tabla_resultados:
            self.resultados_actuales=df_ord;self._actualizar_tabla_treeview_ui(tabla,self.resultados_actuales)
        
        # Actualiza el comando de la cabecera para el siguiente clic (invertir orden)
        tabla.heading(col,command=lambda c=col,t=tabla:self._try_except_wrapper(self._ordenar_columna_tabla_ui,t,c,not rev))
        self._actualizar_mensaje_barra_estado(f"Ordenado por '{col}'.")

    def _actualizar_tabla_treeview_ui(self,tabla,datos,limite_filas=None,columnas_a_mostrar=None,indices_a_resaltar=None):
        # Actualiza el contenido de una tabla Treeview con un DataFrame de pandas.
        is_dicc=tabla==self.tabla_diccionario; tabla_nombre = "Diccionario" if is_dicc else "Resultados" # Determina el nombre de la tabla
        [tabla.delete(i) for i in tabla.get_children()];tabla["columns"]=() # Limpia la tabla
        
        if datos is None or datos.empty: # Si no hay datos, configura la tabla vacía y retorna
            self._configurar_funcionalidad_orden_tabla(tabla); logger.debug(f"Tabla '{tabla_nombre}' vaciada."); return
        
        cols_orig=list(datos.columns); cols_para_usar_en_tabla: List[str] # Columnas originales del DataFrame
        if columnas_a_mostrar: # Si se especificaron columnas a mostrar (por índice o nombre)
            if all(isinstance(c, int) for c in columnas_a_mostrar): # Si son índices
                try: cols_para_usar_en_tabla = [cols_orig[i] for i in columnas_a_mostrar if 0 <= i < len(cols_orig)]
                except IndexError: logger.warning(f"Índices en columnas_a_mostrar fuera de rango para tabla '{tabla_nombre}'. Usando todas."); cols_para_usar_en_tabla = cols_orig
            elif all(isinstance(c, str) for c in columnas_a_mostrar): # Si son nombres de columna
                cols_para_usar_en_tabla = [c for c in columnas_a_mostrar if c in cols_orig]
            else: logger.warning(f"Tipo inesperado para columnas_a_mostrar en tabla '{tabla_nombre}'. Usando todas."); cols_para_usar_en_tabla = cols_orig
            if not cols_para_usar_en_tabla : logger.warning(f"columnas_a_mostrar no resultó en columnas válidas para tabla '{tabla_nombre}'. Usando todas."); cols_para_usar_en_tabla = cols_orig
        else: cols_para_usar_en_tabla = cols_orig # Si no se especificaron, usa todas las columnas originales
        
        if not cols_para_usar_en_tabla: # Si no hay columnas usables, configura tabla vacía y retorna
            self._configurar_funcionalidad_orden_tabla(tabla); logger.debug(f"Tabla '{tabla_nombre}' sin columnas usables."); return
        
        tabla["columns"]=tuple(cols_para_usar_en_tabla) # Establece las columnas de la tabla
        for c in cols_para_usar_en_tabla: # Configura cada columna (cabecera y ancho)
            tabla.heading(c,text=str(c),anchor=tk.W) # Texto y alineación de cabecera
            try: # Cálculo dinámico del ancho de columna
                if c in datos.columns: ancho_contenido = datos[c].astype(str).str.len().quantile(0.95) if not datos[c].empty else 0
                else: ancho_contenido = 0 
                ancho_cabecera = len(str(c)); ancho = max(70, min(int(max(ancho_cabecera * 7, ancho_contenido * 5.5) + 15), 350))
            except Exception as e_ancho: logger.warning(f"Error calculando ancho para columna '{c}' en tabla '{tabla_nombre}': {e_ancho}"); ancho = 100 # Ancho por defecto en caso de error
            tabla.column(c,anchor=tk.W,width=ancho,minwidth=50) # Aplica ancho y ancho mínimo
        
        df_iterar = datos[cols_para_usar_en_tabla]; num_filas_original=len(df_iterar) # DataFrame a iterar
        # Lógica para limitar filas mostradas, excepto si hay resaltados específicos
        mostrar_todo_por_resaltado = is_dicc and indices_a_resaltar and num_filas_original > 0
        if not mostrar_todo_por_resaltado and limite_filas and num_filas_original > limite_filas: df_iterar=df_iterar.head(limite_filas)
        elif mostrar_todo_por_resaltado: logger.debug(f"Mostrando todas {num_filas_original} filas de '{tabla_nombre}' por resaltado.")
        
        for i,(idx,row) in enumerate(df_iterar.iterrows()): # Itera sobre filas del DataFrame a mostrar
            vals=[str(v) if pd.notna(v) else "" for v in row.values];tags=["par" if i%2==0 else "impar"] # Prepara valores y tags de estilo
            if is_dicc and indices_a_resaltar and idx in indices_a_resaltar:tags.append("resaltado_azul") # Aplica tag de resaltado si es necesario
            try: tabla.insert("","end",values=vals,tags=tuple(tags),iid=f"row_{idx}") # Inserta la fila en la tabla
            except Exception as e_ins: logger.warning(f"Error insertando fila {idx} en '{tabla_nombre}': {e_ins}") # Maneja error de inserción
        
        self._configurar_funcionalidad_orden_tabla(tabla); logger.debug(f"Tabla '{tabla_nombre}' actualizada con {len(tabla.get_children())} filas visibles.")

    def _actualizar_etiquetas_archivos_cargados(self):
        # Actualiza las etiquetas que muestran los nombres de los archivos cargados.
        max_l=25;dic_p=self.motor.archivo_diccionario_actual;desc_p=self.motor.archivo_descripcion_actual # Longitud máxima y rutas
        dic_n=dic_p.name if dic_p else "Ninguno";desc_n=desc_p.name if desc_p else "Ninguno" # Nombres de archivo
        # Acorta nombres largos para que quepan en la etiqueta
        dic_d=f"Dic: {dic_n}" if len(dic_n)<=max_l else f"Dic: ...{dic_n[-(max_l-4):]}";
        desc_d=f"Desc: {desc_n}" if len(desc_n)<=max_l else f"Desc: ...{desc_n[-(max_l-4):]}"
        # Configura texto y color de etiquetas
        self.lbl_dic_cargado.config(text=dic_d,foreground="green" if dic_p else "red")
        self.lbl_desc_cargado.config(text=desc_d,foreground="green" if desc_p else "red")

    def _actualizar_estado_general_botones_y_controles(self):
        # Actualiza el estado (habilitado/deshabilitado) de varios botones y controles según el estado de la aplicación.
        dic_ok=self.motor.datos_diccionario is not None;desc_ok=self.motor.datos_descripcion is not None # Verifica si los archivos están cargados
        if dic_ok or desc_ok: self._actualizar_estado_botones_operadores() # Habilita/deshabilita operadores si hay algún archivo
        else: self._deshabilitar_botones_operadores() # Deshabilita operadores si no hay archivos
        
        self.btn_buscar["state"]="normal" if dic_ok and desc_ok else "disabled" # Botón Buscar: habilitado si ambos archivos están cargados
        salvar_ok=False # Flag para habilitar el botón de salvar regla
        if self.ultimo_termino_buscado and self.origen_principal_resultados!=OrigenResultados.NINGUNO: # Si hay una búsqueda previa
            # Lógica para determinar si se puede salvar la regla (basado en el tipo de resultado)
            if self.origen_principal_resultados.es_via_diccionario and \
               ((self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty) or \
                (self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and \
                 self.origen_principal_resultados in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC] )):
                salvar_ok=True
            elif (self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA) and \
                 self.desc_finales_de_ultima_busqueda is not None:
                salvar_ok=True
        self.btn_salvar_regla["state"]="normal" if salvar_ok else "disabled" # Habilita/deshabilita botón Salvar Regla
        # Botón Exportar: habilitado si hay resultados actuales y no están vacíos
        self.btn_exportar["state"]="normal" if (self.resultados_actuales is not None and not self.resultados_actuales.empty) else "disabled"

    def _cargar_diccionario_ui(self):
        # Maneja la acción de cargar el archivo de diccionario desde la UI.
        cfg_path=self.config.get("last_dic_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada=filedialog.askopenfilename(title="Cargar Diccionario",filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada: return # Si el usuario cancela, retorna
        
        nombre_archivo = Path(ruta_seleccionada).name # Nombre del archivo seleccionado
        self._actualizar_mensaje_barra_estado(f"Cargando dicc: {nombre_archivo}...") # Actualiza barra de estado
        # Resetea tablas y estado de búsqueda
        self._actualizar_tabla_treeview_ui(self.tabla_diccionario,None);self._actualizar_tabla_treeview_ui(self.tabla_resultados,None)
        self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None
        self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        
        ok,msg=self.motor.cargar_excel_diccionario(ruta_seleccionada) # Llama al motor para cargar el diccionario
        desc_n_title=Path(self.motor.archivo_descripcion_actual).name if self.motor.archivo_descripcion_actual else "N/A" # Nombre del archivo de descripciones para el título
        
        if ok and self.motor.datos_diccionario is not None: # Si la carga fue exitosa
            self.config["last_dic_path"]=Path(ruta_seleccionada);self._guardar_configuracion_app() # Guarda la ruta en config
            df_d=self.motor.datos_diccionario;n_filas=len(df_d) # DataFrame y número de filas
            # Obtiene columnas para la vista previa
            cols_prev,_=self.motor._obtener_nombres_columnas_busqueda_df(df_d,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
            self.lbl_tabla_diccionario.config(text=f"Diccionario ({n_filas} filas)") # Actualiza etiqueta de tabla
            self._actualizar_tabla_treeview_ui(self.tabla_diccionario,df_d,limite_filas=100,columnas_a_mostrar=cols_prev) # Actualiza tabla
            self.title(f"Buscador - Dic: {nombre_archivo} | Desc: {desc_n_title}") # Actualiza título de ventana
            self._actualizar_mensaje_barra_estado(f"Diccionario '{nombre_archivo}' ({n_filas}) cargado.")
        else: # Si la carga falló
            self._actualizar_mensaje_barra_estado(f"Error cargando diccionario: {msg or 'Desconocido'}");messagebox.showerror("Error Carga Dicc",msg or "Error desconocido")
            self.title(f"Buscador - Dic: N/A (Error) | Desc: {desc_n_title}")
        self._actualizar_etiquetas_archivos_cargados();self._actualizar_estado_general_botones_y_controles() # Actualiza UI

    def _cargar_excel_descripcion_ui(self):
        # Maneja la acción de cargar el archivo de descripciones desde la UI.
        cfg_path=self.config.get("last_desc_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada_str=filedialog.askopenfilename(title="Cargar Descripciones",filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada_str: logger.info("Carga de descripciones cancelada."); return # Si el usuario cancela
        
        nombre_archivo = Path(ruta_seleccionada_str).name; # Nombre del archivo
        self._actualizar_mensaje_barra_estado(f"Cargando descripciones: {nombre_archivo}...") # Actualiza barra de estado
        # Resetea tabla de resultados y estado de búsqueda
        self.resultados_actuales=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO
        self._actualizar_tabla_treeview_ui(self.tabla_resultados,None)
        
        ok, msg_error = self.motor.cargar_excel_descripcion(ruta_seleccionada_str) # Llama al motor para cargar descripciones
        dic_n_title=Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A" # Nombre de archivo de diccionario para título
        
        if ok and self.motor.datos_descripcion is not None: # Si la carga fue exitosa
            self.config["last_desc_path"] = Path(ruta_seleccionada_str); self._guardar_configuracion_app() # Guarda ruta en config
            df_desc = self.motor.datos_descripcion; num_filas = len(df_desc) # DataFrame y número de filas
            self._actualizar_mensaje_barra_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando vista previa...")
            self._actualizar_tabla_treeview_ui(self.tabla_resultados, df_desc, limite_filas=200) # Actualiza tabla de resultados (vista previa)
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: {nombre_archivo}") # Actualiza título de ventana
        else: # Si la carga falló
            error_a_mostrar = msg_error or "Ocurrió un error desconocido al cargar el archivo de descripciones."
            self._actualizar_mensaje_barra_estado(f"Error cargando descripciones: {error_a_mostrar}"); messagebox.showerror("Error al Cargar Archivo de Descripciones", error_a_mostrar)
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: N/A (Error)")
        self._actualizar_etiquetas_archivos_cargados();self._actualizar_estado_general_botones_y_controles() # Actualiza UI

    def _ejecutar_busqueda_ui(self):
        # Ejecuta la búsqueda principal cuando el usuario presiona "Buscar" o Enter.
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None: # Verifica si los archivos están cargados
            messagebox.showwarning("Archivos Faltantes","Cargue Diccionario y Descripciones.");return
        
        term_ui=self.texto_busqueda_var.get();self.ultimo_termino_buscado=term_ui # Obtiene el término de búsqueda
        # Resetea el estado de los resultados previos
        self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None
        self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        self._actualizar_tabla_treeview_ui(self.tabla_resultados,None);self._actualizar_mensaje_barra_estado(f"Buscando '{term_ui}'...")
        
        # Ejecuta la búsqueda a través del motor (inicialmente vía diccionario)
        res_df,origen,fcds,idx_res,err_msg = self.motor.buscar(termino_busqueda_original=term_ui, buscar_via_diccionario_flag=True)
        # Almacena los resultados y el origen
        self.fcds_de_ultima_busqueda=fcds;self.origen_principal_resultados=origen;self.indices_fcds_resaltados=idx_res
        df_desc_cols=self.motor.datos_descripcion.columns if self.motor.datos_descripcion is not None else [] # Columnas de referencia para DF vacío
        
        # Actualiza la tabla del diccionario (FCDs)
        if self.motor.datos_diccionario is not None:
            num_fcds_actual=len(self.indices_fcds_resaltados) if self.indices_fcds_resaltados else 0
            dicc_lbl=f"Diccionario ({len(self.motor.datos_diccionario)} filas)" + \
                     (f" - {num_fcds_actual} FCDs resaltados" if num_fcds_actual>0 and origen.es_via_diccionario and origen!=OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS else "")
            self.lbl_tabla_diccionario.config(text=dicc_lbl)
            cols_prev_dic_actual,_=self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
            # Muestra todas las filas si hay resaltados, sino limita a 100
            limite_filas_dic_preview = None if self.indices_fcds_resaltados and num_fcds_actual > 0 else 100 
            self._actualizar_tabla_treeview_ui(self.tabla_diccionario,self.motor.datos_diccionario,limite_filas=limite_filas_dic_preview,columnas_a_mostrar=cols_prev_dic_actual,indices_a_resaltar=self.indices_fcds_resaltados)
        
        # Manejo de diferentes orígenes de resultados y errores
        if err_msg and origen.es_error_operacional:messagebox.showerror("Error Motor",f"Error interno: {err_msg}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        elif origen.es_error_carga or origen.es_error_configuracion or origen.es_termino_invalido:messagebox.showerror("Error Búsqueda",err_msg or f"Error: {origen.name}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        elif origen in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC]:
            self.resultados_actuales=res_df;self._actualizar_mensaje_barra_estado(f"'{term_ui}': {len(fcds) if fcds is not None else 0} en Dic, {len(res_df) if res_df is not None else 0} en Desc.")
        elif origen==OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS: # Si no se encontró en diccionario
            self.resultados_actuales=res_df ;self._actualizar_mensaje_barra_estado(f"'{term_ui}': No en Diccionario.");
            # Pregunta al usuario si desea buscar directamente en descripciones
            if messagebox.askyesno("Búsqueda Alternativa",f"'{term_ui}' no encontrado en Diccionario.\n\n¿Buscar '{term_ui}' directamente en Descripciones?"):
                self._try_except_wrapper(self._buscar_directo_en_descripciones_y_actualizar_ui, term_ui, df_desc_cols)
            else: self._actualizar_estado_general_botones_y_controles() # Actualiza UI si no
        elif origen in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC]:
            # Si se encontraron FCDs pero no resultados en descripciones o términos válidos
            self.resultados_actuales=res_df;num_fcds_i=len(fcds) if fcds is not None else 0;msg_fcd_i=f"{num_fcds_i} en Diccionario"
            msg_desc_i="pero no se extrajeron términos válidos para Desc." if origen in [OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC] else "pero 0 resultados en Desc."
            self._actualizar_mensaje_barra_estado(f"'{term_ui}': {msg_fcd_i}, {msg_desc_i.replace('.','')} en Desc.")
            # Pregunta al usuario si desea buscar directamente en descripciones
            if messagebox.askyesno("Búsqueda Alternativa",f"{msg_fcd_i} para '{term_ui}', {msg_desc_i}\n\n¿Buscar '{term_ui}' directamente en Descripciones?"):
                self._try_except_wrapper(self._buscar_directo_en_descripciones_y_actualizar_ui, term_ui, df_desc_cols)
            else: self._actualizar_estado_general_botones_y_controles() # Actualiza UI si no
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS: # Búsqueda directa con resultados
            self.resultados_actuales=res_df;self._actualizar_mensaje_barra_estado(f"Búsqueda directa '{term_ui}': {len(res_df) if res_df is not None else 0} resultados.")
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_VACIA: # Búsqueda directa sin resultados (o término vacío)
            self.resultados_actuales=res_df;num_r=len(res_df) if res_df is not None else 0
            self._actualizar_mensaje_barra_estado(f"Mostrando todas las desc ({num_r})." if not term_ui.strip() else f"Búsqueda directa '{term_ui}': 0 resultados.")
            if term_ui.strip() and num_r==0 :messagebox.showinfo("Info",f"No resultados para '{term_ui}' en búsqueda directa.")
        
        if self.resultados_actuales is None:self.resultados_actuales=pd.DataFrame(columns=df_desc_cols) # Asegura que resultados_actuales no sea None
        self.desc_finales_de_ultima_busqueda=self.resultados_actuales.copy(); # Guarda una copia para salvar reglas
        self._actualizar_tabla_treeview_ui(self.tabla_resultados,self.resultados_actuales); # Actualiza la tabla de resultados
        self._actualizar_estado_general_botones_y_controles() # Actualiza estado de botones

    def _buscar_directo_en_descripciones_y_actualizar_ui(self, term_ui_original: str, columnas_df_desc_referencia: List[str]):
        # Realiza una búsqueda directa en descripciones y actualiza la UI.
        self._actualizar_mensaje_barra_estado(f"Iniciando búsqueda directa de '{term_ui_original}' en descripciones...")
        self.indices_fcds_resaltados = None # No hay FCDs resaltados en búsqueda directa
        # Limpia el resaltado de la tabla de diccionario y la recarga
        if self.motor.datos_diccionario is not None:
            cols_prev_dic_alt,_ = self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview")
            self.lbl_tabla_diccionario.config(text=f"Vista Previa Diccionario ({len(self.motor.datos_diccionario)} filas)")
            self._actualizar_tabla_treeview_ui(self.tabla_diccionario, self.motor.datos_diccionario, limite_filas=100, columnas_a_mostrar=cols_prev_dic_alt, indices_a_resaltar=None)
        
        # Ejecuta la búsqueda directa en el motor
        res_df_dir, orig_dir, _, _, msg_error_directo = self.motor.buscar(termino_busqueda_original=term_ui_original, buscar_via_diccionario_flag=False)
        self.origen_principal_resultados = orig_dir; self.fcds_de_ultima_busqueda = None # Actualiza estado de la búsqueda
        
        if msg_error_directo and (orig_dir.es_error_operacional or orig_dir.es_termino_invalido): # Si hay error
            messagebox.showerror("Error Búsqueda Directa", f"Error: {msg_error_directo}"); self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_referencia)
        else: self.resultados_actuales = res_df_dir # Asigna resultados
        
        num_rdd = len(self.resultados_actuales) if self.resultados_actuales is not None else 0 # Número de resultados
        self._actualizar_mensaje_barra_estado(f"Búsqueda directa de '{term_ui_original}': {num_rdd} resultados.")
        if num_rdd == 0 and orig_dir == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and term_ui_original.strip(): # Informa si no hay resultados
            messagebox.showinfo("Info", f"No resultados para '{term_ui_original}' en búsqueda directa.")
        
        if self.resultados_actuales is None: self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_referencia) # Asegura que no sea None
        self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy() # Guarda copia
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, self.resultados_actuales) # Actualiza tabla
        self._actualizar_estado_general_botones_y_controles() # Actualiza UI

    def _salvar_regla_actual_ui(self):
        # Guarda metadatos de la búsqueda actual (funcionalidad de registro, no de re-ejecución).
        origen_nombre = self.origen_principal_resultados.name
        # Verifica si hay algo que salvar
        if not self.ultimo_termino_buscado and not (self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and self.desc_finales_de_ultima_busqueda is not None): 
            messagebox.showerror("Error Salvar", "No hay búsqueda para salvar."); return
        
        df_salvar: Optional[pd.DataFrame] = None; tipo_datos = "DESCONOCIDO" # DataFrame a considerar y tipo de datos
        # Lógica para determinar qué datos se asocian con la regla guardada
        if self.origen_principal_resultados.es_via_diccionario:
            if self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty: 
                df_salvar = self.desc_finales_de_ultima_busqueda; tipo_datos = "DESC_VIA_DICC"
            elif self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty: 
                df_salvar = self.fcds_de_ultima_busqueda; tipo_datos = "FCDS_DICC"
        elif self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            if self.desc_finales_de_ultima_busqueda is not None: 
                df_salvar = self.desc_finales_de_ultima_busqueda; tipo_datos = "DESC_DIRECTA";
            if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and not (self.ultimo_termino_buscado or "").strip(): 
                tipo_datos = "TODAS_DESC" # Caso especial: se mostraron todas las descripciones por término vacío
        
        if df_salvar is not None: # Si se identificaron datos para salvar
            regla = {"termino": self.ultimo_termino_buscado or "N/A", "origen": origen_nombre, "tipo": tipo_datos, "filas": len(df_salvar), "ts": pd.Timestamp.now().isoformat()}
            self.reglas_guardadas.append(regla); self._actualizar_mensaje_barra_estado(f"Búsqueda '{self.ultimo_termino_buscado}' registrada."); messagebox.showinfo("Regla Salvada", f"Metadatos de '{self.ultimo_termino_buscado}' guardados.")
            logger.info(f"Regla guardada: {regla}")
        else: messagebox.showwarning("Nada que Salvar", "No hay datos claros para salvar.")
        self._actualizar_estado_general_botones_y_controles() # Actualiza UI

    def _exportar_resultados_ui(self):
        # Exporta los resultados actuales (tabla de descripciones) a un archivo Excel o CSV.
        if self.resultados_actuales is None or self.resultados_actuales.empty: 
            messagebox.showinfo("Exportar", "No hay resultados para exportar."); return
        
        nombre_archivo_sugerido = f"resultados_{pd.Timestamp.now():%Y%m%d_%H%M%S}" # Nombre de archivo sugerido con timestamp
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")], title="Guardar resultados", initialfile=nombre_archivo_sugerido)
        if not ruta: return # Si el usuario cancela, retorna
        
        try: # Intenta guardar el archivo
            if ruta.endswith(".xlsx"): self.resultados_actuales.to_excel(ruta, index=False)
            elif ruta.endswith(".csv"): self.resultados_actuales.to_csv(ruta, index=False, encoding='utf-8-sig') # UTF-8 con BOM para compatibilidad Excel
            else: messagebox.showerror("Error Formato", "Usar .xlsx o .csv."); return # Formato no soportado
            messagebox.showinfo("Exportado", f"Resultados exportados a:\n{ruta}"); self._actualizar_mensaje_barra_estado(f"Exportado a {Path(ruta).name}")
            logger.info(f"Resultados exportados a {ruta}")
        except Exception as e_export: # Maneja errores de exportación
            logger.error(f"Error al exportar resultados a '{ruta}': {e_export}")
            messagebox.showerror("Error Exportación", f"No se pudo exportar:\n{e_export}")

    def _actualizar_estado_botones_operadores(self):
        # Habilita o deshabilita los botones de operadores (+, |, #, etc.) según el contexto del campo de búsqueda.
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None: 
            self._deshabilitar_botones_operadores(); return # Si no hay datos cargados, deshabilita todos
        
        [btn.config(state="normal") for btn in self.op_buttons.values()] # Habilita todos por defecto
        txt=self.texto_busqueda_var.get();cur_pos=self.entrada_busqueda.index(tk.INSERT) # Texto actual y posición del cursor
        last_char_rel=txt[:cur_pos].strip()[-1:] if txt[:cur_pos].strip() else "" # Último carácter relevante antes del cursor
        
        ops_logicos=["+","|","/"]; ops_comp_pref=[">","<"]; # Define tipos de operadores
        
        # Lógica para deshabilitar operadores según el último carácter
        if not last_char_rel or last_char_rel in ops_logicos + ["#","<",">","=","-"]: # Si es inicio o después de operador lógico/negación/comparación
            if self.op_buttons.get("+"): self.op_buttons["+"]["state"]="disabled"
            if self.op_buttons.get("|"): self.op_buttons["|"]["state"]="disabled"
            # if self.op_buttons.get("/"): self.op_buttons["/"]["state"]="disabled" # '/' ya no es OR
        
        if last_char_rel and last_char_rel not in ops_logicos + [" "]: # Si el último carácter no es operador lógico ni espacio
             if self.op_buttons.get("#"): self.op_buttons["#"]["state"]="disabled" # Deshabilita negación

        if last_char_rel in [">","<","="]: # Si el último es un operador de comparación
            for opk in ops_comp_pref + ["=","-"]: # Deshabilita otros operadores de comparación y rango
                if self.op_buttons.get(opk): self.op_buttons[opk]["state"]="disabled"
            if last_char_rel == ">" and self.op_buttons.get(">="): self.op_buttons[">="]["state"]="disabled"
            if last_char_rel == "<" and self.op_buttons.get("<="): self.op_buttons["<="]["state"]="disabled"
        
        if last_char_rel.isdigit(): # Si el último es un dígito
            for opk_pref in ops_comp_pref + ["=","#"]: # Deshabilita operadores de prefijo numérico y negación
                 if self.op_buttons.get(opk_pref): self.op_buttons[opk_pref]["state"]="disabled"
        elif not last_char_rel or last_char_rel in [" ","+","|","/"]: # Si es inicio, espacio u operador lógico
            if self.op_buttons.get("-"): self.op_buttons["-"]["state"]="disabled" # Deshabilita el guion de rango

    def _insertar_operador_validado(self,op_limpio: str):
        # Inserta un operador en el campo de búsqueda con los espacios adecuados.
        ops_con_espacio_alrededor = ["+", "|", "/"] 
        texto_a_insertar: str
        if op_limpio in ops_con_espacio_alrededor: texto_a_insertar = f" {op_limpio} " # Operadores lógicos con espacios
        elif op_limpio == "-": texto_a_insertar = f"{op_limpio}" # Guion de rango, sin espacios forzados
        elif op_limpio in [">=", "<="]: texto_a_insertar = f"{op_limpio}" # Operadores de comparación de dos caracteres
        elif op_limpio in [">", "<", "="]: texto_a_insertar = f"{op_limpio}" # Operadores de comparación de un caracter
        elif op_limpio == "#": texto_a_insertar = f"{op_limpio} " # Negación con espacio después
        else: texto_a_insertar = op_limpio # Caso por defecto
        
        self.entrada_busqueda.insert(tk.INSERT,texto_a_insertar);self.entrada_busqueda.focus_set() # Inserta y pone foco
        self._actualizar_estado_botones_operadores() # Actualiza estado de botones

    def _deshabilitar_botones_operadores(self): 
        # Deshabilita todos los botones de operadores.
        [btn.config(state="disabled") for btn in self.op_buttons.values()]

    def on_closing_app(self):
        # Maneja el evento de cierre de la ventana de la aplicación.
        try:
            logger.info("Cerrando aplicación Buscador Avanzado...")
            self._guardar_configuracion_app() # Guarda la configuración antes de cerrar
            self.destroy() # Cierra la ventana principal de Tkinter
        except Exception as e: # Maneja cualquier error durante el cierre
            func_name = "on_closing_app"; error_type = type(e).__name__; error_msg = str(e); tb_str = traceback.format_exc()
            logger.critical(f"Error en {func_name}: {error_type} - {error_msg}\n{tb_str}")
            print(f"--- TRACEBACK COMPLETO (desde {func_name}) ---\n{tb_str}")
            self.destroy() # Intenta cerrar de todas formas

# --- Punto de Entrada Principal de la Aplicación ---
if __name__ == "__main__":
    LOG_FILE_NAME = "Buscador_Avanzado_App_v1.10.1.log" # Nombre del archivo de log, versión actualizada
    # Configuración básica del logging para guardar en archivo y mostrar en consola
    logging.basicConfig(
        level=logging.DEBUG, # Nivel mínimo de mensajes a registrar
        format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s", # Formato del mensaje de log
        handlers=[
            logging.FileHandler(LOG_FILE_NAME, encoding="utf-8", mode="w"), # Handler para escribir en archivo (sobrescribe en cada ejecución)
            logging.StreamHandler() # Handler para mostrar en consola
        ])
    root_logger = logging.getLogger() # Obtiene el logger raíz
    # Mensaje de inicio de la aplicación
    root_logger.info(f"--- Iniciando Buscador Avanzado v1.10.1 (Parseo Numérico Mejorado) (Script: {Path(__file__).name}) ---")
    root_logger.info(f"Logs siendo guardados en: {Path(LOG_FILE_NAME).resolve()}")

    # Verificación de dependencias
    dependencias_faltantes_main: List[str] = []
    try: import pandas as pd_check_main; root_logger.info(f"Pandas: {pd_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("pandas")
    try: import openpyxl as opxl_check_main; root_logger.info(f"openpyxl: {opxl_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("openpyxl") # Necesario para .xlsx
    try: import numpy as np_check_main; root_logger.info(f"Numpy: {np_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("numpy")
    try: import xlrd as xlrd_check_main; root_logger.info(f"xlrd: {xlrd_check_main.__version__}") # Necesario para .xls antiguos
    except ImportError: root_logger.warning("xlrd no encontrado. Carga de .xls antiguos podría fallar.")

    # Si faltan dependencias críticas, muestra error y sale
    if dependencias_faltantes_main:
        mensaje_error_deps_main = (f"Faltan dependencias críticas: {', '.join(dependencias_faltantes_main)}.\nInstale con: pip install {' '.join(dependencias_faltantes_main)}")
        root_logger.critical(mensaje_error_deps_main)
        try: # Intenta mostrar el error en una ventana de Tkinter si es posible
            root_error_tk_main = tk.Tk(); root_error_tk_main.withdraw() # Crea ventana oculta para messagebox
            messagebox.showerror("Dependencias Faltantes", mensaje_error_deps_main); root_error_tk_main.destroy()
        except Exception as e_tk_dep_main: print(f"ERROR CRITICO (Error al mostrar mensaje Tkinter: {e_tk_dep_main}): {mensaje_error_deps_main}") # Fallback a print si Tkinter falla
        exit(1) # Termina la ejecución
    
    # Inicia la aplicación
    try: 
        app=InterfazGrafica() # Crea la instancia de la interfaz gráfica
        app.mainloop() # Inicia el bucle principal de Tkinter
    except Exception as e_main_app_exc: # Captura errores fatales no controlados
        root_logger.critical("Error fatal no controlado en la aplicación principal:", exc_info=True)
        tb_str_fatal = traceback.format_exc()
        print(f"--- TRACEBACK FATAL (desde bloque __main__) ---\n{tb_str_fatal}") # Imprime traceback
        try: # Intenta mostrar mensaje de error fatal
            root_fatal_tk_main = tk.Tk(); root_fatal_tk_main.withdraw()
            messagebox.showerror("Error Fatal Inesperado", f"Error crítico: {e_main_app_exc}\nConsulte '{LOG_FILE_NAME}' y la consola."); root_fatal_tk_main.destroy()
        except: print(f"ERROR FATAL: {e_main_app_exc}. Revise '{LOG_FILE_NAME}'.") # Fallback si Tkinter falla
    finally: 
        root_logger.info(f"--- Finalizando Buscador ---") # Mensaje de finalización
