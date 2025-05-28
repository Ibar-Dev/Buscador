# -*- coding: utf-8 -*-
# Se especifica la codificación UTF-8 para asegurar la correcta interpretación de caracteres especiales.

# Importaciones de la biblioteca estándar y de terceros
import re  # Módulo para trabajar con expresiones regulares, fundamental para el análisis de texto.
import tkinter as tk  # Biblioteca para la creación de interfaces gráficas de usuario (GUI).
from tkinter import ttk  # Módulo de Tkinter que provee widgets temáticos (mejorados).
from tkinter import messagebox  # Para mostrar cuadros de diálogo estándar (información, error, advertencia).
from tkinter import filedialog  # Para mostrar diálogos de selección de archivos y directorios.
import pandas as pd  # Biblioteca para la manipulación y análisis de datos, especialmente con DataFrames.
from typing import (  # Módulo para proporcionar indicaciones de tipo (type hints), mejorando la legibilidad y ayudando al análisis estático.
    Optional,  # Indica que un tipo puede ser el tipo especificado o None.
    List,  # Indica una lista de un tipo específico.
    Tuple,  # Indica una tupla de tipos específicos.
    Union,  # Indica que un tipo puede ser uno de varios tipos especificados.
    Set,  # Indica un conjunto de un tipo específico.
    Dict,  # Indica un diccionario con tipos específicos para claves y valores.
    Any,  # Indica un tipo no restringido.
)
from enum import Enum, auto  # Módulo para crear enumeraciones, que son conjuntos de constantes simbólicas.

import platform  # Módulo para acceder a datos de identificación de la plataforma subyacente (sistema operativo).
import unicodedata  # Módulo para acceder a la Base de Datos de Caracteres Unicode (UCD).
import logging  # Módulo para emitir mensajes de registro desde bibliotecas y aplicaciones.
import json  # Módulo para trabajar con el formato de datos JSON.
import os  # Módulo que proporciona una forma de usar funcionalidades dependientes del sistema operativo.
from pathlib import Path  # Módulo que ofrece clases para representar rutas de sistema de archivos con semántica para diferentes SO.
import traceback # Módulo para obtener y formatear tracebacks de excepciones.

import numpy as np  # Biblioteca para computación numérica, fundamental para operaciones con arrays.

# --- Configuración del Logging ---
logger = logging.getLogger(__name__)

# --- Enumeraciones ---
class OrigenResultados(Enum):
    NINGUNO = 0
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto()
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto()
    DICCIONARIO_SIN_COINCIDENCIAS = auto()
    DIRECTO_DESCRIPCION_CON_RESULTADOS = auto()
    DIRECTO_DESCRIPCION_VACIA = auto()
    ERROR_CARGA_DICCIONARIO = auto()
    ERROR_CARGA_DESCRIPCION = auto()
    ERROR_CONFIGURACION_COLUMNAS_DICC = auto()
    ERROR_CONFIGURACION_COLUMNAS_DESC = auto()
    ERROR_BUSQUEDA_INTERNA_MOTOR = auto()
    TERMINO_INVALIDO = auto()
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC = auto()

    @property
    def es_via_diccionario(self) -> bool:
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
        return self in {OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, OrigenResultados.DIRECTO_DESCRIPCION_VACIA}
    @property
    def es_error_carga(self) -> bool:
        return self in {OrigenResultados.ERROR_CARGA_DICCIONARIO, OrigenResultados.ERROR_CARGA_DESCRIPCION}
    @property
    def es_error_configuracion(self) -> bool:
        return self in {OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC}
    @property
    def es_error_operacional(self) -> bool: return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR
    @property
    def es_termino_invalido(self) -> bool: return self == OrigenResultados.TERMINO_INVALIDO

class ExtractorMagnitud:
    MAPEO_MAGNITUDES_PREDEFINIDO: Dict[str, List[str]] = {}

    def __init__(self, mapeo_magnitudes: Optional[Dict[str, List[str]]] = None):
        self.sinonimo_a_canonico_normalizado: Dict[str, str] = {}
        # Utiliza el mapeo proporcionado o el predefinido (que está vacío por defecto).
        mapeo_a_usar = mapeo_magnitudes if mapeo_magnitudes is not None else self.MAPEO_MAGNITUDES_PREDEFINIDO
        
        # Itera sobre el mapeo para construir el diccionario interno normalizado.
        # La clave del mapeo_a_usar es la forma canónica, y el valor es una lista de sus sinónimos.
        for forma_canonica_original, lista_sinonimos_originales in mapeo_a_usar.items():
            # Normaliza la forma canónica.
            canonico_norm = self._normalizar_texto(forma_canonica_original)
            if not canonico_norm: # Si la forma canónica es inválida después de normalizar, la ignora.
                logger.warning(f"Forma canónica '{forma_canonica_original}' resultó vacía tras normalizar y fue ignorada en ExtractorMagnitud.")
                continue
            
            # Mapea la forma canónica normalizada a sí misma.
            self.sinonimo_a_canonico_normalizado[canonico_norm] = canonico_norm
            
            # Procesa y mapea cada sinónimo a la forma canónica normalizada.
            for sinonimo_original in lista_sinonimos_originales:
                sinonimo_norm = self._normalizar_texto(str(sinonimo_original)) # Asegura que sea string
                if sinonimo_norm: # Si el sinónimo es válido después de normalizar.
                    self.sinonimo_a_canonico_normalizado[sinonimo_norm] = canonico_norm
        logger.debug(f"ExtractorMagnitud inicializado/actualizado con {len(self.sinonimo_a_canonico_normalizado)} mapeos normalizados.")


    @staticmethod
    def _normalizar_texto(texto: str) -> str:
        if not isinstance(texto, str) or not texto: return ""
        try:
            texto_upper = texto.upper()
            forma_normalizada = unicodedata.normalize("NFKD", texto_upper)
            res = "".join(c for c in forma_normalizada if not unicodedata.combining(c) and (c.isalnum() or c.isspace() or c in ['.', '-', '_', '/']))
            return ' '.join(res.split())
        except TypeError:
            logger.error(f"TypeError en _normalizar_texto (ExtractorMagnitud) con entrada: {texto}")
            return ""

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        if not texto_unidad: return None
        normalizada = self._normalizar_texto(texto_unidad)
        return self.sinonimo_a_canonico_normalizado.get(normalizada) if normalizada else None

class ManejadorExcel:
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        ruta = Path(ruta_archivo)
        if not ruta.exists():
            mensaje_error = f"¡Archivo no encontrado! Ruta: {ruta}"
            logger.error(f"ManejadorExcel: {mensaje_error}")
            return None, mensaje_error
        try:
            engine: Optional[str] = None
            if ruta.suffix.lower() == ".xlsx": engine = "openpyxl"
            logger.info(f"ManejadorExcel: Cargando '{ruta.name}' con engine='{engine or 'auto (pandas intentará xlrd para .xls)'}'...")
            df = pd.read_excel(ruta, engine=engine)
            logger.info(f"ManejadorExcel: Archivo '{ruta.name}' ({len(df)} filas) cargado exitosamente.")
            return df, None
        except ImportError as ie:
            mensaje_error_usuario = (f"Error al cargar '{ruta.name}': Falta librería.\nPara .xlsx: pip install openpyxl\nPara .xls: pip install xlrd\nDetalle: {ie}")
            logger.exception(f"ManejadorExcel: Falta dependencia para leer '{ruta.name}'. Error: {ie}")
            return None, mensaje_error_usuario
        except Exception as e:
            mensaje_error_usuario = (f"No se pudo cargar '{ruta.name}': {e}\nVerifique formato, permisos y si está en uso.")
            logger.exception(f"ManejadorExcel: Error genérico al cargar '{ruta.name}'.")
            return None, mensaje_error_usuario

class MotorBusqueda:
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[Path] = None
        self.archivo_descripcion_actual: Optional[Path] = None
        self.indices_columnas_busqueda_dic_preview: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else []
        logger.info(f"MotorBusqueda inicializado. Índices preview dicc: {self.indices_columnas_busqueda_dic_preview or 'Todas texto/objeto'}")
        self.patron_comparacion = re.compile(r"^\s*([<>]=?)\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$")
        self.patron_rango = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$")
        self.patron_termino_negado = re.compile(r'#\s*(?:\"([^\"]+)\"|([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\.\-\_]+))', re.IGNORECASE | re.UNICODE)
        self.patron_num_unidad_df = re.compile(r"(\d+(?:[.,]\d+)?)[\s\-]*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?")
        self.extractor_magnitud = ExtractorMagnitud() # Se inicializa con el mapeo predefinido (vacío)

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        """
        Carga el archivo Excel que actúa como "diccionario" de términos (FCDs).
        Actualiza el ExtractorMagnitud con las formas canónicas y sinónimos del diccionario.
        """
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            self.extractor_magnitud = ExtractorMagnitud() # Resetear a predefinido si falla la carga
            return False, error_msg_carga

        # Construir mapeo para ExtractorMagnitud
        mapeo_dinamico_para_extractor: Dict[str, List[str]] = {}
        
        # Asumir que la primera columna (índice 0) es la forma canónica de la unidad/término.
        # Asumir que las columnas desde el índice 3 (cuarta columna) en adelante pueden ser sinónimos.
        # Esto debe coincidir con la estructura real de tu archivo "Diccionario_ACCESO draft vX.X.xlsx".
        if df_cargado.shape[1] > 0: # El DF debe tener al menos la columna canónica.
            columna_canonica_nombre = df_cargado.columns[0]
            # Determinar columnas de sinónimos (ej. de la 4ta hasta la 14ava, o hasta que no haya más)
            # Para el ejemplo del log: SINONIMOS (idx 3), Unnamed: 4 (idx 4) ... Unnamed: 13 (idx 13)
            # Usaremos un rango de índices que parece cubrir tus columnas de sinónimos.
            # Puedes ajustar estos índices si la estructura de tu Excel cambia.
            inicio_col_sinonimos = 3 
            # Definimos un máximo de columnas de sinónimos a leer para evitar errores si hay muchas columnas inesperadas.
            # O podríamos iterar todas las columnas restantes si se prefiere.
            max_cols_a_chequear_para_sinonimos = df_cargado.shape[1] # Chequear todas las columnas restantes
                                                                    # después de la columna canónica y las intermedias (concepto, tipo)

            for _, fila in df_cargado.iterrows():
                forma_canonica_raw = fila[columna_canonica_nombre]
                if pd.isna(forma_canonica_raw) or str(forma_canonica_raw).strip() == "":
                    continue # Saltar si la forma canónica está vacía

                # La forma canónica también es un sinónimo de sí misma.
                # El extractor se encarga de normalizarla internamente.
                forma_canonica_str = str(forma_canonica_raw).strip()
                
                # Lista para almacenar todos los sinónimos de esta forma canónica.
                # ExtractorMagnitud normalizará estos al inicializarse.
                sinonimos_para_esta_canonica: List[str] = [forma_canonica_str] 

                # Iterar sobre las columnas designadas como sinónimos
                for i in range(inicio_col_sinonimos, max_cols_a_chequear_para_sinonimos):
                    if i < df_cargado.shape[1]: # Asegurar que el índice de columna es válido
                        sinonimo_celda_raw = fila[df_cargado.columns[i]]
                        if pd.notna(sinonimo_celda_raw) and str(sinonimo_celda_raw).strip() != "":
                            sinonimos_para_esta_canonica.append(str(sinonimo_celda_raw).strip())
                
                # Añadir al mapeo. ExtractorMagnitud manejará la normalización y evitará duplicados
                # si la misma forma canónica (normalizada) aparece múltiples veces.
                # Lo correcto es que la *clave* del mapeo sea la forma canónica original (o normalizada)
                # y el *valor* la lista de sinónimos originales (o normalizados).
                # El constructor de ExtractorMagnitud se encarga de la normalización interna.
                forma_canonica_clave_para_mapeo = self.extractor_magnitud._normalizar_texto(forma_canonica_str)
                if forma_canonica_clave_para_mapeo:
                    # Si la forma canónica ya existe, extendemos su lista de sinónimos (evitando duplicados)
                    # La lógica de ExtractorMagnitud ya lo hace internamente al construir sinonimo_a_canonico_normalizado
                    # Aquí, simplemente pasamos la lista completa de sinónimos para esta canónica.
                    # Si una forma canónica aparece múltiples veces, la última definición de sinónimos para esa canónica prevalecerá.
                    # Para un mejor manejo de duplicados de formas canónicas, se podría acumular.
                    # Pero para el ExtractorMagnitud, él mapea cada sinónimo normalizado a una única canónica normalizada.
                    # Lo importante es que cada sinónimo se asocie a su canónica correcta.
                    mapeo_dinamico_para_extractor[forma_canonica_str] = list(set(sinonimos_para_esta_canonica)) # Usar set para eliminar duplicados de la lista de sinónimos.


            if mapeo_dinamico_para_extractor:
                self.extractor_magnitud = ExtractorMagnitud(mapeo_magnitudes=mapeo_dinamico_para_extractor)
                logger.info(f"Extractor de magnitudes actualizado desde '{ruta.name}' usando formas canónicas y sinónimos.")
            else:
                logger.warning(f"No se extrajeron mapeos de unidad válidos desde '{ruta.name}'. ExtractorMagnitud usará su predefinido (si existe) o estará vacío.")
                self.extractor_magnitud = ExtractorMagnitud() # Resetear a predefinido si no se cargó nada
        else:
            logger.warning(f"El archivo de diccionario '{ruta.name}' no tiene columnas. No se pudo actualizar el extractor de magnitudes.")
            self.extractor_magnitud = ExtractorMagnitud() # Resetear

        self.datos_diccionario = df_cargado
        self.archivo_diccionario_actual = ruta

        if logger.isEnabledFor(logging.DEBUG) and self.datos_diccionario is not None:
            logger.debug(f"Archivo de diccionario '{ruta.name}' cargado (primeras 3 filas):\n{self.datos_diccionario.head(3).to_string()}")
        return True, None

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        # ... (sin cambios en esta función)
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)
        if df_cargado is None:
            self.datos_descripcion = None; self.archivo_descripcion_actual = None
            return False, error_msg_carga
        self.datos_descripcion = df_cargado; self.archivo_descripcion_actual = ruta
        logger.info(f"Archivo de descripciones '{ruta.name}' cargado.")
        return True, None

    def _obtener_nombres_columnas_busqueda_df(self, df: pd.DataFrame, indices_cfg: List[int], tipo_busqueda: str) -> Tuple[Optional[List[str]], Optional[str]]:
        # ... (sin cambios en esta función)
        if df is None or df.empty: return None, f"DF para '{tipo_busqueda}' vacío."
        columnas_disponibles = list(df.columns); num_cols_df = len(columnas_disponibles)
        if num_cols_df == 0: return None, f"DF '{tipo_busqueda}' sin columnas."
        usar_columnas_por_defecto = not indices_cfg or indices_cfg == [-1]
        if usar_columnas_por_defecto:
            cols_texto_obj = [col for col in columnas_disponibles if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
            if cols_texto_obj:
                logger.debug(f"Para '{tipo_busqueda}', usando columnas de tipo texto/objeto (defecto): {cols_texto_obj}")
                return cols_texto_obj, None
            else:
                logger.warning(f"Para '{tipo_busqueda}', no hay cols texto/objeto. Usando todas las {num_cols_df} columnas: {columnas_disponibles}")
                return columnas_disponibles, None
        nombres_columnas_seleccionadas: List[str] = []
        indices_invalidos: List[str] = []
        for i in indices_cfg:
            if not (isinstance(i, int) and 0 <= i < num_cols_df): indices_invalidos.append(str(i))
            else: nombres_columnas_seleccionadas.append(columnas_disponibles[i])
        if indices_invalidos: return None, f"Índice(s) {', '.join(indices_invalidos)} inválido(s) para '{tipo_busqueda}'. Columnas: {num_cols_df} (0 a {num_cols_df-1})."
        if not nombres_columnas_seleccionadas: return None, f"Config de índices {indices_cfg} no resultó en columnas válidas para '{tipo_busqueda}'."
        logger.debug(f"Para '{tipo_busqueda}', usando columnas por índices {indices_cfg}: {nombres_columnas_seleccionadas}")
        return nombres_columnas_seleccionadas, None

    def _normalizar_para_busqueda(self, texto: str) -> str:
        # ... (sin cambios en esta función)
        if not isinstance(texto, str) or not texto: return ""
        try:
            texto_upper = texto.upper()
            texto_norm_nfkd = unicodedata.normalize('NFKD', texto_upper)
            texto_sin_acentos = "".join([c for c in texto_norm_nfkd if not unicodedata.combining(c)])
            texto_limpio_final = re.sub(r'[^\w\s\.\-\/\_]', '', texto_sin_acentos)
            return ' '.join(texto_limpio_final.split()).strip()
        except Exception as e:
            logger.error(f"Error al normalizar el texto '{texto[:50]}...': {e}")
            return str(texto).upper().strip()

    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, cols: List[str], texto: str) -> Tuple[pd.DataFrame, str, List[str]]:
        # ... (sin cambios en esta función)
        texto_limpio_entrada = texto.strip(); terminos_negados_encontrados: List[str] = []
        df_a_procesar = df_original.copy() if df_original is not None else pd.DataFrame()
        if not texto_limpio_entrada: return df_a_procesar, "", terminos_negados_encontrados
        partes_positivas: List[str] = []; ultimo_indice_fin_negado = 0
        for match_negado in self.patron_termino_negado.finditer(texto_limpio_entrada):
            partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:match_negado.start()])
            ultimo_indice_fin_negado = match_negado.end()
            termino_negado_raw = match_negado.group(1) or match_negado.group(2)
            if termino_negado_raw:
                termino_negado_normalizado = self._normalizar_para_busqueda(termino_negado_raw.strip('"'))
                if termino_negado_normalizado and termino_negado_normalizado not in terminos_negados_encontrados:
                    terminos_negados_encontrados.append(termino_negado_normalizado)
        partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:])
        terminos_positivos_final_str = ' '.join("".join(partes_positivas).split()).strip()
        if df_a_procesar.empty or not terminos_negados_encontrados or not cols:
            logger.debug(f"Parseo negación: Query='{texto_limpio_entrada}', Positivos='{terminos_positivos_final_str}', Negados={terminos_negados_encontrados}. No se aplicó filtro al DF.")
            return df_a_procesar, terminos_positivos_final_str, terminos_negados_encontrados
        mascara_exclusion_total = pd.Series(False, index=df_a_procesar.index)
        for termino_negado_actual in terminos_negados_encontrados:
            if not termino_negado_actual: continue
            mascara_para_este_termino_negado = pd.Series(False, index=df_a_procesar.index)
            patron_regex_negado = r"\b" + re.escape(termino_negado_actual) + r"\b"
            for nombre_columna in cols:
                if nombre_columna not in df_a_procesar.columns: continue
                try:
                    serie_columna_normalizada = df_a_procesar[nombre_columna].astype(str).map(self._normalizar_para_busqueda)
                    mascara_para_este_termino_negado |= serie_columna_normalizada.str.contains(patron_regex_negado, regex=True, na=False)
                except Exception as e_neg_col: logger.error(f"Error aplicando negación en col '{nombre_columna}', term '{termino_negado_actual}': {e_neg_col}")
            mascara_exclusion_total |= mascara_para_este_termino_negado
        df_resultado_filtrado = df_a_procesar[~mascara_exclusion_total]
        logger.info(f"Filtrado por negación (Query='{texto_limpio_entrada}'): {len(df_a_procesar)} -> {len(df_resultado_filtrado)} filas. Negados: {terminos_negados_encontrados}. Positivos: '{terminos_positivos_final_str}'")
        return df_resultado_filtrado, terminos_positivos_final_str, terminos_negados_encontrados

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        # ... (MODIFICADO para ya no tratar '/' como OR)
        texto_limpio = texto_complejo.strip();
        if not texto_limpio: return "OR", []
        if '+' in texto_limpio and not (texto_limpio.startswith("(") and texto_limpio.endswith(")")):
             logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Detectado '+' de alto nivel, tratando como AND. Segmento=['{texto_limpio}']")
             return "AND", [texto_limpio]

        separadores_or = [(r"\s*\|\s*", "|")] # <-- '/' ELIMINADO DE AQUÍ
        for sep_regex, sep_char_literal in separadores_or:
            if '+' not in texto_complejo and sep_char_literal in texto_limpio:
                segmentos_potenciales = [s.strip() for s in re.split(sep_regex, texto_limpio) if s.strip()]
                if len(segmentos_potenciales) > 1 or (len(segmentos_potenciales) == 1 and texto_limpio != segmentos_potenciales[0]):
                    logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Op=OR, Segs={segmentos_potenciales}")
                    return "OR", segmentos_potenciales
        logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Op=AND (no OR explícito de alto nivel), Seg=['{texto_limpio}']")
        return "AND", [texto_limpio]

    def _descomponer_nivel2_and(self, termino_segmento_n1: str) -> Tuple[str, List[str]]:
        # ... (sin cambios en esta función)
        termino_limpio = termino_segmento_n1.strip();
        if not termino_limpio: return "AND", []
        partes_crudas = re.split(r'\s+\+\s+', termino_limpio)
        partes_limpias_finales = [p.strip() for p in partes_crudas if p.strip()]
        logger.debug(f"Descomp. N2 (AND) para '{termino_segmento_n1}': Partes={partes_limpias_finales}")
        return "AND", partes_limpias_finales

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        # ... (sin cambios en esta función, pero _parse_numero que llama sí cambió)
        terminos_analizados: List[Dict[str, Any]] = []
        for termino_original_bruto in terminos_brutos:
            termino_original_procesado = str(termino_original_bruto).strip()
            es_frase_exacta = False
            termino_final_para_analisis = termino_original_procesado
            if len(termino_final_para_analisis) >= 2 and \
               termino_final_para_analisis.startswith('"') and \
               termino_final_para_analisis.endswith('"'):
                termino_final_para_analisis = termino_final_para_analisis[1:-1]
                es_frase_exacta = True
            if not termino_final_para_analisis: continue
            item_analizado: Dict[str, Any] = {"original": termino_final_para_analisis}
            match_comparacion = self.patron_comparacion.match(termino_final_para_analisis)
            match_rango = self.patron_rango.match(termino_final_para_analisis)
            if match_comparacion and not es_frase_exacta:
                operador_str, valor_str, unidad_str_raw = match_comparacion.groups()
                valor_numerico = self._parse_numero(valor_str) # <--- LLAMA A _parse_numero MODIFICADO
                if valor_numerico is not None:
                    mapa_operadores = {">": "gt", "<": "lt", ">=": "ge", "<=": "le", "=": "eq"}
                    unidad_canonica: Optional[str] = None
                    if unidad_str_raw and unidad_str_raw.strip(): unidad_canonica = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_raw.strip())
                    item_analizado.update({"tipo": mapa_operadores.get(operador_str), "valor": valor_numerico, "unidad_busqueda": unidad_canonica})
                else: item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            elif match_rango and not es_frase_exacta:
                valor1_str, valor2_str, unidad_str_r_raw = match_rango.groups()
                valor1_num = self._parse_numero(valor1_str); valor2_num = self._parse_numero(valor2_str) # <--- LLAMA A _parse_numero MODIFICADO
                if valor1_num is not None and valor2_num is not None:
                    unidad_canonica_r: Optional[str] = None
                    if unidad_str_r_raw and unidad_str_r_raw.strip(): unidad_canonica_r = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_r_raw.strip())
                    item_analizado.update({"tipo": "range", "valor": sorted([valor1_num, valor2_num]), "unidad_busqueda": unidad_canonica_r})
                else: item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            else:
                item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            terminos_analizados.append(item_analizado)
        logger.debug(f"Términos (post-AND) analizados para búsqueda detallada: {terminos_analizados}")
        return terminos_analizados

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        """
        Convierte una cadena que representa un número a float,
        manejando comas como decimales y puntos como miles según la regla de 3+ dígitos.
        Ej: "10.000" -> 10000.0; "10,00" -> 10.0; "1.234,56" -> 1234.56
        """
        if isinstance(num_str, (int, float)): # Si ya es numérico, lo devuelve como float.
            return float(num_str)
        if not isinstance(num_str, str): # Si no es un string, no se puede parsear.
            return None

        s_limpio = num_str.strip() # Limpia espacios al inicio/final.
        if not s_limpio: # Si está vacío después de limpiar, retorna None.
            return None

        # Paso 1: Normalizar comas a puntos para un manejo uniforme de puntos.
        partes = con_coma = s_limpio.split(',')
        
    

        try:
            if len(partes) == 1:
                # No hay puntos, o era originalmente un entero sin comas (ej. "100")
                logger.debug(f"Parseo num: '{num_str}' -> '{con_coma}' -> partes={partes}. Caso entero/simple float.")
                return float(partes[0])
            
            # len(partes) > 1, significa que hay al menos un punto.
            ultima_parte = partes[-1]
            partes_principales_str = "".join(partes[:-1])

            if ultima_parte.isdigit(): # Verifica si la última parte es completamente numérica.
                if len(ultima_parte) >= 3:
                    # El último punto (y todos los anteriores) se consideran separadores de miles.
                    # Se unen todas las partes (incluida la última) sin ningún punto.
                    numero_reconstruido_str = "".join(partes) # Ej: "10", "000" -> "10000"
                    logger.debug(f"Parseo num: '{num_str}' -> '{con_coma}' -> partes={partes}. Última parte '{ultima_parte}' (>=3 dig) -> miles. Reconstruido: '{numero_reconstruido_str}'")
                    return float(numero_reconstruido_str)
                elif len(ultima_parte) == 1 or len(ultima_parte) == 2 :
                    # El último punto es un separador decimal.
                    # Las partes principales (concatenadas sin puntos) forman la parte entera.
                    numero_reconstruido_str = f"{partes_principales_str}.{ultima_parte}" # Ej: "10" + "." + "00" -> "10.00"
                    logger.debug(f"Parseo num: '{num_str}' -> '{con_coma}' -> partes={partes}. Última parte '{ultima_parte}' (1-2 dig) -> decimal. Reconstruido: '{numero_reconstruido_str}'")
                    return float(numero_reconstruido_str)
                else: # len(ultima_parte) == 0 (ej: num_str era "10." o "1.234.")
                    if not ultima_parte: # El string original terminaba en un punto.
                        if partes_principales_str.isdigit(): # Asegurarse que lo que queda es un número
                           logger.debug(f"Parseo num: '{num_str}' -> '{con_coma}' -> partes={partes}. Última parte vacía. Reconstruido: '{partes_principales_str}'")
                           return float(partes_principales_str) # "10." -> 10.0 ; "1.234." -> 1234.0
                        else:
                           logger.warning(f"Parseo num: Formato no reconocido tras quitar punto final para '{num_str}'. Parte principal: '{partes_principales_str}'")
                           return None
                    else: # Última parte es dígito pero de longitud 0 (imposible) o no reconocida
                        logger.warning(f"Parseo num: Formato de última parte '{ultima_parte}' no reconocido para '{num_str}'.")
                        return None
            else: # Última parte no es enteramente dígitos (ej. "10.5V" si regex falló, o "10.A5")
                logger.warning(f"Parseo num: Última parte '{ultima_parte}' de '{num_str}' (procesado como '{con_coma}') no es puramente numérica.")
                # Intento de fallback: si el string original (normalizado con comas a puntos) es directamente convertible
                # Esto manejaría "10.5" donde el punto es decimal y no hay más puntos.
                if con_coma.count('.') <= 1: # Solo si hay 0 o 1 punto en total
                    try:
                        val_fallback = float(con_coma)
                        logger.debug(f"Parseo num: Fallback a conversión directa para '{con_coma}' -> {val_fallback}")
                        return val_fallback
                    except ValueError:
                        logger.warning(f"Parseo num: Fallback falló para '{con_coma}'.")
                        return None
                return None
        except ValueError:
            logger.warning(f"Parseo num: ValueError final al convertir '{num_str}' (procesado como '{con_coma}' o reconstruido) a float.")
            return None

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols: List[str], term_an: Dict[str, Any]) -> pd.Series:
        # ... (sin cambios significativos aquí, ya que depende de _parse_numero)
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
                            num_celda_val = self._parse_numero(num_celda_str) # <-- Usa _parse_numero mejorado
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
        # ... (sin cambios significativos aquí, pero puede ser afectado por _parse_numero en _analizar_terminos)
        if df is None or df.empty or not cols: return pd.Series(False, index=df.index if df is not None else None)
        if not term_an_seg: return pd.Series(False, index=df.index)
        mascara_final = pd.Series(True, index=df.index)
        for term_ind_an in term_an_seg:
            if term_ind_an["tipo"] == "str" and \
               ("|" in term_ind_an["original"] or "/" in term_ind_an["original"]) and \
               term_ind_an["original"].startswith("(") and term_ind_an["original"].endswith(")"):
                logger.debug(f"Segmento AND contiene sub-query OR: '{term_ind_an['original']}'. Se procesará por separado.")
                sub_mascara_or, err_sub_or = self._procesar_busqueda_en_df_objetivo(df, cols, term_ind_an["original"], None)
                if err_sub_or:
                    logger.warning(f"Sub-query OR '{term_ind_an['original']}' falló: {err_sub_or}")
                    return pd.Series(False, index=df.index)
                mascara_este_term = sub_mascara_or.reindex(df.index, fill_value=False)
            else:
                mascara_este_term = self._generar_mascara_para_un_termino(df, cols, term_ind_an)
            mascara_final &= mascara_este_term
            if not mascara_final.any(): break
        return mascara_final

    def _combinar_mascaras_de_segmentos_or(self, lista_mascaras: List[pd.Series], df_idx_ref: Optional[pd.Index] = None) -> pd.Series:
        # ... (sin cambios en esta función)
        if not lista_mascaras:
            return pd.Series(False, index=df_idx_ref) if df_idx_ref is not None else pd.Series(dtype=bool)
        idx_usar = df_idx_ref
        if idx_usar is None or idx_usar.empty:
            if lista_mascaras and not lista_mascaras[0].empty:
                idx_usar = lista_mascaras[0].index
        if idx_usar is None or idx_usar.empty:
            return pd.Series(dtype=bool)
        mascara_final = pd.Series(False, index=idx_usar)
        for masc_seg in lista_mascaras:
            if masc_seg.empty: continue
            mascara_alineada = masc_seg
            if not masc_seg.index.equals(idx_usar):
                try: mascara_alineada = masc_seg.reindex(idx_usar, fill_value=False)
                except Exception as e_reidx: logger.error(f"Fallo reindex máscara OR: {e_reidx}. Máscara ignorada."); continue
            mascara_final |= mascara_alineada
        return mascara_final

    def _procesar_busqueda_en_df_objetivo(self, df_obj: pd.DataFrame, cols_obj: List[str], termino_busqueda_original_para_este_df: str, terminos_negativos_adicionales: Optional[List[str]] = None) -> Tuple[pd.DataFrame, Optional[str]]:
        # ... (sin cambios significativos aquí, pero afectado por _parse_numero vía _analizar_terminos)
        logger.debug(f"Proc. búsqueda DF: Query='{termino_busqueda_original_para_este_df}' en {len(cols_obj)} cols de DF ({len(df_obj)} filas). Neg. Adic: {terminos_negativos_adicionales}")
        df_despues_negaciones_query, terminos_positivos_de_query, terminos_negados_de_query = \
            self._aplicar_negaciones_y_extraer_positivos(df_obj, cols_obj, termino_busqueda_original_para_este_df)
        df_actual_procesando = df_despues_negaciones_query
        if terminos_negativos_adicionales and not df_actual_procesando.empty:
            set_negados_de_query = set(terminos_negados_de_query)
            negativos_unicamente_adicionales_norm: List[str] = [neg_adic_norm for neg_adic_norm in terminos_negativos_adicionales if neg_adic_norm and neg_adic_norm not in set_negados_de_query]
            if negativos_unicamente_adicionales_norm:
                logger.debug(f"Aplicando neg. ADICIONALES: {negativos_unicamente_adicionales_norm} a {len(df_actual_procesando)} filas.")
                mascara_excluir_adicionales = pd.Series(False, index=df_actual_procesando.index)
                for term_neg_adic_norm_actual in negativos_unicamente_adicionales_norm:
                    mascara_temporal_neg_adic = pd.Series(False, index=df_actual_procesando.index)
                    patron_regex_neg_adic = r"\b" + re.escape(term_neg_adic_norm_actual) + r"\b"
                    for nombre_col in cols_obj:
                        if nombre_col not in df_actual_procesando.columns: continue
                        try:
                            serie_norm_df = df_actual_procesando[nombre_col].astype(str).map(self._normalizar_para_busqueda)
                            mascara_temporal_neg_adic |= serie_norm_df.str.contains(patron_regex_neg_adic, regex=True, na=False)
                        except Exception as e_neg_adic_col: logger.error(f"Error neg. adic. col '{nombre_col}', term '{term_neg_adic_norm_actual}': {e_neg_adic_col}")
                    mascara_excluir_adicionales |= mascara_temporal_neg_adic
                df_actual_procesando = df_actual_procesando[~mascara_excluir_adicionales]
                logger.info(f"Filtrado por neg. ADICIONALES: {len(df_despues_negaciones_query)} -> {len(df_actual_procesando)} filas.")
        terminos_positivos_final_para_parseo = terminos_positivos_de_query
        if df_actual_procesando.empty and not terminos_positivos_final_para_parseo.strip():
            logger.debug("DF vacío post-negaciones y sin términos positivos. Devolviendo DF vacío.")
            return df_actual_procesando.copy(), None
        if not terminos_positivos_final_para_parseo.strip():
            logger.debug(f"Sin términos positivos ('{terminos_positivos_final_para_parseo}'). Devolviendo DF post-negaciones ({len(df_actual_procesando)} filas).")
            return df_actual_procesando.copy(), None
        operador_nivel1, segmentos_nivel1_or = self._descomponer_nivel1_or(terminos_positivos_final_para_parseo)
        if not segmentos_nivel1_or:
            if termino_busqueda_original_para_este_df.strip() or terminos_positivos_final_para_parseo.strip():
                msg_error_segmentos = f"Térm. positivo '{terminos_positivos_final_para_parseo}' (de '{termino_busqueda_original_para_este_df}') inválido post-OR."
                logger.warning(msg_error_segmentos)
                return pd.DataFrame(columns=df_actual_procesando.columns), msg_error_segmentos
            else:
                logger.debug("Query original y positiva post-negación vacías. Devolviendo DF post-negaciones.")
                return df_actual_procesando.copy(), None
        lista_mascaras_para_or: List[pd.Series] = []
        for segmento_or_actual in segmentos_nivel1_or:
            _operador_nivel2, terminos_brutos_nivel2_and = self._descomponer_nivel2_and(segmento_or_actual)
            terminos_atomicos_analizados_and = self._analizar_terminos(terminos_brutos_nivel2_and)
            mascara_para_segmento_or_actual: pd.Series
            if not terminos_atomicos_analizados_and:
                if operador_nivel1 == "AND":
                    msg_error_and = f"Segmento AND '{segmento_or_actual}' sin términos atómicos válidos. Falla."
                    logger.warning(msg_error_and)
                    return pd.DataFrame(columns=df_actual_procesando.columns), msg_error_and
                logger.debug(f"Segmento OR '{segmento_or_actual}' sin términos atómicos. Se ignora para OR.")
                if not df_actual_procesando.empty: mascara_para_segmento_or_actual = pd.Series(False, index=df_actual_procesando.index)
                else: mascara_para_segmento_or_actual = pd.Series(dtype=bool)
            else:
                mascara_para_segmento_or_actual = self._aplicar_mascara_combinada_para_segmento_and(df_actual_procesando, cols_obj, terminos_atomicos_analizados_and)
            lista_mascaras_para_or.append(mascara_para_segmento_or_actual)
        if not lista_mascaras_para_or and not df_actual_procesando.empty :
            logger.error("Error interno: no se generaron máscaras OR a pesar de segmentos N1 y DF no vacío.")
            return pd.DataFrame(columns=df_actual_procesando.columns), "Error interno: no se generaron máscaras OR."
        elif not lista_mascaras_para_or and df_actual_procesando.empty:
            return df_actual_procesando.copy(), None
        mascara_final_df_objetivo = self._combinar_mascaras_de_segmentos_or(lista_mascaras_para_or, df_actual_procesando.index if not df_actual_procesando.empty else None)
        if mascara_final_df_objetivo.empty and not df_actual_procesando.empty:
             df_resultado_final = pd.DataFrame(columns=df_actual_procesando.columns)
        elif mascara_final_df_objetivo.empty and df_actual_procesando.empty:
             df_resultado_final = df_actual_procesando.copy()
        else:
             df_resultado_final = df_actual_procesando[mascara_final_df_objetivo].copy()
        logger.debug(f"Resultado _procesar_busqueda_en_df_objetivo para '{termino_busqueda_original_para_este_df}': {len(df_resultado_final)} filas.")
        return df_resultado_final, None

    def _extraer_terminos_de_fila_completa(self, fila_df: pd.Series) -> Set[str]:
        # ... (sin cambios en esta función)
        terminos_extraidos_de_fila: Set[str] = set()
        if fila_df is None or fila_df.empty: return terminos_extraidos_de_fila
        for valor_celda in fila_df.values:
            if pd.notna(valor_celda):
                texto_celda_str = str(valor_celda).strip()
                if texto_celda_str:
                    texto_celda_norm = self._normalizar_para_busqueda(texto_celda_str)
                    palabras_significativas_celda = [palabra for palabra in texto_celda_norm.split() if len(palabra) > 1 and not palabra.isdigit()]
                    if palabras_significativas_celda: terminos_extraidos_de_fila.update(palabras_significativas_celda)
                    elif texto_celda_norm and len(texto_celda_norm) > 1 and not texto_celda_norm.isdigit() and not self._parse_numero(texto_celda_norm):
                        terminos_extraidos_de_fila.add(texto_celda_norm)
        return terminos_extraidos_de_fila

    def buscar(self, termino_busqueda_original: str, buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[List[int]], Optional[str]]:
        # ... (Lógica de AND secuencial ya implementada en la v1.8/v1.8.1 se mantiene) ...
        # La modificación clave fue en _parse_numero y la lógica de carga de ExtractorMagnitud.
        logger.info(f"Motor.buscar INICIO: termino='{termino_busqueda_original}', via_dicc={buscar_via_diccionario_flag}")
        columnas_descripcion_ref = self.datos_descripcion.columns if self.datos_descripcion is not None else []
        df_vacio_para_descripciones = pd.DataFrame(columns=columnas_descripcion_ref)
        fcds_obtenidos_final_para_ui: Optional[pd.DataFrame] = None
        indices_fcds_a_resaltar_en_preview: Optional[List[int]] = None

        if not termino_busqueda_original.strip():
            if self.datos_descripcion is not None:
                logger.info("Término vacío. Devolviendo todas las descripciones.")
                return self.datos_descripcion.copy(), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
            else:
                logger.warning("Término vacío y descripciones no cargadas.")
                return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, "Descripciones no cargadas."

        if buscar_via_diccionario_flag:
            if self.datos_diccionario is None: return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, None, "Diccionario no cargado."
            columnas_dic_para_fcds, err_msg_cols_dic = self._obtener_nombres_columnas_busqueda_df(self.datos_diccionario, [], "diccionario_fcds_inicial")
            if not columnas_dic_para_fcds: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, None, err_msg_cols_dic

            _df_dummy, terminos_positivos_globales, terminos_negativos_globales = self._aplicar_negaciones_y_extraer_positivos(pd.DataFrame(), [], termino_busqueda_original)
            logger.info(f"Parseo global: Positivos='{terminos_positivos_globales}', Negativos Globales={terminos_negativos_globales}")

            if "+" in terminos_positivos_globales:
                logger.info(f"Detectada búsqueda AND en positivos globales: '{terminos_positivos_globales}'")
                partes_and = [p.strip() for p in terminos_positivos_globales.split("+") if p.strip()]
                df_resultado_acumulado_desc = self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame(columns=columnas_descripcion_ref)
                fcds_indices_acumulados = set()
                todas_partes_and_produjeron_terminos_validos = True
                hay_error_en_busqueda_de_parte_o_desc = False
                error_msg_critico_partes: Optional[str] = None

                if self.datos_descripcion is None:
                     logger.error("Archivo de descripciones no cargado, no se puede proceder con búsqueda AND vía diccionario.")
                     return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas para búsqueda AND."
                columnas_desc_para_filtrado, err_cols_desc_fil = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                if not columnas_desc_para_filtrado:
                    return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_fil

                for i, parte_and_actual_str in enumerate(partes_and):
                    if not parte_and_actual_str: continue
                    logger.debug(f"Procesando parte AND '{parte_and_actual_str}' (parte {i+1}/{len(partes_and)}) en diccionario...")
                    fcds_para_esta_parte, error_fcd_parte = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, columnas_dic_para_fcds, parte_and_actual_str, None)
                    if error_fcd_parte:
                        todas_partes_and_produjeron_terminos_validos = False; hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_fcd_parte
                        logger.warning(f"Parte AND '{parte_and_actual_str}' falló en diccionario con error: {error_fcd_parte}"); break
                    if fcds_para_esta_parte is None or fcds_para_esta_parte.empty:
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no encontró FCDs en diccionario."); break
                    fcds_indices_acumulados.update(fcds_para_esta_parte.index.tolist())
                    terminos_extraidos_de_esta_parte_set: Set[str] = set()
                    for _, fila_fcd in fcds_para_esta_parte.iterrows(): terminos_extraidos_de_esta_parte_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                    if not terminos_extraidos_de_esta_parte_set:
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' encontró FCDs, pero no se extrajeron términos de ellas."); break
                    terminos_or_con_comillas_actual = [f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_extraidos_de_esta_parte_set if t]
                    query_or_simple_actual = " | ".join(terminos_or_con_comillas_actual)
                    if not query_or_simple_actual:
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no generó una query OR válida para descripciones."); break
                    if df_resultado_acumulado_desc.empty and i >= 0:
                         logger.info(f"Resultados acumulados de descripción vacíos antes de aplicar filtro para '{parte_and_actual_str}'. Búsqueda AND final será vacía."); break
                    logger.info(f"Aplicando filtro OR para '{parte_and_actual_str}' (Query: '{query_or_simple_actual[:100]}...') sobre {len(df_resultado_acumulado_desc)} filas de descripción.")
                    df_resultado_acumulado_desc, error_sub_busqueda_desc = self._procesar_busqueda_en_df_objetivo(df_resultado_acumulado_desc, columnas_desc_para_filtrado, query_or_simple_actual, None)
                    if error_sub_busqueda_desc:
                        hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_sub_busqueda_desc
                        logger.error(f"Error en sub-búsqueda OR para '{query_or_simple_actual}': {error_sub_busqueda_desc}"); break
                    if df_resultado_acumulado_desc.empty:
                        logger.info(f"Filtro OR para '{parte_and_actual_str}' no encontró coincidencias en resultados acumulados. Búsqueda AND final será vacía."); break
                
                if fcds_indices_acumulados and self.datos_diccionario is not None:
                    fcds_obtenidos_final_para_ui = self.datos_diccionario.loc[list(fcds_indices_acumulados)].drop_duplicates().copy()
                    indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                else:
                    fcds_obtenidos_final_para_ui = pd.DataFrame(columns=self.datos_diccionario.columns if self.datos_diccionario is not None else [])
                    indices_fcds_a_resaltar_en_preview = []
                
                if hay_error_en_busqueda_de_parte_o_desc:
                    return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_msg_critico_partes
                if not todas_partes_and_produjeron_terminos_validos or df_resultado_acumulado_desc.empty:
                    origen_fallo_and = OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS if not todas_partes_and_produjeron_terminos_validos else OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                    logger.info(f"Búsqueda AND '{terminos_positivos_globales}' no produjo resultados finales en descripciones (Origen: {origen_fallo_and.name}).")
                    return df_vacio_para_descripciones, origen_fallo_and, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                
                resultados_desc_final_filtrado_and = df_resultado_acumulado_desc
                if not resultados_desc_final_filtrado_and.empty and terminos_negativos_globales:
                    logger.info(f"Aplicando negativos globales {terminos_negativos_globales} a {len(resultados_desc_final_filtrado_and)} filas (resultado del AND de ORs)")
                    query_solo_negados_globales = " ".join([f"#{neg}" for neg in terminos_negativos_globales])
                    df_temp_neg, _, _ = self._aplicar_negaciones_y_extraer_positivos(resultados_desc_final_filtrado_and, columnas_desc_para_filtrado, query_solo_negados_globales)
                    resultados_desc_final_filtrado_and = df_temp_neg
                
                logger.info(f"Búsqueda AND '{terminos_positivos_globales}' vía diccionario produjo {len(resultados_desc_final_filtrado_and)} resultados en descripciones.")
                return resultados_desc_final_filtrado_and, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
            else: 
                origen_propuesto_flujo_simple: OrigenResultados = OrigenResultados.NINGUNO
                fcds_query_simple: Optional[pd.DataFrame] = None
                if terminos_positivos_globales.strip():
                    logger.info(f"BUSCAR EN DICC (FCDs) - Positivos (sin AND de alto nivel): Query='{terminos_positivos_globales}'")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
                    try:
                        fcds_temp, error_dic_pos = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, columnas_dic_para_fcds, terminos_positivos_globales, None)
                        if error_dic_pos: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_pos
                        fcds_query_simple = fcds_temp
                    except Exception as e_dic_pos:
                        logger.exception("Excepción búsqueda en diccionario (positivos simples)."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-positivos simples): {e_dic_pos}"
                elif terminos_negativos_globales:
                    logger.info(f"BUSCAR EN DICC (FCDs) - Puramente Negativo: Negs Globales={terminos_negativos_globales}")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC
                    try:
                        fcds_temp, error_dic_neg = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, columnas_dic_para_fcds, "", terminos_negativos_adicionales=terminos_negativos_globales)
                        if error_dic_neg: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_neg
                        fcds_query_simple = fcds_temp
                    except Exception as e_dic_neg:
                        logger.exception("Excepción búsqueda en diccionario (puramente negativo)."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-negativo): {e_dic_neg}"
                else: return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None

                fcds_obtenidos_final_para_ui = fcds_query_simple
                if fcds_obtenidos_final_para_ui is not None and not fcds_obtenidos_final_para_ui.empty:
                    indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                    logger.info(f"FCDs obtenidas del diccionario (flujo simple/negativo): {len(fcds_obtenidos_final_para_ui)} filas.")
                else:
                    logger.info(f"No se encontraron FCDs en diccionario para '{termino_busqueda_original}' (flujo simple/negativo).")
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Descripciones no cargadas."
                terminos_para_buscar_en_descripcion_set: Set[str] = set()
                for _, fila_fcd in fcds_obtenidos_final_para_ui.iterrows(): terminos_para_buscar_en_descripcion_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                if not terminos_para_buscar_en_descripcion_set:
                    logger.info("FCDs encontrados (flujo simple/negativo), pero no se extrajeron términos para descripciones.")
                    origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                    if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC: origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                    return df_vacio_para_descripciones, origen_final_sinterm, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                logger.info(f"Términos para desc ({len(terminos_para_buscar_en_descripcion_set)} únicos, muestra): {sorted(list(terminos_para_buscar_en_descripcion_set))[:10]}...")
                terminos_or_con_comillas_desc = [f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_para_buscar_en_descripcion_set if t]
                query_or_para_desc_simple = " | ".join(terminos_or_con_comillas_desc)
                if not query_or_para_desc_simple:
                    origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                    if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC: origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                    return df_vacio_para_descripciones, origen_q_vacia, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Query OR para descripciones vacía."
                columnas_desc_final_simple, err_cols_desc_final_simple = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                if not columnas_desc_final_simple: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, err_cols_desc_final_simple
                negativos_a_aplicar_desc_simple = terminos_negativos_globales
                if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC: negativos_a_aplicar_desc_simple = []
                logger.info(f"BUSCAR EN DESC (vía FCD simple/negativa): Query='{query_or_para_desc_simple[:200]}...'. Neg. Globales a aplicar: {negativos_a_aplicar_desc_simple}")
                try:
                    resultados_desc_final_simple, error_busqueda_desc_simple = self._procesar_busqueda_en_df_objetivo(self.datos_descripcion, columnas_desc_final_simple, query_or_para_desc_simple,terminos_negativos_adicionales=negativos_a_aplicar_desc_simple)
                    if error_busqueda_desc_simple: return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_busqueda_desc_simple
                    if resultados_desc_final_simple is None or resultados_desc_final_simple.empty:
                        origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                        if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC: origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                        return df_vacio_para_descripciones, origen_res_desc_vacio_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                    else: return resultados_desc_final_simple, origen_propuesto_flujo_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                except Exception as e_desc_proc_simple:
                    logger.exception("Excepción búsqueda final en descripciones (flujo simple/negativo)."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, f"Error motor (desc final simple/negativo): {e_desc_proc_simple}"
        else: # Búsqueda directa en descripciones
            if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas."
            columnas_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion")
            if not columnas_desc_directo: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_directo
            try:
                logger.info(f"BUSCAR EN DESC (DIRECTO): Query '{termino_busqueda_original}'")
                resultados_directos_desc, error_busqueda_desc_dir = self._procesar_busqueda_en_df_objetivo(self.datos_descripcion, columnas_desc_directo, termino_busqueda_original, None)
                if error_busqueda_desc_dir: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_busqueda_desc_dir
                if resultados_directos_desc is None or resultados_directos_desc.empty: return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
                else: return resultados_directos_desc, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None, None
            except Exception as e_desc_dir_proc:
                logger.exception("Excepción búsqueda directa en descripciones."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (desc directa): {e_desc_dir_proc}"

# --- Interfaz Gráfica (sin cambios significativos más allá del título y _parse_numero indirectamente) ---
# ... (El resto de la clase InterfazGrafica y el bloque if __name__ == "__main__" se mantienen como en v1.8.1,
#      solo actualizando el número de versión en el título y el nombre del archivo de log)
class InterfazGrafica(tk.Tk):
    CONFIG_FILE_NAME = "config_buscador_avanzado_ui.json" 

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado v1.10 (Sinónimos Unidad Mejorados)") # Actualizado
        self.geometry("1250x800")
        self.config: Dict[str, Any] = self._cargar_configuracion_app()
        indices_cfg_preview_dic = self.config.get("indices_columnas_busqueda_dic_preview", [])
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg_preview_dic)
        self.resultados_actuales: Optional[pd.DataFrame] = None
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        self.ultimo_termino_buscado: Optional[str] = None
        self.reglas_guardadas: List[Dict[str, Any]] = []
        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.indices_fcds_resaltados: Optional[List[int]] = None
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO
        self.color_fila_par: str = "white"; self.color_fila_impar: str = "#f0f0f0"; self.color_resaltado_dic: str = "sky blue"
        self.op_buttons: Dict[str, ttk.Button] = {}
        self._configurar_estilo_ttk_app()
        self._crear_widgets_app()
        self._configurar_grid_layout_app()
        self._configurar_eventos_globales_app()
        self._configurar_tags_estilo_treeview_app()
        self._configurar_funcionalidad_orden_tabla(self.tabla_resultados)
        self._configurar_funcionalidad_orden_tabla(self.tabla_diccionario)
        self._actualizar_mensaje_barra_estado("Listo. Cargue Diccionario y Descripciones.")
        self._deshabilitar_botones_operadores()
        self._actualizar_estado_general_botones_y_controles()
        logger.info(f"Interfaz Gráfica (v1.10 Sinónimos Unidad Mejorados) inicializada.")

    def _try_except_wrapper(self, func, *args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            func_name = func.__name__; error_type = type(e).__name__; error_msg = str(e); tb_str = traceback.format_exc()
            logger.critical(f"Error en {func_name}: {error_type} - {error_msg}\n{tb_str}")
            print(f"--- TRACEBACK COMPLETO (desde _try_except_wrapper para {func_name}) ---\n{tb_str}")
            messagebox.showerror(f"Error Interno en {func_name}", f"Ocurrió un error inesperado:\n{error_type}: {error_msg}\n\nConsulte el log y la consola para el traceback completo.")
            if func_name in ["_cargar_diccionario_ui", "_cargar_excel_descripcion_ui"]: self._actualizar_etiquetas_archivos_cargados(); self._actualizar_estado_general_botones_y_controles()
            return None

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str): self._actualizar_estado_botones_operadores()
    def _cargar_configuracion_app(self) -> Dict[str, Any]:
        config_cargada: Dict[str, Any] = {}; ruta_archivo_config = Path(self.CONFIG_FILE_NAME)
        if ruta_archivo_config.exists():
            try:
                with ruta_archivo_config.open("r", encoding="utf-8") as f: config_cargada = json.load(f)
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE_NAME}")
            except Exception as e: logger.error(f"Error al cargar config '{self.CONFIG_FILE_NAME}': {e}")
        else: logger.info(f"Archivo config '{self.CONFIG_FILE_NAME}' no encontrado.")
        for clave_ruta in ["last_dic_path", "last_desc_path"]:
            valor_ruta = config_cargada.get(clave_ruta)
            config_cargada[clave_ruta] = Path(valor_ruta) if valor_ruta else None
        config_cargada.setdefault("indices_columnas_busqueda_dic_preview", [])
        return config_cargada
    def _guardar_configuracion_app(self):
        self.config["last_dic_path"] = str(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else None
        self.config["last_desc_path"] = str(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else None
        self.config["indices_columnas_busqueda_dic_preview"] = self.motor.indices_columnas_busqueda_dic_preview
        try:
            with open(self.CONFIG_FILE_NAME, "w", encoding="utf-8") as f: json.dump(self.config, f, indent=4)
            logger.info(f"Configuración guardada en: {self.CONFIG_FILE_NAME}")
        except Exception as e: logger.error(f"Error al guardar config '{self.CONFIG_FILE_NAME}': {e}")
    def _configurar_estilo_ttk_app(self):
        style = ttk.Style(self); os_name = platform.system(); prefs = {"Windows":["vista","xpnative"],"Darwin":["aqua"],"Linux":["clam","alt"]}
        theme = next((t for t in prefs.get(os_name,["clam"]) if t in style.theme_names()), style.theme_use() or "default")
        try: style.theme_use(theme); style.configure("Operator.TButton",padding=(2,1),font=("TkDefaultFont",9)); logger.info(f"Tema TTK: {theme}")
        except: logger.warning(f"Fallo al aplicar tema {theme}")
    def _crear_widgets_app(self):
        self.marco_controles=ttk.LabelFrame(self,text="Controles")
        self.btn_cargar_diccionario=ttk.Button(self.marco_controles,text="Cargar Diccionario",command=lambda: self._try_except_wrapper(self._cargar_diccionario_ui))
        self.lbl_dic_cargado=ttk.Label(self.marco_controles,text="Dic: Ninguno",width=25,anchor=tk.W,relief=tk.SUNKEN,borderwidth=1)
        self.btn_cargar_descripciones=ttk.Button(self.marco_controles,text="Cargar Descripciones",command=lambda: self._try_except_wrapper(self._cargar_excel_descripcion_ui))
        self.lbl_desc_cargado=ttk.Label(self.marco_controles,text="Desc: Ninguno",width=25,anchor=tk.W,relief=tk.SUNKEN,borderwidth=1)
        self.frame_ops=ttk.Frame(self.marco_controles)
        op_buttons_defs = [("+","+"),("|","|"),("#","#"),("> ",">"),("< ","<"),("≥ ",">="),("≤ ","<="),("-","-")]
        for i, (text, op_val_clean) in enumerate(op_buttons_defs):
            btn = ttk.Button(self.frame_ops,text=text,command=lambda op=op_val_clean: self._insertar_operador_validado(op),style="Operator.TButton",width=3)
            btn.grid(row=0,column=i,padx=1,pady=1,sticky="nsew"); self.op_buttons[op_val_clean] = btn
        self.entrada_busqueda=ttk.Entry(self.marco_controles,width=60,textvariable=self.texto_busqueda_var)
        self.btn_buscar=ttk.Button(self.marco_controles,text="Buscar",command=lambda: self._try_except_wrapper(self._ejecutar_busqueda_ui))
        self.btn_salvar_regla=ttk.Button(self.marco_controles,text="Salvar Regla",command=lambda: self._try_except_wrapper(self._salvar_regla_actual_ui),state="disabled")
        self.btn_ayuda=ttk.Button(self.marco_controles,text="?",command=self._mostrar_ayuda_ui,width=3)
        self.btn_exportar=ttk.Button(self.marco_controles,text="Exportar",command=lambda: self._try_except_wrapper(self._exportar_resultados_ui),state="disabled")
        self.lbl_tabla_diccionario=ttk.Label(self,text="Vista Previa Diccionario:")
        self.frame_tabla_diccionario=ttk.Frame(self);self.tabla_diccionario=ttk.Treeview(self.frame_tabla_diccionario,show="headings",height=8);self.scrolly_diccionario=ttk.Scrollbar(self.frame_tabla_diccionario,orient="vertical",command=self.tabla_diccionario.yview);self.scrollx_diccionario=ttk.Scrollbar(self.frame_tabla_diccionario,orient="horizontal",command=self.tabla_diccionario.xview);self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set,xscrollcommand=self.scrollx_diccionario.set)
        self.lbl_tabla_resultados=ttk.Label(self,text="Resultados / Descripciones:");self.frame_tabla_resultados=ttk.Frame(self);self.tabla_resultados=ttk.Treeview(self.frame_tabla_resultados,show="headings");self.scrolly_resultados=ttk.Scrollbar(self.frame_tabla_resultados,orient="vertical",command=self.tabla_resultados.yview);self.scrollx_resultados=ttk.Scrollbar(self.frame_tabla_resultados,orient="horizontal",command=self.tabla_resultados.xview);self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set,xscrollcommand=self.scrollx_resultados.set)
        self.barra_estado=ttk.Label(self,text="Listo.",relief=tk.SUNKEN,anchor=tk.W,borderwidth=1);self._actualizar_etiquetas_archivos_cargados()
    def _configurar_grid_layout_app(self):
        self.grid_rowconfigure(2,weight=1);self.grid_rowconfigure(4,weight=3);self.grid_columnconfigure(0,weight=1);self.marco_controles.grid(row=0,column=0,sticky="new",padx=10,pady=(10,5));self.marco_controles.grid_columnconfigure(1,weight=1);self.marco_controles.grid_columnconfigure(3,weight=1);self.btn_cargar_diccionario.grid(row=0,column=0,padx=(5,0),pady=5,sticky="w");self.lbl_dic_cargado.grid(row=0,column=1,padx=(2,10),pady=5,sticky="ew");self.btn_cargar_descripciones.grid(row=0,column=2,padx=(5,0),pady=5,sticky="w");self.lbl_desc_cargado.grid(row=0,column=3,padx=(2,5),pady=5,sticky="ew");self.frame_ops.grid(row=1,column=0,columnspan=6,padx=5,pady=(5,0),sticky="ew");[self.frame_ops.grid_columnconfigure(i,weight=1) for i in range(len(self.op_buttons))];self.entrada_busqueda.grid(row=2,column=0,columnspan=2,padx=5,pady=(0,5),sticky="ew");self.btn_buscar.grid(row=2,column=2,padx=(2,0),pady=(0,5),sticky="w");self.btn_salvar_regla.grid(row=2,column=3,padx=(2,0),pady=(0,5),sticky="w");self.btn_ayuda.grid(row=2,column=4,padx=(2,0),pady=(0,5),sticky="w");self.btn_exportar.grid(row=2,column=5,padx=(10,5),pady=(0,5),sticky="e");self.lbl_tabla_diccionario.grid(row=1,column=0,sticky="sw",padx=10,pady=(10,0));self.frame_tabla_diccionario.grid(row=2,column=0,sticky="nsew",padx=10,pady=(0,10));self.frame_tabla_diccionario.grid_rowconfigure(0,weight=1);self.frame_tabla_diccionario.grid_columnconfigure(0,weight=1);self.tabla_diccionario.grid(row=0,column=0,sticky="nsew");self.scrolly_diccionario.grid(row=0,column=1,sticky="ns");self.scrollx_diccionario.grid(row=1,column=0,sticky="ew");self.lbl_tabla_resultados.grid(row=3,column=0,sticky="sw",padx=10,pady=(0,0));self.frame_tabla_resultados.grid(row=4,column=0,sticky="nsew",padx=10,pady=(0,10));self.frame_tabla_resultados.grid_rowconfigure(0,weight=1);self.frame_tabla_resultados.grid_columnconfigure(0,weight=1);self.tabla_resultados.grid(row=0,column=0,sticky="nsew");self.scrolly_resultados.grid(row=0,column=1,sticky="ns");self.scrollx_resultados.grid(row=1,column=0,sticky="ew");self.barra_estado.grid(row=5,column=0,sticky="sew",padx=0,pady=(5,0))
    def _configurar_eventos_globales_app(self): self.entrada_busqueda.bind("<Return>",lambda e:self._try_except_wrapper(self._ejecutar_busqueda_ui));self.protocol("WM_DELETE_WINDOW",self.on_closing_app)
    def _actualizar_mensaje_barra_estado(self,m): self.barra_estado.config(text=m);logger.info(f"Mensaje UI (BarraEstado): {m}");self.update_idletasks()
    def _mostrar_ayuda_ui(self):
        texto_ayuda = ("Sintaxis:\n- Texto: `router cisco`\n- AND: `tarjeta + 16 puertos`\n- OR: `modulo | SFP` (Nota: `/` ya no es OR)\n"
                       "- Numérico: `>1000W`, `<50V`, `>=48A`, `<=10.5W`\n- Rango: `10-20V`\n- Frase: `\"rack 19\"`\n- Negación: `#palabra` o `# \"frase\"`\n\n"
                       "Flujo Vía Diccionario:\n1. Query 'A+B': Parte 'A' y 'B' se buscan individualmente en Diccionario (FCDs).\n"
                       "2. Sinónimos: De las FCDs de 'A' se extraen Sinónimos_A. De las FCDs de 'B' se extraen Sinónimos_B.\n"
                       "3. Búsqueda en Descripciones: Se buscan filas que contengan (ALGÚN Sinónimo_A) Y (ALGÚN Sinónimo_B) mediante filtrado secuencial.\n"
                       "4. Negativos (#global): Se aplican al final sobre los resultados de descripciones.\n"
                       "5. Falla en Diccionario: Si 'A' o 'B' no da FCDs/sinónimos, se ofrece búsqueda directa de 'A+B' en Descripciones.")
        messagebox.showinfo("Ayuda - Sintaxis y Flujo", texto_ayuda)
    def _configurar_tags_estilo_treeview_app(self):
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure("par", background=self.color_fila_par); tabla.tag_configure("impar", background=self.color_fila_impar)
        self.tabla_diccionario.tag_configure("resaltado_azul", background=self.color_resaltado_dic, foreground="black")
    def _configurar_funcionalidad_orden_tabla(self,tabla):
        cols = tabla["columns"]
        if cols: [tabla.heading(c,text=str(c),anchor=tk.W,command=lambda col=c,tbl=tabla: self._try_except_wrapper(self._ordenar_columna_tabla_ui,tbl,col,False)) for c in cols]
    def _ordenar_columna_tabla_ui(self,tabla,col,rev):
        df_copia=None;idx_resaltar=None
        if tabla==self.tabla_diccionario and self.motor.datos_diccionario is not None:df_copia=self.motor.datos_diccionario.copy();idx_resaltar=self.indices_fcds_resaltados
        elif tabla==self.tabla_resultados and self.resultados_actuales is not None:df_copia=self.resultados_actuales.copy()
        else: tabla.heading(col,command=lambda c=col,t=tabla:self._try_except_wrapper(self._ordenar_columna_tabla_ui,t,c,not rev));return
        if df_copia.empty or col not in df_copia.columns: tabla.heading(col,command=lambda c=col,t=tabla:self._try_except_wrapper(self._ordenar_columna_tabla_ui,t,c,not rev));return
        df_num=pd.to_numeric(df_copia[col],errors='coerce')
        df_ord=df_copia.sort_values(by=col,ascending=not rev,na_position='last',key=(lambda x:pd.to_numeric(x,errors='coerce')) if not df_num.isna().all() else (lambda x:x.astype(str).str.lower()))
        columnas_para_diccionario_ordenado = None
        if tabla==self.tabla_diccionario and self.motor.datos_diccionario is not None:
            # Obtener nombres de columna basados en los índices configurados, si existen, para la vista previa
            # Esto asegura que se usen las mismas columnas que en la carga inicial si están definidas.
            columnas_para_diccionario_ordenado, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                df_ord, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview"
            )
            if not columnas_para_diccionario_ordenado: columnas_para_diccionario_ordenado = list(df_ord.columns)
        if tabla==self.tabla_diccionario:self._actualizar_tabla_treeview_ui(tabla,df_ord,limite_filas=None,columnas_a_mostrar=columnas_para_diccionario_ordenado, indices_a_resaltar=idx_resaltar)
        elif tabla==self.tabla_resultados:self.resultados_actuales=df_ord;self._actualizar_tabla_treeview_ui(tabla,self.resultados_actuales)
        tabla.heading(col,command=lambda c=col,t=tabla:self._try_except_wrapper(self._ordenar_columna_tabla_ui,t,c,not rev));self._actualizar_mensaje_barra_estado(f"Ordenado por '{col}'.")

    def _actualizar_tabla_treeview_ui(self,tabla,datos,limite_filas=None,columnas_a_mostrar=None,indices_a_resaltar=None):
        is_dicc=tabla==self.tabla_diccionario; tabla_nombre = "Diccionario" if is_dicc else "Resultados"
        [tabla.delete(i) for i in tabla.get_children()];tabla["columns"]=()
        if datos is None or datos.empty:self._configurar_funcionalidad_orden_tabla(tabla); logger.debug(f"Tabla '{tabla_nombre}' vaciada."); return
        cols_orig=list(datos.columns); cols_para_usar_en_tabla: List[str]
        if columnas_a_mostrar:
            if all(isinstance(c, int) for c in columnas_a_mostrar):
                try: cols_para_usar_en_tabla = [cols_orig[i] for i in columnas_a_mostrar if 0 <= i < len(cols_orig)]
                except IndexError: logger.warning(f"Índices en columnas_a_mostrar fuera de rango para tabla '{tabla_nombre}'. Usando todas."); cols_para_usar_en_tabla = cols_orig
            elif all(isinstance(c, str) for c in columnas_a_mostrar): cols_para_usar_en_tabla = [c for c in columnas_a_mostrar if c in cols_orig]
            else: logger.warning(f"Tipo inesperado para columnas_a_mostrar en tabla '{tabla_nombre}'. Usando todas."); cols_para_usar_en_tabla = cols_orig
            if not cols_para_usar_en_tabla : logger.warning(f"columnas_a_mostrar no resultó en columnas válidas para tabla '{tabla_nombre}'. Usando todas."); cols_para_usar_en_tabla = cols_orig
        else: cols_para_usar_en_tabla = cols_orig
        if not cols_para_usar_en_tabla:self._configurar_funcionalidad_orden_tabla(tabla); logger.debug(f"Tabla '{tabla_nombre}' sin columnas usables."); return
        tabla["columns"]=tuple(cols_para_usar_en_tabla)
        for c in cols_para_usar_en_tabla:
            tabla.heading(c,text=str(c),anchor=tk.W)
            try:
                if c in datos.columns: ancho_contenido = datos[c].astype(str).str.len().quantile(0.95) if not datos[c].empty else 0
                else: ancho_contenido = 0 
                ancho_cabecera = len(str(c)); ancho = max(70, min(int(max(ancho_cabecera * 7, ancho_contenido * 5.5) + 15), 350))
            except Exception as e_ancho: logger.warning(f"Error calculando ancho para columna '{c}' en tabla '{tabla_nombre}': {e_ancho}"); ancho = 100
            tabla.column(c,anchor=tk.W,width=ancho,minwidth=50)
        df_iterar = datos[cols_para_usar_en_tabla]; num_filas_original=len(df_iterar)
        mostrar_todo_por_resaltado = is_dicc and indices_a_resaltar and num_filas_original > 0
        if not mostrar_todo_por_resaltado and limite_filas and num_filas_original > limite_filas: df_iterar=df_iterar.head(limite_filas)
        elif mostrar_todo_por_resaltado: logger.debug(f"Mostrando todas {num_filas_original} filas de '{tabla_nombre}' por resaltado.")
        for i,(idx,row) in enumerate(df_iterar.iterrows()):
            vals=[str(v) if pd.notna(v) else "" for v in row.values];tags=["par" if i%2==0 else "impar"]
            if is_dicc and indices_a_resaltar and idx in indices_a_resaltar:tags.append("resaltado_azul")
            try: tabla.insert("","end",values=vals,tags=tuple(tags),iid=f"row_{idx}")
            except Exception as e_ins: logger.warning(f"Error insertando fila {idx} en '{tabla_nombre}': {e_ins}")
        self._configurar_funcionalidad_orden_tabla(tabla); logger.debug(f"Tabla '{tabla_nombre}' actualizada con {len(tabla.get_children())} filas visibles.")

    def _actualizar_etiquetas_archivos_cargados(self):
        max_l=25;dic_p=self.motor.archivo_diccionario_actual;desc_p=self.motor.archivo_descripcion_actual
        dic_n=dic_p.name if dic_p else "Ninguno";desc_n=desc_p.name if desc_p else "Ninguno"
        dic_d=f"Dic: {dic_n}" if len(dic_n)<=max_l else f"Dic: ...{dic_n[-(max_l-4):]}";desc_d=f"Desc: {desc_n}" if len(desc_n)<=max_l else f"Desc: ...{desc_n[-(max_l-4):]}"
        self.lbl_dic_cargado.config(text=dic_d,foreground="green" if dic_p else "red");self.lbl_desc_cargado.config(text=desc_d,foreground="green" if desc_p else "red")
    def _actualizar_estado_general_botones_y_controles(self):
        dic_ok=self.motor.datos_diccionario is not None;desc_ok=self.motor.datos_descripcion is not None
        if dic_ok or desc_ok: self._actualizar_estado_botones_operadores()
        else: self._deshabilitar_botones_operadores()
        self.btn_buscar["state"]="normal" if dic_ok and desc_ok else "disabled";salvar_ok=False
        if self.ultimo_termino_buscado and self.origen_principal_resultados!=OrigenResultados.NINGUNO:
            if self.origen_principal_resultados.es_via_diccionario and ((self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty)or(self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and self.origen_principal_resultados in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC] )):salvar_ok=True
            elif (self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados==OrigenResultados.DIRECTO_DESCRIPCION_VACIA) and self.desc_finales_de_ultima_busqueda is not None:salvar_ok=True
        self.btn_salvar_regla["state"]="normal" if salvar_ok else "disabled";self.btn_exportar["state"]="normal" if (self.resultados_actuales is not None and not self.resultados_actuales.empty) else "disabled"

    def _cargar_diccionario_ui(self):
        cfg_path=self.config.get("last_dic_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada=filedialog.askopenfilename(title="Cargar Diccionario",filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada: return
        nombre_archivo = Path(ruta_seleccionada).name
        self._actualizar_mensaje_barra_estado(f"Cargando dicc: {nombre_archivo}...")
        self._actualizar_tabla_treeview_ui(self.tabla_diccionario,None);self._actualizar_tabla_treeview_ui(self.tabla_resultados,None);self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        ok,msg=self.motor.cargar_excel_diccionario(ruta_seleccionada)
        desc_n_title=Path(self.motor.archivo_descripcion_actual).name if self.motor.archivo_descripcion_actual else "N/A"
        if ok and self.motor.datos_diccionario is not None:
            self.config["last_dic_path"]=Path(ruta_seleccionada);self._guardar_configuracion_app();df_d=self.motor.datos_diccionario;n_filas=len(df_d)
            cols_prev,_=self.motor._obtener_nombres_columnas_busqueda_df(df_d,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
            self.lbl_tabla_diccionario.config(text=f"Diccionario ({n_filas} filas)");self._actualizar_tabla_treeview_ui(self.tabla_diccionario,df_d,limite_filas=100,columnas_a_mostrar=cols_prev)
            self.title(f"Buscador - Dic: {nombre_archivo} | Desc: {desc_n_title}");self._actualizar_mensaje_barra_estado(f"Diccionario '{nombre_archivo}' ({n_filas}) cargado.")
        else:
            self._actualizar_mensaje_barra_estado(f"Error cargando diccionario: {msg or 'Desconocido'}");messagebox.showerror("Error Carga Dicc",msg or "Error desconocido")
            self.title(f"Buscador - Dic: N/A (Error) | Desc: {desc_n_title}")
        self._actualizar_etiquetas_archivos_cargados();self._actualizar_estado_general_botones_y_controles()

    def _cargar_excel_descripcion_ui(self):
        cfg_path=self.config.get("last_desc_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada_str=filedialog.askopenfilename(title="Cargar Descripciones",filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada_str: logger.info("Carga de descripciones cancelada."); return
        nombre_archivo = Path(ruta_seleccionada_str).name;
        self._actualizar_mensaje_barra_estado(f"Cargando descripciones: {nombre_archivo}...")
        self.resultados_actuales=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self._actualizar_tabla_treeview_ui(self.tabla_resultados,None)
        ok, msg_error = self.motor.cargar_excel_descripcion(ruta_seleccionada_str)
        dic_n_title=Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A"
        if ok and self.motor.datos_descripcion is not None:
            self.config["last_desc_path"] = Path(ruta_seleccionada_str); self._guardar_configuracion_app()
            df_desc = self.motor.datos_descripcion; num_filas = len(df_desc)
            self._actualizar_mensaje_barra_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando vista previa...")
            self._actualizar_tabla_treeview_ui(self.tabla_resultados, df_desc, limite_filas=200)
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: {nombre_archivo}")
        else:
            error_a_mostrar = msg_error or "Ocurrió un error desconocido al cargar el archivo de descripciones."
            self._actualizar_mensaje_barra_estado(f"Error cargando descripciones: {error_a_mostrar}"); messagebox.showerror("Error al Cargar Archivo de Descripciones", error_a_mostrar)
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: N/A (Error)")
        self._actualizar_etiquetas_archivos_cargados();self._actualizar_estado_general_botones_y_controles()

    def _ejecutar_busqueda_ui(self):
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:messagebox.showwarning("Archivos Faltantes","Cargue Diccionario y Descripciones.");return
        term_ui=self.texto_busqueda_var.get();self.ultimo_termino_buscado=term_ui
        self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        self._actualizar_tabla_treeview_ui(self.tabla_resultados,None);self._actualizar_mensaje_barra_estado(f"Buscando '{term_ui}'...")
        res_df,origen,fcds,idx_res,err_msg = self.motor.buscar(termino_busqueda_original=term_ui, buscar_via_diccionario_flag=True)
        self.fcds_de_ultima_busqueda=fcds;self.origen_principal_resultados=origen;self.indices_fcds_resaltados=idx_res
        df_desc_cols=self.motor.datos_descripcion.columns if self.motor.datos_descripcion is not None else []
        if self.motor.datos_diccionario is not None:
            num_fcds_actual=len(self.indices_fcds_resaltados) if self.indices_fcds_resaltados else 0
            dicc_lbl=f"Diccionario ({len(self.motor.datos_diccionario)} filas)" + (f" - {num_fcds_actual} FCDs resaltados" if num_fcds_actual>0 and origen.es_via_diccionario and origen!=OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS else "")
            self.lbl_tabla_diccionario.config(text=dicc_lbl)
            cols_prev_dic_actual,_=self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
            limite_filas_dic_preview = None if self.indices_fcds_resaltados else 100
            self._actualizar_tabla_treeview_ui(self.tabla_diccionario,self.motor.datos_diccionario,limite_filas=limite_filas_dic_preview,columnas_a_mostrar=cols_prev_dic_actual,indices_a_resaltar=self.indices_fcds_resaltados)
        if err_msg and origen.es_error_operacional:messagebox.showerror("Error Motor",f"Error interno: {err_msg}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        elif origen.es_error_carga or origen.es_error_configuracion or origen.es_termino_invalido:messagebox.showerror("Error Búsqueda",err_msg or f"Error: {origen.name}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        elif origen in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC]:
            self.resultados_actuales=res_df;self._actualizar_mensaje_barra_estado(f"'{term_ui}': {len(fcds) if fcds is not None else 0} en Dic, {len(res_df) if res_df is not None else 0} en Desc.")
        elif origen==OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS:
            self.resultados_actuales=res_df ;self._actualizar_mensaje_barra_estado(f"'{term_ui}': No en Diccionario.");
            if messagebox.askyesno("Búsqueda Alternativa",f"'{term_ui}' no encontrado en Diccionario.\n\n¿Buscar '{term_ui}' directamente en Descripciones?"):
                self._try_except_wrapper(self._buscar_directo_en_descripciones_y_actualizar_ui, term_ui, df_desc_cols)
            else: self._actualizar_estado_general_botones_y_controles()
        elif origen in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC]:
            self.resultados_actuales=res_df;num_fcds_i=len(fcds) if fcds is not None else 0;msg_fcd_i=f"{num_fcds_i} en Diccionario"
            msg_desc_i="pero no se extrajeron términos válidos para Desc." if origen in [OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC] else "pero 0 resultados en Desc."
            self._actualizar_mensaje_barra_estado(f"'{term_ui}': {msg_fcd_i}, {msg_desc_i.replace('.','')} en Desc.")
            if messagebox.askyesno("Búsqueda Alternativa",f"{msg_fcd_i} para '{term_ui}', {msg_desc_i}\n\n¿Buscar '{term_ui}' directamente en Descripciones?"):
                self._try_except_wrapper(self._buscar_directo_en_descripciones_y_actualizar_ui, term_ui, df_desc_cols)
            else: self._actualizar_estado_general_botones_y_controles()
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS:self.resultados_actuales=res_df;self._actualizar_mensaje_barra_estado(f"Búsqueda directa '{term_ui}': {len(res_df) if res_df is not None else 0} resultados.")
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            self.resultados_actuales=res_df;num_r=len(res_df) if res_df is not None else 0
            self._actualizar_mensaje_barra_estado(f"Mostrando todas las desc ({num_r})." if not term_ui.strip() else f"Búsqueda directa '{term_ui}': 0 resultados.")
            if term_ui.strip() and num_r==0 :messagebox.showinfo("Info",f"No resultados para '{term_ui}' en búsqueda directa.")
        if self.resultados_actuales is None:self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        self.desc_finales_de_ultima_busqueda=self.resultados_actuales.copy();self._actualizar_tabla_treeview_ui(self.tabla_resultados,self.resultados_actuales);self._actualizar_estado_general_botones_y_controles()

    def _buscar_directo_en_descripciones_y_actualizar_ui(self, term_ui_original: str, columnas_df_desc_referencia: List[str]):
        self._actualizar_mensaje_barra_estado(f"Iniciando búsqueda directa de '{term_ui_original}' en descripciones...")
        self.indices_fcds_resaltados = None
        if self.motor.datos_diccionario is not None:
            cols_prev_dic_alt,_ = self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview")
            self.lbl_tabla_diccionario.config(text=f"Vista Previa Diccionario ({len(self.motor.datos_diccionario)} filas)")
            self._actualizar_tabla_treeview_ui(self.tabla_diccionario, self.motor.datos_diccionario, limite_filas=100, columnas_a_mostrar=cols_prev_dic_alt, indices_a_resaltar=None)
        res_df_dir, orig_dir, _, _, msg_error_directo = self.motor.buscar(termino_busqueda_original=term_ui_original, buscar_via_diccionario_flag=False)
        self.origen_principal_resultados = orig_dir; self.fcds_de_ultima_busqueda = None
        if msg_error_directo and (orig_dir.es_error_operacional or orig_dir.es_termino_invalido):
            messagebox.showerror("Error Búsqueda Directa", f"Error: {msg_error_directo}"); self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_referencia)
        else: self.resultados_actuales = res_df_dir
        num_rdd = len(self.resultados_actuales) if self.resultados_actuales is not None else 0
        self._actualizar_mensaje_barra_estado(f"Búsqueda directa de '{term_ui_original}': {num_rdd} resultados.")
        if num_rdd == 0 and orig_dir == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and term_ui_original.strip():
            messagebox.showinfo("Info", f"No resultados para '{term_ui_original}' en búsqueda directa.")
        if self.resultados_actuales is None: self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_referencia)
        self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy()
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, self.resultados_actuales)
        self._actualizar_estado_general_botones_y_controles()

    def _salvar_regla_actual_ui(self):
        origen_nombre = self.origen_principal_resultados.name
        if not self.ultimo_termino_buscado and not (self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and self.desc_finales_de_ultima_busqueda is not None): messagebox.showerror("Error Salvar", "No hay búsqueda para salvar."); return
        df_salvar: Optional[pd.DataFrame] = None; tipo_datos = "DESCONOCIDO"
        if self.origen_principal_resultados.es_via_diccionario:
            if self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty: df_salvar = self.desc_finales_de_ultima_busqueda; tipo_datos = "DESC_VIA_DICC"
            elif self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty: df_salvar = self.fcds_de_ultima_busqueda; tipo_datos = "FCDS_DICC"
        elif self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            if self.desc_finales_de_ultima_busqueda is not None: df_salvar = self.desc_finales_de_ultima_busqueda; tipo_datos = "DESC_DIRECTA";
            if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and not (self.ultimo_termino_buscado or "").strip(): tipo_datos = "TODAS_DESC"
        if df_salvar is not None:
            regla = {"termino": self.ultimo_termino_buscado or "N/A", "origen": origen_nombre, "tipo": tipo_datos, "filas": len(df_salvar), "ts": pd.Timestamp.now().isoformat()}
            self.reglas_guardadas.append(regla); self._actualizar_mensaje_barra_estado(f"Búsqueda '{self.ultimo_termino_buscado}' registrada."); messagebox.showinfo("Regla Salvada", f"Metadatos de '{self.ultimo_termino_buscado}' guardados.")
            logger.info(f"Regla guardada: {regla}")
        else: messagebox.showwarning("Nada que Salvar", "No hay datos claros para salvar.")
        self._actualizar_estado_general_botones_y_controles()

    def _exportar_resultados_ui(self):
        if self.resultados_actuales is None or self.resultados_actuales.empty: messagebox.showinfo("Exportar", "No hay resultados para exportar."); return
        nombre_archivo_sugerido = f"resultados_{pd.Timestamp.now():%Y%m%d_%H%M%S}"
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")], title="Guardar resultados", initialfile=nombre_archivo_sugerido)
        if not ruta: return
        if ruta.endswith(".xlsx"): self.resultados_actuales.to_excel(ruta, index=False)
        elif ruta.endswith(".csv"): self.resultados_actuales.to_csv(ruta, index=False, encoding='utf-8-sig')
        else: messagebox.showerror("Error Formato", "Usar .xlsx o .csv."); return
        messagebox.showinfo("Exportado", f"Resultados exportados a:\n{ruta}"); self._actualizar_mensaje_barra_estado(f"Exportado a {Path(ruta).name}")

    def _actualizar_estado_botones_operadores(self):
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None: self._deshabilitar_botones_operadores(); return
        [btn.config(state="normal") for btn in self.op_buttons.values()]
        txt=self.texto_busqueda_var.get();cur_pos=self.entrada_busqueda.index(tk.INSERT)
        last_char_rel=txt[:cur_pos].strip()[-1:] if txt[:cur_pos].strip() else ""
        ops_logicos=["+","|","/"]; ops_comp_pref=[">","<"];
        if not last_char_rel or last_char_rel in ops_logicos + ["#","<",">","=","-"]:
            if self.op_buttons.get("+"): self.op_buttons["+"]["state"]="disabled"
            if self.op_buttons.get("|"): self.op_buttons["|"]["state"]="disabled"
        if last_char_rel and last_char_rel not in ops_logicos + [" "]:
             if self.op_buttons.get("#"): self.op_buttons["#"]["state"]="disabled"
        if last_char_rel in [">","<","="]:
            for opk in ops_comp_pref + ["=","-"]:
                if self.op_buttons.get(opk): self.op_buttons[opk]["state"]="disabled"
            if last_char_rel == ">" and self.op_buttons.get(">="): self.op_buttons[">="]["state"]="disabled"
            if last_char_rel == "<" and self.op_buttons.get("<="): self.op_buttons["<="]["state"]="disabled"
        if last_char_rel.isdigit():
            for opk_pref in ops_comp_pref + ["=","#"]:
                 if self.op_buttons.get(opk_pref): self.op_buttons[opk_pref]["state"]="disabled"
        elif not last_char_rel or last_char_rel in [" ","+","|","/"]:
            if self.op_buttons.get("-"): self.op_buttons["-"]["state"]="disabled"
    def _insertar_operador_validado(self,op_limpio):
        ops_con_espacio_alrededor = ["+", "|", "/"] # Eliminado '-' de aquí, ya que en rangos no lleva espacios obligatorios
        texto_a_insertar: str
        if op_limpio in ops_con_espacio_alrededor: texto_a_insertar = f" {op_limpio} "
        elif op_limpio == "-": texto_a_insertar = f"{op_limpio}" # Para rangos tipo "10-20"
        elif op_limpio in [">=", "<="]: texto_a_insertar = f"{op_limpio}"
        elif op_limpio in [">", "<", "="]: texto_a_insertar = f"{op_limpio}"
        elif op_limpio == "#": texto_a_insertar = f"{op_limpio} "
        else: texto_a_insertar = op_limpio
        self.entrada_busqueda.insert(tk.INSERT,texto_a_insertar);self.entrada_busqueda.focus_set()
        self._actualizar_estado_botones_operadores()
    def _deshabilitar_botones_operadores(self): [btn.config(state="disabled") for btn in self.op_buttons.values()]
    def on_closing_app(self):
        try:
            logger.info("Cerrando aplicación Buscador Avanzado...")
            self._guardar_configuracion_app()
            self.destroy()
        except Exception as e:
            func_name = "on_closing_app"; error_type = type(e).__name__; error_msg = str(e); tb_str = traceback.format_exc()
            logger.critical(f"Error en {func_name}: {error_type} - {error_msg}\n{tb_str}")
            print(f"--- TRACEBACK COMPLETO (desde {func_name}) ---\n{tb_str}")
            self.destroy()

# --- Punto de Entrada Principal de la Aplicación ---
if __name__ == "__main__":
    LOG_FILE_NAME = "Buscador_Avanzado_App_v1.10.log" # Versión con Sinónimos Unidad Mejorados
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s",
        handlers=[logging.FileHandler(LOG_FILE_NAME, encoding="utf-8", mode="w"), logging.StreamHandler()])
    root_logger = logging.getLogger()
    root_logger.info(f"--- Iniciando Buscador Avanzado v1.10 (Sinónimos Unidad Mejorados) (Script: {Path(__file__).name}) ---")
    root_logger.info(f"Logs siendo guardados en: {Path(LOG_FILE_NAME).resolve()}")

    dependencias_faltantes_main: List[str] = []
    try: import pandas as pd_check_main; root_logger.info(f"Pandas: {pd_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("pandas")
    try: import openpyxl as opxl_check_main; root_logger.info(f"openpyxl: {opxl_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("openpyxl")
    try: import numpy as np_check_main; root_logger.info(f"Numpy: {np_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("numpy")
    try: import xlrd as xlrd_check_main; root_logger.info(f"xlrd: {xlrd_check_main.__version__}")
    except ImportError: root_logger.warning("xlrd no encontrado. Carga de .xls antiguos podría fallar.")

    if dependencias_faltantes_main:
        mensaje_error_deps_main = (f"Faltan dependencias críticas: {', '.join(dependencias_faltantes_main)}.\nInstale con: pip install {' '.join(dependencias_faltantes_main)}")
        root_logger.critical(mensaje_error_deps_main)
        try:
            root_error_tk_main = tk.Tk(); root_error_tk_main.withdraw()
            messagebox.showerror("Dependencias Faltantes", mensaje_error_deps_main); root_error_tk_main.destroy()
        except Exception as e_tk_dep_main: print(f"ERROR CRITICO (Error al mostrar mensaje Tkinter: {e_tk_dep_main}): {mensaje_error_deps_main}")
        exit(1)
    try: app=InterfazGrafica();app.mainloop()
    except Exception as e_main_app_exc:
        root_logger.critical("Error fatal no controlado en la aplicación principal:", exc_info=True)
        tb_str_fatal = traceback.format_exc()
        print(f"--- TRACEBACK FATAL (desde bloque __main__) ---\n{tb_str_fatal}")
        try:
            root_fatal_tk_main = tk.Tk(); root_fatal_tk_main.withdraw()
            messagebox.showerror("Error Fatal Inesperado", f"Error crítico: {e_main_app_exc}\nConsulte '{LOG_FILE_NAME}' y la consola."); root_fatal_tk_main.destroy()
        except: print(f"ERROR FATAL: {e_main_app_exc}. Revise '{LOG_FILE_NAME}'.")
    finally: root_logger.info(f"--- Finalizando Buscador ---")
