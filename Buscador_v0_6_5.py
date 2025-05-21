# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import (
    Optional,
    List,
    Tuple,
    Union,
    Set,
    Callable,
    Dict,
    Any,
    Literal,
)
from enum import Enum, auto
import traceback
import platform
import unicodedata
import logging
import json
import os
from pathlib import Path
import string

# --- Configuración del Logging ---
# Configurado globalmente al inicio del script en if __name__ == "__main__"
logger = logging.getLogger(__name__)


# --- Enumeraciones (Versión Mejorada del segundo script) ---
class OrigenResultados(Enum):
    NINGUNO = 0
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto() # FCDs exist, but no terms to search in desc
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto()  # FCDs exist, but no desc match
    DIRECTO_DESCRIPCION_CON_RESULTADOS = auto()
    DIRECTO_DESCRIPCION_VACIA = auto() # Direct search performed, no results / or search term was empty
    ERROR_CARGA_DICCIONARIO = auto()
    ERROR_CARGA_DESCRIPCION = auto()
    ERROR_CONFIGURACION_COLUMNAS_DICC = auto() # Specific to dictionary column config
    ERROR_CONFIGURACION_COLUMNAS_DESC = auto() # Specific to description column config (if applicable)
    ERROR_BUSQUEDA_INTERNA_MOTOR = auto() # Generic error within motor's search logic

    @property
    def es_via_diccionario(self) -> bool:
        return self in {
            OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
            OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC,
        }

    @property
    def es_directo_descripcion(self) -> bool:
        return self in {
            OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS,
            OrigenResultados.DIRECTO_DESCRIPCION_VACIA,
        }

    @property
    def es_error_carga(self) -> bool:
        return self in {
            OrigenResultados.ERROR_CARGA_DICCIONARIO,
            OrigenResultados.ERROR_CARGA_DESCRIPCION,
        }

    @property
    def es_error_configuracion(self) -> bool:
        return self in {
            OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC,
            OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC,
        }

    @property
    def es_error_operacional(self) -> bool: # Error durante la lógica de búsqueda en sí
        return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR


# --- Clases de Utilidad ---
class ExtractorMagnitud:
    MAGNITUDES_PREDEFINIDAS: List[str] = sorted(list(set([ # sorted(list(set(...))) para unicidad y orden
        "A","AMP","AMPS","AH","ANTENNA","BASE","BIT","ETH","FE","G","GB",
        "GBE","GE","GIGABIT","GBASE","GBASEWAN","GBIC","GBIT","GBPS","GH",
        "GHZ","HZ","KHZ","KM","KVA","KW","LINEAS","LINES","MHZ","NM","PORT",
        "PORTS","PTOS","PUERTO","PUERTOS","P","V","VA","VAC","VC","VCC",
        "VCD","VDC","W","WATTS","E","POTS","STM", "VOLTIOS", "VATIOS", "AMPERIOS"
        # "FE", "GBE", "GE" eran duplicados en script original, set los elimina. "E" estaba
    ])))

    def __init__(self, magnitudes: Optional[List[str]] = None):
        self.magnitudes_normalizadas: Dict[str, str] = {
            self._normalizar_texto(m.upper()): m
            for m in (
                magnitudes if magnitudes is not None else self.MAGNITUDES_PREDEFINIDAS
            )
        }
        # logger.debug(f"ExtractorMagnitud inicializado con {len(self.magnitudes_normalizadas)} magnitudes.")

    @staticmethod
    def _normalizar_texto(texto: str) -> str: # Versión primordial
        if not isinstance(texto, str) or not texto: return ""
        try:
            forma_normalizada = unicodedata.normalize('NFKD', texto.upper())
            return ''.join(c for c in forma_normalizada if not unicodedata.combining(c))
        except TypeError:
            # logger.warning(f"TypeError al normalizar texto: {texto}")
            return "" # Devuelve vacío si la normalización falla

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        normalizada = self._normalizar_texto(texto_unidad)
        return self.magnitudes_normalizadas.get(normalizada)


class ManejadorExcel: # (Versión mejorada del segundo script)
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        """ Carga un archivo Excel. Retorna (DataFrame, mensaje_error_o_none). """
        ruta = Path(ruta_archivo)
        logger.info(f"Intentando cargar archivo Excel: {ruta}")
        if not ruta.exists():
            error_msg = f"¡Archivo no encontrado! Ruta: {ruta}"
            logger.error(error_msg)
            return None, error_msg
        try:
            engine = "openpyxl" if ruta.suffix.lower() == ".xlsx" else None # .lower() para robustez
            df = pd.read_excel(ruta, engine=engine)
            logger.info(f"Archivo '{ruta.name}' cargado ({len(df)} filas).")
            return df, None
        except Exception as e:
            error_msg = (
                f"No se pudo cargar el archivo:\n{ruta}\n\nError: {e}\n\n"
                "Posibles causas:\n"
                "- El archivo está siendo usado por otro programa.\n"
                "- No tiene instalado 'openpyxl' para .xlsx (o 'xlrd' para .xls).\n"
                "- El archivo está corrupto o en formato no soportado."
            )
            logger.exception(f"Error inesperado al cargar archivo Excel: {ruta}")
            return None, error_msg


# --- Clase MotorBusqueda (Fusionada y Mejorada) ---
class MotorBusqueda:
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[Path] = None
        self.archivo_descripcion_actual: Optional[Path] = None

        self.indices_columnas_busqueda_dic: List[int] = (
            indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else []
        )
        logger.info(
            f"MotorBusqueda inicializado. Índices búsqueda diccionario: {self.indices_columnas_busqueda_dic or 'Todas las de texto'}"
        )

        self.patron_comparacion_compilado = re.compile(
            r"^([<>]=?)(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ]+)?(.*)$"
        )
        self.patron_rango_compilado = re.compile(
            r"^(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ]+)?$"
        )
        self.patron_negacion_compilado = re.compile(r"^#(.+)$")
        self.patron_num_unidad_df = re.compile(
            r"(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ]+)?"
        )
        self.extractor_magnitud = ExtractorMagnitud()

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        """ Carga y valida el Excel del diccionario. Retorna (exito, mensaje_error_o_none). """
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            return False, error_msg_carga or "Error desconocido al cargar diccionario."

        self.datos_diccionario = df_cargado
        self.archivo_diccionario_actual = ruta
        
        valido, msg_val_cols = self._validar_columnas_df(
            self.datos_diccionario, self.indices_columnas_busqueda_dic, "diccionario"
        )
        if not valido:
            logger.warning(
                f"Validación de columnas del diccionario fallida. Carga invalidada. Causa: {msg_val_cols}"
            )
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            return False, msg_val_cols or "Validación de columnas del diccionario fallida."
        return True, None

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        """ Carga el Excel de descripciones. Retorna (exito, mensaje_error_o_none). """
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_descripcion = None
            self.archivo_descripcion_actual = None
            return False, error_msg_carga or "Error desconocido al cargar descripciones."

        self.datos_descripcion = df_cargado
        self.archivo_descripcion_actual = ruta
        # Podríamos añadir validación de columnas para descripciones si es necesario en el futuro
        return True, None

    def _validar_columnas_df(
        self, df: Optional[pd.DataFrame], indices_cfg: List[int], nombre_df_log: str
    ) -> Tuple[bool, Optional[str]]:
        """ Valida si los índices configurados son aplicables al DataFrame. Retorna (valido, mensaje_error_o_none). """
        if df is None:
            msg = f"DataFrame '{nombre_df_log}' es None, no se puede validar."
            logger.error(msg)
            return False, msg
        
        num_cols_df = len(df.columns)

        if not indices_cfg or indices_cfg == [-1]: # Modo "todas las de texto" o "todas"
            if num_cols_df == 0:
                msg = f"El archivo del {nombre_df_log} está vacío o no contiene columnas (modo 'todas')."
                logger.error(msg)
                return False, msg
            return True, None # Válido si tiene columnas

        if not all(isinstance(idx, int) and idx >= 0 for idx in indices_cfg):
            msg = f"Configuración de índices para {nombre_df_log} inválida: {indices_cfg}. Deben ser enteros no negativos."
            logger.error(msg)
            return False, msg

        max_indice_requerido = max(indices_cfg) if indices_cfg else -1

        if num_cols_df == 0:
            msg = f"El {nombre_df_log} no tiene columnas."
            logger.error(msg)
            return False, msg
        elif max_indice_requerido >= num_cols_df:
            msg = (
                f"El {nombre_df_log} necesita al menos {max_indice_requerido + 1} columnas "
                f"para los índices configurados ({indices_cfg}), pero solo tiene {num_cols_df}."
            )
            logger.error(msg)
            return False, msg
        return True, None

    def _obtener_nombres_columnas_busqueda_df(
        self, df: Optional[pd.DataFrame], indices_cfg: List[int], nombre_df_log: str
    ) -> Tuple[Optional[List[str]], Optional[str]]:
        """ Obtiene los nombres de las columnas para búsqueda. Retorna (lista_columnas_o_none, mensaje_error_o_none). """
        if df is None:
            msg = f"Intentando obtener nombres de columnas de un DataFrame ({nombre_df_log}) que es None."
            logger.error(msg)
            return None, msg

        columnas_disponibles = df.columns
        num_cols_df = len(columnas_disponibles)

        if not indices_cfg or indices_cfg == [-1]: # Modo "todas las de texto"
            cols_texto_obj = [
                col for col in df.columns
                if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])
            ]
            if cols_texto_obj:
                logger.info(
                    f"Buscando en columnas de texto/object (detectadas) del {nombre_df_log}: {cols_texto_obj}"
                )
                return cols_texto_obj, None
            elif num_cols_df > 0: # Si no hay de texto pero hay otras, usar todas
                logger.warning(
                    f"No se encontraron columnas de texto/object en {nombre_df_log}. Se usarán todas las {num_cols_df} columnas."
                )
                return list(df.columns), None
            else:
                msg = f"El DataFrame del {nombre_df_log} no tiene columnas."
                logger.error(msg)
                return None, msg

        nombres_columnas_seleccionadas = []
        indices_validos_usados = []
        for indice in indices_cfg:
            if 0 <= indice < num_cols_df:
                nombres_columnas_seleccionadas.append(columnas_disponibles[indice])
                indices_validos_usados.append(indice)
            else:
                logger.warning(
                    f"Índice {indice} para {nombre_df_log} es inválido o fuera de rango (0 a {num_cols_df-1}). Ignorado."
                )

        if not nombres_columnas_seleccionadas:
            msg = f"No se encontraron columnas válidas en {nombre_df_log} con los índices configurados: {indices_cfg}"
            logger.error(msg)
            return None, msg

        logger.debug(
            f"Se buscará en columnas del {nombre_df_log}: {nombres_columnas_seleccionadas} (índices: {indices_validos_usados})"
        )
        return nombres_columnas_seleccionadas, None

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        # (Lógica del primordial, que ya era bastante completa)
        palabras_analizadas = []
        for term_orig_bruto in terminos_brutos:
            term_orig = str(term_orig_bruto) 
            term = term_orig.strip()
            if not term: continue

            item_analizado: Dict[str, Any] = {'original': term_orig, 'negate': False}
            match_neg = self.patron_negacion_compilado.match(term)
            if match_neg:
                item_analizado['negate'] = True
                term = match_neg.group(1).strip()
                if not term: continue

            match_comp = self.patron_comparacion_compilado.match(term)
            match_range = self.patron_rango_compilado.match(term)

            if match_comp:
                op, v_str, unidad_str, _ = match_comp.groups()
                v_num = self._parse_numero(v_str)
                if v_num is not None:
                    op_map = {'>': 'gt', '<': 'lt', '>=': 'ge', '<=': 'le'}
                    item_analizado.update({
                        'tipo': op_map[op],
                        'valor': v_num,
                        'unidad_busqueda': self.extractor_magnitud._normalizar_texto(unidad_str.strip()) if unidad_str else None
                    })
                else: # Si no se pudo parsear como número, tratar como string
                    item_analizado.update({'tipo': 'str', 'valor': term.upper()})
            elif match_range:
                v1_str, v2_str, unidad_rango_str = match_range.groups()
                v1, v2 = self._parse_numero(v1_str), self._parse_numero(v2_str)
                if v1 is not None and v2 is not None:
                    item_analizado.update({
                        'tipo': 'range',
                        'valor': sorted([v1, v2]), # Guardar siempre [min, max]
                        'unidad_busqueda': self.extractor_magnitud._normalizar_texto(unidad_rango_str.strip()) if unidad_rango_str else None
                    })
                else: # Si no se pudo parsear como rango numérico, tratar como string
                    item_analizado.update({'tipo': 'str', 'valor': term.upper()})
            else: # Texto simple
                item_analizado.update({'tipo': 'str', 'valor': term.upper()})
            palabras_analizadas.append(item_analizado)
        logger.debug(f"Términos analizados (motor): {palabras_analizadas}")
        return palabras_analizadas

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        if not isinstance(num_str, (str, int, float)): return None
        try:
            return float(str(num_str).replace(',', '.'))
        except ValueError:
            return None

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols_a_buscar: List[str], termino_analizado: Dict[str, Any]) -> pd.Series:
        # (Lógica del primordial, que ya es funcional y maneja la iteración por celda para numéricos)
        mascara_total_subtermino = pd.Series(False, index=df.index)
        tipo_sub = termino_analizado['tipo']
        valor_sub = termino_analizado['valor'] 
        unidad_sub_requerida = termino_analizado.get('unidad_busqueda')
        es_negado = termino_analizado.get('negate', False)

        for col_nombre in cols_a_buscar:
            if col_nombre not in df.columns:
                logger.warning(f"Columna '{col_nombre}' no encontrada en el DataFrame al generar máscara. Saltando.")
                continue
            col_series = df[col_nombre]
            mascara_col_actual = pd.Series(False, index=df.index)

            if tipo_sub in ['gt', 'lt', 'ge', 'le', 'range']:
                for idx, valor_celda in col_series.items(): # Iterar sobre cada celda
                    if pd.isna(valor_celda) or str(valor_celda).strip() == "": continue
                    # Buscar todos los números con unidad en la celda actual
                    for match_celda in self.patron_num_unidad_df.finditer(str(valor_celda)):
                        try:
                            num_celda_val = float(match_celda.group(1).replace(',', '.'))
                            unidad_celda_val = self.extractor_magnitud._normalizar_texto(match_celda.group(2)) if match_celda.group(2) else None
                            
                            if unidad_sub_requerida and unidad_celda_val != unidad_sub_requerida:
                                continue # La unidad no coincide, probar siguiente match en la celda
                            
                            cond_ok = False
                            if tipo_sub == 'gt' and num_celda_val > valor_sub: cond_ok = True
                            elif tipo_sub == 'lt' and num_celda_val < valor_sub: cond_ok = True
                            elif tipo_sub == 'ge' and num_celda_val >= valor_sub: cond_ok = True
                            elif tipo_sub == 'le' and num_celda_val <= valor_sub: cond_ok = True
                            elif tipo_sub == 'range' and valor_sub[0] <= num_celda_val <= valor_sub[1]: cond_ok = True
                            
                            if cond_ok:
                                mascara_col_actual.at[idx] = True
                                break # Coincidencia encontrada en esta celda para esta columna
                        except ValueError: # Si el número parseado de la celda no es válido
                            continue
                    if mascara_col_actual.at[idx]: continue # Si ya se marcó, pasar a la siguiente fila del df
            elif tipo_sub == 'str':
                termino_regex_escapado = r"\b" + re.escape(str(valor_sub)) + r"\b" # valor_sub ya es upper
                try:
                    # Convertir toda la columna a string, luego a mayúsculas y buscar
                    mascara_col_actual = col_series.astype(str).str.upper().str.contains(termino_regex_escapado, regex=True, na=False)
                except Exception as e_conv_str:
                    logger.warning(f"No se pudo convertir/buscar string en columna '{col_nombre}': {e_conv_str}")
            
            mascara_total_subtermino |= mascara_col_actual.fillna(False)
        
        return ~mascara_total_subtermino if es_negado else mascara_total_subtermino

    def _aplicar_mascara_combinada(self, df: pd.DataFrame, cols_a_buscar: List[str], terminos_analizados: List[Dict[str, Any]], op_principal: str) -> pd.Series:
        # (Lógica del primordial)
        if df is None or df.empty or not cols_a_buscar or not terminos_analizados:
            return pd.Series(False, index=df.index if df is not None else None)

        cols_ok_en_df = [c for c in cols_a_buscar if c in df.columns]
        if not cols_ok_en_df:
            logger.error(f"Ninguna columna de búsqueda ({cols_a_buscar}) existe en el DataFrame provisto.")
            return pd.Series(False, index=df.index)

        op_es_and = op_principal.upper() == 'AND'
        mascara_final_combinada = pd.Series(True, index=df.index) if op_es_and else pd.Series(False, index=df.index)

        if not terminos_analizados: 
             return mascara_final_combinada

        for termino_individual_analizado in terminos_analizados:
            mascara_este_termino = self._generar_mascara_para_un_termino(df, cols_ok_en_df, termino_individual_analizado)
            if op_es_and:
                mascara_final_combinada &= mascara_este_termino
            else: # OR
                mascara_final_combinada |= mascara_este_termino
        
        return mascara_final_combinada

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, cols_nombres: List[str]) -> Set[str]:
        # (Lógica del primordial)
        terminos_clave: Set[str] = set()
        if df_coincidencias is None or df_coincidencias.empty or not cols_nombres:
            return terminos_clave

        columnas_validas_en_df = [c for c in cols_nombres if c in df_coincidencias.columns]
        if not columnas_validas_en_df:
            logger.warning("Ninguna de las columnas configuradas para el diccionario existe en el DataFrame de coincidencias.")
            return terminos_clave

        for col_nombre in columnas_validas_en_df:
            try:
                nuevos_terminos = df_coincidencias[col_nombre].dropna().astype(str).str.upper().unique()
                terminos_clave.update(t for t in nuevos_terminos if t and not t.isspace())
            except Exception as e:
                logger.warning(f"Error extrayendo términos clave de la columna '{col_nombre}' del diccionario: {e}")
        
        logger.info(f"Se extrajeron {len(terminos_clave)} términos clave del diccionario.")
        logger.debug(f"Términos clave extraídos (muestra): {list(terminos_clave)[:10]}...")
        return terminos_clave

    def _buscar_terminos_en_descripciones(self, df_desc_a_buscar: pd.DataFrame, terminos_clave: Set[str], require_all_terminos_clave: bool = False) -> pd.DataFrame:
        # (Lógica del primordial, con case=False por si los términos clave no fueran ya upper)
        if df_desc_a_buscar is None or df_desc_a_buscar.empty:
            logger.warning("DataFrame de descripciones (para búsqueda de términos clave) no provisto o vacío.")
            return pd.DataFrame(columns=(df_desc_a_buscar.columns if df_desc_a_buscar is not None else []))
        
        if not terminos_clave:
            logger.info("No hay términos clave para buscar en las descripciones.")
            return pd.DataFrame(columns=df_desc_a_buscar.columns)

        logger.info(f"Buscando {len(terminos_clave)} términos clave en {len(df_desc_a_buscar)} descripciones (require_all={require_all_terminos_clave}).")
        
        # Combinar todas las columnas de descripción en una sola serie de texto por fila, en mayúsculas
        texto_filas_desc = df_desc_a_buscar.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
        
        # Escapar términos clave (que ya deberían estar en mayúsculas desde _extraer_terminos_diccionario)
        terminos_escapados = [r"\b" + re.escape(t) + r"\b" for t in terminos_clave if t]
        if not terminos_escapados:
            return pd.DataFrame(columns=df_desc_a_buscar.columns)

        mascara_final_desc: pd.Series
        if require_all_terminos_clave:
            try:
                mascara_final_desc = pd.Series(True, index=texto_filas_desc.index)
                for pat in terminos_escapados: # Los patrones son los términos clave escapados
                    mascara_final_desc &= texto_filas_desc.str.contains(pat, regex=True, case=False, na=False) # case=False es seguro
            except Exception as e:
                logger.exception(f"Error durante la búsqueda 'require_all' en descripciones: {e}")
                return pd.DataFrame(columns=df_desc_a_buscar.columns)
        else: # require_any (OR)
            patron_or_regex = '|'.join(terminos_escapados)
            try:
                mascara_final_desc = texto_filas_desc.str.contains(patron_or_regex, regex=True, case=False, na=False)
            except Exception as e:
                logger.exception(f"Error durante la búsqueda 'OR' en descripciones: {e}")
                return pd.DataFrame(columns=df_desc_a_buscar.columns)

        resultados_desc = df_desc_a_buscar[mascara_final_desc]
        logger.info(f"Búsqueda de términos clave en descripciones completada. Resultados: {len(resultados_desc)} filas.")
        return resultados_desc

    def buscar(
        self,
        df_dic_para_busqueda: Optional[pd.DataFrame], 
        df_desc_para_busqueda: Optional[pd.DataFrame],
        terminos_analizados_usuario: List[Dict[str, Any]],
        operador_principal_usuario: str,
        buscar_via_diccionario_flag: bool,
    ) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[str]]:
        """ Lógica de búsqueda principal. Retorna (resultados_finales, origen, fcds_si_aplica, error_msg_motor). """
        logger.info(
            f"Motor.buscar: op='{operador_principal_usuario}', via_dicc={buscar_via_diccionario_flag}, terminos={terminos_analizados_usuario}"
        )

        fcds_obtenidos: Optional[pd.DataFrame] = None
        # error_motor_msg: Optional[str] = None # No se usa directamente para el retorno, OrigenResultados lo maneja

        if buscar_via_diccionario_flag:
            if df_dic_para_busqueda is None:
                return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, "Diccionario no cargado para búsqueda."
            
            cols_dic, err_cols_dic = self._obtener_nombres_columnas_busqueda_df(
                df_dic_para_busqueda, self.indices_columnas_busqueda_dic, "diccionario (búsqueda)"
            )
            if not cols_dic:
                return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, err_cols_dic or "Error en configuración de columnas del diccionario."

            try:
                mascara_en_dic = self._aplicar_mascara_combinada(
                    df_dic_para_busqueda, cols_dic, terminos_analizados_usuario, operador_principal_usuario
                )
                fcds_obtenidos = df_dic_para_busqueda[mascara_en_dic].copy()
                logger.info(f"FCDs encontrados: {len(fcds_obtenidos)}")
            except Exception as e_motor_dic:
                logger.exception("Error en motor al buscar en diccionario.")
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, f"Error interno buscando en diccionario: {e_motor_dic}"


            if df_desc_para_busqueda is None: # Necesario para el siguiente paso
                return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos, "Descripciones no cargadas para completar búsqueda vía diccionario."

            if fcds_obtenidos.empty: # No se encontraron FCDs
                return (pd.DataFrame(columns=df_desc_para_busqueda.columns), # Resultados finales vacíos
                        OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, # Indica que no hubo FCDs
                        fcds_obtenidos, None) 

            terminos_extraidos_de_fcds = self._extraer_terminos_diccionario(fcds_obtenidos, cols_dic)
            
            if not terminos_extraidos_de_fcds:
                return (pd.DataFrame(columns=df_desc_para_busqueda.columns),
                        OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, # FCDs existían, pero no se extrajeron términos
                        fcds_obtenidos, None)

            try:
                # require_all_terminos_clave=False: si FCD da "ROUTER CISCO", busca "ROUTER" OR "CISCO" en desc.
                resultados_desc_via_dic = self._buscar_terminos_en_descripciones(
                    df_desc_para_busqueda, terminos_extraidos_de_fcds, require_all_terminos_clave=False  
                )
            except Exception as e_motor_desc_via_dic:
                logger.exception("Error en motor al buscar términos de FCDs en descripciones.")
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos, f"Error interno buscando términos de FCDs en descripciones: {e_motor_desc_via_dic}"

            if resultados_desc_via_dic.empty:
                return (pd.DataFrame(columns=df_desc_para_busqueda.columns), # Vacío pero con estructura
                        OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, # FCDs existían, términos se extrajeron, pero no hubo match en desc
                        fcds_obtenidos, None)
            else:
                return (resultados_desc_via_dic,
                        OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
                        fcds_obtenidos, None)
        else: # Búsqueda directa en descripciones
            if df_desc_para_busqueda is None:
                return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, "Descripciones no cargadas para búsqueda directa."

            cols_desc, err_cols_desc = self._obtener_nombres_columnas_busqueda_df(
                df_desc_para_busqueda, [], "descripciones (búsqueda directa)" # Vacío para todas las de texto
            )
            if not cols_desc:
                return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, err_cols_desc or "Error en configuración de columnas de descripción."
            
            try:
                mascara_en_desc_directa = self._aplicar_mascara_combinada(
                    df_desc_para_busqueda, cols_desc, terminos_analizados_usuario, operador_principal_usuario
                )
                resultados_directos = df_desc_para_busqueda[mascara_en_desc_directa].copy()
                logger.info(f"Resultados directos en descripción: {len(resultados_directos)}")
            except Exception as e_motor_desc_directa:
                logger.exception("Error en motor al buscar directamente en descripciones.")
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, f"Error interno buscando directamente en descripciones: {e_motor_desc_directa}"


            if resultados_directos.empty:
                return (pd.DataFrame(columns=df_desc_para_busqueda.columns), # Vacío pero con estructura
                        OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None)
            else:
                return (resultados_directos,
                        OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None)


# --- Clase InterfazGrafica (Base primordial, con adaptaciones) ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE = "config_buscador_FUSION_v2.json" # Nuevo nombre para la config fusionada

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado (Fusionado)") # Título actualizado
        self.geometry("1250x800")

        self.config = self._cargar_configuracion()
        indices_cfg = self.config.get("indices_columnas_busqueda_dic", [])
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg)

        self.resultados_actuales: Optional[pd.DataFrame] = None # Contendrá los resultados finales mostrados en tabla_resultados
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        self.ultimo_termino_buscado: Optional[str] = None
        self.reglas_guardadas: List[Dict[str, Any]] = []
        
        # Variables de estado para los datos de la última búsqueda (para salvar reglas)
        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None # Coincidencias del diccionario
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None # Resultados de descripción (vía dic o directa)
        
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO
        self.color_fila_par = "white"; self.color_fila_impar = "#f0f0f0"
        self.op_buttons: Dict[str, ttk.Button] = {} # Para gestionar botones de operadores

        self._configurar_estilo_ttk()
        self._crear_widgets() # Usará la gestión de botones de 'refactor'
        self._configurar_grid()
        self._configurar_eventos()
        self._configurar_tags_treeview()
        self._configurar_orden_tabla(self.tabla_resultados)
        self._configurar_orden_tabla(self.tabla_diccionario)

        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones.")
        self._deshabilitar_botones_operadores()
        self._actualizar_botones_estado_general()
        logger.info("Interfaz Gráfica (Fusionada) inicializada.")

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str):
        self._actualizar_estado_botones_operadores()

    def _cargar_configuracion(self) -> Dict:
        # (Lógica de primordial, usando self.CONFIG_FILE)
        config = {}
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f: config = json.load(f)
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE}")
            except Exception as e:
                logger.error(f"Error al cargar config: {e}")
                # No messagebox aquí para no interrumpir inicio
        else:
            logger.info(f"Archivo de configuración '{self.CONFIG_FILE}' no encontrado. Se creará al cerrar.")
        
        last_dic_path_str = config.get("last_dic_path")
        config["last_dic_path"] = str(Path(last_dic_path_str)) if last_dic_path_str else None
        last_desc_path_str = config.get("last_desc_path")
        config["last_desc_path"] = str(Path(last_desc_path_str)) if last_desc_path_str else None
        config.setdefault("indices_columnas_busqueda_dic", [])
        return config

    def _guardar_configuracion(self):
        # (Lógica de primordial, usando self.CONFIG_FILE)
        self.config["last_dic_path"] = str(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else None
        self.config["last_desc_path"] = str(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else None
        self.config["indices_columnas_busqueda_dic"] = self.motor.indices_columnas_busqueda_dic
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            logger.info(f"Configuración guardada en: {self.CONFIG_FILE}")
        except Exception as e:
            logger.error(f"Error al guardar config: {e}")
            messagebox.showerror("Error Configuración", f"No se pudo guardar config:\n{e}")
    
    def _configurar_estilo_ttk(self):
        # (Lógica de primordial)
        style = ttk.Style(self); themes = style.theme_names(); os_name = platform.system()
        prefs = {"Windows":["vista","xpnative","clam"],"Darwin":["aqua","clam"],"Linux":["clam","alt","default"]}
        theme_to_use = next((t for t in prefs.get(os_name, ["clam","default"]) if t in themes), None)
        if not theme_to_use:
            theme_to_use = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))
        if theme_to_use:
            logger.info(f"Aplicando tema TTK: {theme_to_use}")
            try: 
                style.theme_use(theme_to_use)
                # Estilo para botones de operadores del 'refactor'
                style.configure("Operator.TButton", padding=(2, 1), font=('TkDefaultFont', 9)) # width no parece tener efecto en todos los temas
            except tk.TclError as e: logger.warning(f"No se pudo aplicar tema '{theme_to_use}': {e}.")
        else: logger.warning("No se encontró tema TTK disponible.")

    def _crear_widgets(self):
        # (Lógica de primordial, pero con la creación de botones de operadores de 'refactor')
        self.marco_controles = ttk.LabelFrame(self, text="Controles")
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.lbl_dic_cargado = ttk.Label(self.marco_controles, text="Dic: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion)
        self.lbl_desc_cargado = ttk.Label(self.marco_controles, text="Desc: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)

        self.frame_ops = ttk.Frame(self.marco_controles)
        op_buttons_defs = [ # Definiciones de los botones de operadores
            ("+", "+"), ("|", "|"), ("#", "#"), (">", ">"), 
            ("<", "<"), ("≥", ">="), ("≤", "<="), ("-", "-")
        ]
        for i, (text, op_val) in enumerate(op_buttons_defs):
            btn = ttk.Button(
                self.frame_ops, text=text, 
                command=lambda op=op_val: self._insertar_operador_validado(op),
                style="Operator.TButton", width=3 # width=3 para un poco más de espacio
            )
            btn.grid(row=0, column=i, padx=1, pady=1, sticky="nsew")
            self.op_buttons[op_val] = btn # Guardar referencia

        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50, textvariable=self.texto_busqueda_var)
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda)
        self.btn_salvar_regla = ttk.Button(self.marco_controles, text="Salvar Regla", command=self._salvar_regla_actual)
        self.btn_ayuda = ttk.Button(self.marco_controles, text="?", command=self._mostrar_ayuda, width=3)
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar", command=self._exportar_resultados)

        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario:")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados / Descripciones:")
        
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings", height=8)
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)

        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings")
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)

        self.barra_estado = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W, borderwidth=1)
        self._actualizar_etiquetas_archivos()

    def _configurar_grid(self):
        # (Lógica de primordial, ajustada para el frame_ops)
        self.grid_rowconfigure(2, weight=1); self.grid_rowconfigure(4, weight=3)
        self.grid_columnconfigure(0, weight=1)
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        self.marco_controles.grid_columnconfigure(1, weight=1) 
        self.marco_controles.grid_columnconfigure(3, weight=1) 

        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew")
        
        self.frame_ops.grid(row=1, column=0, columnspan=4, padx=5, pady=(5,0), sticky="ew")
        for i in range(len(self.op_buttons)): self.frame_ops.grid_columnconfigure(i, weight=1) # Que los botones se expandan

        self.entrada_busqueda.grid(row=2, column=0, columnspan=2, padx=5, pady=(0,5), sticky="ew")
        self.btn_buscar.grid(row=2, column=2, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_salvar_regla.grid(row=2, column=3, padx=(2,0), pady=(0,5), sticky="w") 
        self.btn_ayuda.grid(row=2, column=4, padx=(2,0), pady=(0,5), sticky="w")     
        self.btn_exportar.grid(row=2, column=5, padx=(10, 5), pady=(0,5), sticky="e") 

        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0))
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1); self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew"); self.scrolly_diccionario.grid(row=0, column=1, sticky="ns"); self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")
        
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1); self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew"); self.scrolly_resultados.grid(row=0, column=1, sticky="ns"); self.scrollx_resultados.grid(row=1, column=0, sticky="ew")
        
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0))

    def _configurar_eventos(self):
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _actualizar_estado(self, mensaje: str):
        self.barra_estado.config(text=mensaje)
        logger.info(f"Estado UI: {mensaje}")
        self.update_idletasks()

    def _mostrar_ayuda(self):
        # (Ayuda combinada/mejorada)
        ayuda = """Sintaxis de Búsqueda:
-------------------------------------
- Texto simple: Busca la palabra o frase (insensible a mayús/minús y acentos). Ej: `router cisco`
- Operadores Lógicos:
  * `término1 + término2`: Busca filas con AMBOS (AND). Ej: `tarjeta + 16 puertos`
  * `término1 | término2` (o `/`): Busca filas con AL MENOS UNO (OR). Ej: `modulo | SFP`
- Comparaciones numéricas (unidad opcional, si se usa debe coincidir):
  * `>num[UNIDAD]`: Mayor. Ej: `>1000` o `>1000w`
  * `<num[UNIDAD]`: Menor. Ej: `<50` o `<50v`
  * `>=num[UNIDAD]` o `≥num[UNIDAD]`: Mayor o igual. Ej: `>=48a`
  * `<=num[UNIDAD]` o `≤num[UNIDAD]`: Menor o igual. Ej: `<=10.5w`
- Rangos numéricos (unidad opcional, ambos extremos incluidos):
  * `num1-num2[UNIDAD]`: Entre num1 y num2. Ej: `10-20` o `50-100V`
- Negación (excluir):
  * `#término`: `término` puede ser texto, comparación o rango.
    Ej: `switch + #gestionable` (switch Y NO gestionable)
    Ej: `tarjeta + #>8 puertos` (tarjeta Y NO mayor que 8 puertos)

Modo de Búsqueda:
1. El término se busca primero en el Diccionario.
2. Si hay coincidencias en el Diccionario (FCDs):
   - Se extraen todos los textos de las columnas de búsqueda de esos FCDs.
   - Estos textos extraídos se buscan (con OR) en las Descripciones.
3. Si NO hay coincidencias en el Diccionario O la búsqueda vía diccionario no da resultados en descripciones:
   - Se preguntará si desea buscar el término original directamente en las Descripciones.
   - Esta búsqueda directa usará la misma lógica de operadores.
4. Búsqueda vacía: Muestra todas las descripciones.
"""
        messagebox.showinfo("Ayuda - Sintaxis de Búsqueda", ayuda)

    def _configurar_tags_treeview(self):
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par)
            tabla.tag_configure('impar', background=self.color_fila_impar)

    def _configurar_orden_tabla(self, tabla: ttk.Treeview):
        cols = tabla["columns"]
        if cols: # Solo si hay columnas definidas
            for col in cols:
                tabla.heading(col, text=str(col), anchor=tk.W,
                              command=lambda c=col, t=tabla: self._ordenar_columna(t, c, False))

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool):
        # (Lógica de primordial, adaptada para DataFrame)
        df_para_ordenar = None
        if tabla == self.tabla_diccionario:
            df_para_ordenar = self.motor.datos_diccionario
        elif tabla == self.tabla_resultados:
            df_para_ordenar = self.resultados_actuales # Usar los resultados actualmente mostrados
        
        if df_para_ordenar is None or df_para_ordenar.empty or col not in df_para_ordenar.columns:
            logging.debug(f"No se puede ordenar la tabla por columna '{col}'. Datos no disponibles o columna inexistente.")
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, not reverse))
            return

        logging.info(f"Ordenando tabla por columna '{col}', descendente={reverse}")
        try:
            col_to_sort_by = df_para_ordenar[col]
            if pd.api.types.is_numeric_dtype(col_to_sort_by.infer_objects().dtype):
                col_numeric = pd.to_numeric(col_to_sort_by, errors='coerce')
                # Usar .loc para reordenar el DataFrame original basado en los índices ordenados
                df_ordenado = df_para_ordenar.loc[col_numeric.sort_values(ascending=not reverse, na_position='last').index]
            else: # Ordenar como string (case-insensitive)
                df_ordenado = df_para_ordenar.sort_values(
                    by=col, 
                    ascending=not reverse, 
                    na_position='last', 
                    key=lambda x: x.astype(str).str.lower() if pd.api.types.is_string_dtype(x) or pd.api.types.is_object_dtype(x) else x
                )
            
            if tabla == self.tabla_diccionario:
                self.motor.datos_diccionario = df_ordenado 
                self._actualizar_tabla(tabla, df_ordenado, limite_filas=100) 
            elif tabla == self.tabla_resultados:
                self.resultados_actuales = df_ordenado
                self._actualizar_tabla(tabla, df_ordenado)
            
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, not reverse))
            self._actualizar_estado(f"Tabla ordenada por '{col}' ({'Asc' if not reverse else 'Desc'}).")
        except Exception as e:
            logging.exception(f"Error al intentar ordenar por columna '{col}'")
            messagebox.showerror("Error al Ordenar", f"No se pudo ordenar por '{col}':\n{e}")
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, False))

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None, columnas_a_mostrar: Optional[List[str]] = None):
        # (Lógica de primordial, con mejoras de 'refactor' para ancho de columna y reconfiguración de orden)
        is_diccionario = tabla == self.tabla_diccionario
        logger.debug(f"Actualizando tabla {'Diccionario' if is_diccionario else 'Resultados'}.")
        try:
            for i in tabla.get_children(): tabla.delete(i)
        except tk.TclError as e: logger.warning(f"Error Tcl al limpiar tabla: {e}"); pass
        tabla["columns"] = ()

        if datos is None or datos.empty:
            logger.debug("No hay datos para mostrar en la tabla.")
            self._configurar_orden_tabla(tabla) # Reconfigurar orden en tabla vacía para que las cabeceras sigan siendo clickables
            return

        datos_a_usar = datos
        cols_df = list(datos_a_usar.columns)
        cols_finales = [c for c in (columnas_a_mostrar or cols_df) if c in cols_df] or cols_df

        if not cols_finales:
            logger.warning("DataFrame no tiene columnas para mostrar o columnas seleccionadas no existen.")
            self._configurar_orden_tabla(tabla)
            return

        df_para_mostrar_vista = datos_a_usar[cols_finales]
        tabla["columns"] = tuple(cols_finales)

        for col in cols_finales:
            tabla.heading(col, text=str(col), anchor=tk.W) # El command se setea en _configurar_orden_tabla
            try:
                col_as_str = df_para_mostrar_vista[col].dropna().astype(str)
                ancho_contenido = col_as_str.str.len().max() if not col_as_str.empty else 0
                ancho_cabecera = len(str(col))
                ancho = max(70, min(int(max(ancho_cabecera * 8, ancho_contenido * 6.5) + 25), 400))
                tabla.column(col, anchor=tk.W, width=ancho, minwidth=70)
            except Exception: tabla.column(col, anchor=tk.W, width=100, minwidth=50)

        num_filas_a_iterar = limite_filas if is_diccionario and limite_filas is not None else len(df_para_mostrar_vista)
        df_iterar = df_para_mostrar_vista.head(num_filas_a_iterar)

        for i, (_, row) in enumerate(df_iterar.iterrows()):
            vals = [str(v) if pd.notna(v) else "" for v in row.values]
            tag = 'par' if i % 2 == 0 else 'impar'
            try:
                tabla.insert("", "end", values=vals, tags=(tag,))
            except tk.TclError: # Fallback para caracteres problemáticos
                try:
                    vals_ascii = [v.encode('ascii', 'ignore').decode('ascii') for v in vals]
                    tabla.insert("", "end", values=vals_ascii, tags=(tag,))
                except Exception as e_inner: logger.error(f"Fallo el fallback ASCII para fila {i}: {e_inner}")
        
        self._configurar_orden_tabla(tabla) # Reconfigurar orden después de cambiar columnas/datos

    def _actualizar_etiquetas_archivos(self):
        # (Lógica de primordial)
        dic_path = self.motor.archivo_diccionario_actual
        desc_path = self.motor.archivo_descripcion_actual
        dic_name = dic_path.name if dic_path else "Ninguno"
        desc_name = desc_path.name if desc_path else "Ninguno"
        
        max_len_label = 25 
        dic_display = f"Dic: {dic_name}" if len(dic_name) <= max_len_label else f"Dic: ...{dic_name[-(max_len_label-4):]}"
        desc_display = f"Desc: {desc_name}" if len(desc_name) <= max_len_label else f"Desc: ...{desc_name[-(max_len_label-4):]}"
        
        self.lbl_dic_cargado.config(text=dic_display, foreground="green" if dic_path else "red")
        self.lbl_desc_cargado.config(text=desc_display, foreground="green" if desc_path else "red")

    def _actualizar_botones_estado_general(self):
        # (Lógica de 'refactor' adaptada)
        dic_cargado = self.motor.datos_diccionario is not None
        desc_cargado = self.motor.datos_descripcion is not None

        if dic_cargado or desc_cargado: # Habilitar operadores si al menos uno está cargado (para búsqueda directa)
            self._actualizar_estado_botones_operadores()
        else:
            self._deshabilitar_botones_operadores()

        self.btn_buscar['state'] = 'normal' if dic_cargado and desc_cargado else 'disabled'
        
        puede_salvar_algo = False
        if self.ultimo_termino_buscado and self.origen_principal_resultados != OrigenResultados.NINGUNO:
            if self.origen_principal_resultados.es_via_diccionario:
                if (self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty) or \
                   (self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and \
                    self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC):
                    puede_salvar_algo = True
            elif self.origen_principal_resultados.es_directo_descripcion or \
                 self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA: # Búsqueda vacía también se puede guardar
                 if self.desc_finales_de_ultima_busqueda is not None: 
                    puede_salvar_algo = True
        
        self.btn_salvar_regla['state'] = 'normal' if puede_salvar_algo else 'disabled'
        self.btn_exportar['state'] = 'normal' if self.reglas_guardadas else 'disabled'

    def _cargar_diccionario(self):
        # (Lógica de 'refactor' para manejo de error de carga)
        last_dir_str = self.config.get("last_dic_path")
        last_dir = str(Path(last_dir_str).parent) if last_dir_str and Path(last_dir_str).exists() else os.getcwd()
        
        ruta = filedialog.askopenfilename(
            title="Seleccionar Archivo Diccionario",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            initialdir=last_dir
        )
        if not ruta: logger.info("Carga de diccionario cancelada."); return

        nombre_archivo = Path(ruta).name
        self._actualizar_estado(f"Cargando diccionario: {nombre_archivo}...")
        self._actualizar_tabla(self.tabla_diccionario, None)
        self._actualizar_tabla(self.tabla_resultados, None) 
        self.resultados_actuales = None; self.fcds_de_ultima_busqueda = None; self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        exito_carga, msg_error_carga = self.motor.cargar_excel_diccionario(ruta)
        if exito_carga:
            self.config["last_dic_path"] = ruta 
            self._guardar_configuracion() 
            df_dic = self.motor.datos_diccionario
            if df_dic is not None:
                num_filas = len(df_dic)
                cols_busqueda_nombres, _ = self.motor._obtener_nombres_columnas_busqueda_df(df_dic, self.motor.indices_columnas_busqueda_dic, "diccionario (preview)")
                
                indices_str = ', '.join(map(str, self.motor.indices_columnas_busqueda_dic)) if self.motor.indices_columnas_busqueda_dic and self.motor.indices_columnas_busqueda_dic != [-1] else "Todas Texto"
                lbl_text_dic_preview = f"Vista Previa Diccionario ({num_filas} filas)"
                if cols_busqueda_nombres:
                    lbl_text_dic_preview = f"Vista Previa Diccionario (Cols: {', '.join(cols_busqueda_nombres)} - Índices: {indices_str})"
                self.lbl_tabla_diccionario.config(text=lbl_text_dic_preview)
                
                self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols_busqueda_nombres)
                self.title(f"Buscador - Dic: {nombre_archivo}")
                self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado.")
        else:
            self._actualizar_estado(f"Error al cargar diccionario: {msg_error_carga or 'Desconocido'}")
            messagebox.showerror("Error Carga Diccionario", msg_error_carga or "No se pudo cargar el diccionario.")
            self.title("Buscador Avanzado (Fusionado)")
        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _cargar_excel_descripcion(self):
        # (Lógica de 'refactor' para manejo de error de carga)
        last_dir_str = self.config.get("last_desc_path")
        last_dir = str(Path(last_dir_str).parent) if last_dir_str and Path(last_dir_str).exists() else os.getcwd()
        
        ruta = filedialog.askopenfilename(
            title="Seleccionar Archivo de Descripciones",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            initialdir=last_dir
        )
        if not ruta: logger.info("Carga de descripciones cancelada."); return

        nombre_archivo = Path(ruta).name
        self._actualizar_estado(f"Cargando descripciones: {nombre_archivo}...")
        self.resultados_actuales = None; self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        exito_carga, msg_error_carga = self.motor.cargar_excel_descripcion(ruta)
        if exito_carga:
            self.config["last_desc_path"] = ruta
            self._guardar_configuracion()
            df_desc = self.motor.datos_descripcion
            if df_desc is not None:
                num_filas = len(df_desc)
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando datos...")
                self._actualizar_tabla(self.tabla_resultados, df_desc) 
                dic_n_title = Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A"
                self.title(f"Buscador - Dic: {dic_n_title} | Desc: {nombre_archivo}")
        else:
            self._actualizar_estado(f"Error al cargar descripciones: {msg_error_carga or 'Desconocido'}")
            messagebox.showerror("Error Carga Descripciones", msg_error_carga or "No se pudieron cargar las descripciones.")
        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _parsear_termino_busqueda_inicial(self, termino_raw: str) -> Tuple[str, List[str]]:
        # (Lógica de primordial)
        termino_limpio = termino_raw.strip()
        op_principal = 'OR'; terminos_brutos = [] 
        if not termino_limpio: return op_principal, []

        if '+' in termino_limpio:
            op_principal = 'AND'
            terminos_brutos = [p.strip() for p in termino_limpio.split('+') if p.strip()]
        elif '|' in termino_limpio:
            op_principal = 'OR'
            terminos_brutos = [p.strip() for p in termino_limpio.split('|') if p.strip()]
        elif '/' in termino_limpio:
            op_principal = 'OR'
            terminos_brutos = [p.strip() for p in termino_limpio.split('/') if p.strip()]
        else:
            op_principal = 'AND' 
            terminos_brutos = [termino_limpio]

        if not any(terminos_brutos):
            logger.warning(f"Término '{termino_raw}' vacío tras parseo inicial."); return 'OR', []
        return op_principal, terminos_brutos

    def _ejecutar_busqueda(self):
        # (Lógica de 'refactor' adaptada para UI primordial)
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", "Cargue Diccionario y Descripciones.")
            return

        termino_busqueda_actual = self.texto_busqueda_var.get()
        self.ultimo_termino_buscado = termino_busqueda_actual

        self.resultados_actuales = None
        self._actualizar_tabla(self.tabla_resultados, None)
        self.fcds_de_ultima_busqueda = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        if not termino_busqueda_actual.strip(): # Búsqueda vacía
            logger.info("Búsqueda vacía. Mostrando todas las descripciones.")
            df_desc_all = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc_all)
            self.resultados_actuales = df_desc_all.copy() if df_desc_all is not None else pd.DataFrame()
            self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy()
            self.origen_principal_resultados = OrigenResultados.DIRECTO_DESCRIPCION_VACIA
            self._actualizar_estado(f"Mostrando todas las {len(df_desc_all) if df_desc_all is not None else 0} descripciones.")
            self._actualizar_botones_estado_general()
            return

        self._actualizar_estado(f"Buscando '{termino_busqueda_actual}'...")
        
        df_dic_motor = self.motor.datos_diccionario
        df_desc_motor = self.motor.datos_descripcion

        op_principal_usuario, terminos_brutos_usuario = self._parsear_termino_busqueda_inicial(termino_busqueda_actual)
        if not terminos_brutos_usuario:
            messagebox.showwarning("Término Inválido", "Término de búsqueda vacío o malformado.")
            self._actualizar_estado("Término de búsqueda inválido."); self._actualizar_botones_estado_general(); return
            
        terminos_analizados_usuario = self.motor._analizar_terminos(terminos_brutos_usuario)
        if not terminos_analizados_usuario:
            messagebox.showwarning("Término Inválido", f"No se pudieron analizar los términos en '{termino_busqueda_actual}'. Verifique sintaxis.")
            self._actualizar_estado(f"Análisis de '{termino_busqueda_actual}' fallido."); self._actualizar_botones_estado_general(); return

        # Intenta búsqueda vía diccionario primero
        res_df, origen_res, fcds_res, err_msg_motor = self.motor.buscar(
            df_dic_motor, df_desc_motor, terminos_analizados_usuario, 
            op_principal_usuario, buscar_via_diccionario_flag=True
        )

        self.fcds_de_ultima_busqueda = fcds_res # Guardar FCDs siempre, incluso si son None o vacíos
        self.origen_principal_resultados = origen_res # Guardar el origen de esta primera fase

        if err_msg_motor: # Error operacional del motor
            messagebox.showerror("Error de Búsqueda (Motor)", f"Error interno del motor: {err_msg_motor}")
            self._actualizar_estado(f"Error en motor: {err_msg_motor}")
            self.resultados_actuales = pd.DataFrame(columns=df_desc_motor.columns if df_desc_motor is not None else [])
            self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy()
        elif origen_res.es_error_carga or origen_res.es_error_configuracion:
            # Estos errores deberían haber sido manejados al cargar los archivos.
            # Si el motor los devuelve aquí, es una doble comprobación o un error en el flujo.
            msg_shown = err_msg_motor or f"Error pre-búsqueda impidió la operación: {origen_res.name}"
            messagebox.showerror("Error de Búsqueda", msg_shown)
            self._actualizar_estado(msg_shown)
            self.resultados_actuales = pd.DataFrame(columns=df_desc_motor.columns if df_desc_motor is not None else [])
            self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy()
        elif origen_res == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC:
            self.resultados_actuales = res_df
            self.desc_finales_de_ultima_busqueda = res_df.copy() if res_df is not None else pd.DataFrame()
            self._actualizar_estado(f"'{termino_busqueda_actual}': {len(fcds_res) if fcds_res is not None else 0} en Dic, {len(res_df) if res_df is not None else 0} en Desc.")
        
        # Si la búsqueda vía diccionario no produjo resultados FINALES en descripción, o no hubo FCDs para empezar,
        # o no se pudieron extraer términos válidos de los FCDs, entonces preguntar para búsqueda directa.
        elif origen_res in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS] or \
             (fcds_res is not None and fcds_res.empty and origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC): # Caso: no hubo FCDs
            
            self.resultados_actuales = res_df # Esto será un DF vacío con columnas de descripción
            self.desc_finales_de_ultima_busqueda = res_df.copy() if res_df is not None else pd.DataFrame()

            msg_info_fcd = f"{len(fcds_res) if fcds_res is not None else 0} coincidencias en Diccionario"
            if fcds_res is not None and fcds_res.empty:
                 msg_info_fcd = "Ninguna coincidencia en Diccionario"

            if origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS:
                msg_desc_issue = "pero no se extrajeron términos válidos para buscar en descripciones."
                self._actualizar_estado(f"'{termino_busqueda_actual}': {msg_info_fcd}, sin términos válidos para Desc.")
            else: # VIA_DICCIONARIO_SIN_RESULTADOS_DESC (incluye caso sin FCDs)
                msg_desc_issue = "lo que no produjo resultados en Descripciones."
                self._actualizar_estado(f"'{termino_busqueda_actual}': {msg_info_fcd}, 0 en Desc.")

            if messagebox.askyesno(
                "Búsqueda Alternativa",
                f"{msg_info_fcd} para '{termino_busqueda_actual}', {msg_desc_issue}\n\n"
                f"¿Desea buscar '{termino_busqueda_actual}' directamente en las Descripciones?",
            ):
                self._actualizar_estado(f"Buscando directamente '{termino_busqueda_actual}' en descripciones...")
                # Reutilizar el mismo operador que se usó en la búsqueda inicial
                res_df_directo, origen_directo, _, err_msg_motor_directo = self.motor.buscar(
                    df_dic_motor, df_desc_motor, terminos_analizados_usuario,
                    op_principal_usuario, buscar_via_diccionario_flag=False
                )
                if err_msg_motor_directo:
                    messagebox.showerror("Error Búsqueda Directa", f"Error interno: {err_msg_motor_directo}")
                    self._actualizar_estado(f"Error búsqueda directa: {err_msg_motor_directo}")
                    # Mantener resultados_actuales como el DF vacío de la fase anterior
                else:
                    self.resultados_actuales = res_df_directo
                    self.desc_finales_de_ultima_busqueda = res_df_directo.copy() if res_df_directo is not None else pd.DataFrame()
                    self.origen_principal_resultados = origen_directo # Actualizar el origen principal
                    self.fcds_de_ultima_busqueda = None # La búsqueda final fue directa
                    num_rdd = len(self.resultados_actuales) if self.resultados_actuales is not None else 0
                    self._actualizar_estado(f"Búsqueda directa '{termino_busqueda_actual}': {num_rdd} resultados.")
                    if num_rdd == 0:
                        messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_busqueda_actual}' en búsqueda directa.")
            # Si no se hace búsqueda directa, los resultados_actuales y desc_finales ya son el DF vacío.
        
        # Asegurar que self.resultados_actuales y self.desc_finales_de_ultima_busqueda no sean None si la búsqueda falló o fue cancelada
        if self.resultados_actuales is None:
            cols = df_desc_motor.columns if df_desc_motor is not None else []
            self.resultados_actuales = pd.DataFrame(columns=cols)
        if self.desc_finales_de_ultima_busqueda is None:
            self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy()
            
        self._actualizar_tabla(self.tabla_resultados, self.resultados_actuales)
        self._actualizar_botones_estado_general()

        if df_dic_motor is not None and not df_dic_motor.empty:
            self._buscar_y_enfocar_en_preview()

    def _buscar_y_enfocar_en_preview(self):
        # (Lógica de 'refactor' para limpieza de término)
        df_completo_dic = self.motor.datos_diccionario
        if df_completo_dic is None or df_completo_dic.empty: return

        termino_buscar_raw = self.texto_busqueda_var.get()
        if not termino_buscar_raw.strip(): return
        
        _, terminos_brutos_preview = self._parsear_termino_busqueda_inicial(termino_buscar_raw)
        
        clean_terminos_preview = []
        for t_prev in terminos_brutos_preview:
            t_clean = t_prev.strip()
            if t_clean.startswith("#"): t_clean = t_clean[1:].strip() # Quitar negación
            
            # Intentar quitar patrones de comparación y rango para enfocar en el "núcleo" del término
            # Esto es heurístico y podría necesitar ajustes
            t_clean = re.sub(self.motor.patron_comparacion_compilado.pattern, lambda m: m.group(4).strip() or m.group(2), t_clean) 
            t_clean = re.sub(self.motor.patron_rango_compilado.pattern, lambda m: m.group(1), t_clean) # Toma el primer número del rango
            
            # Quitar unidades conocidas si es un término numérico simple (post-limpieza)
            match_num_unidad = self.motor.patron_num_unidad_df.fullmatch(t_clean)
            if match_num_unidad and match_num_unidad.group(2): # Tiene número y unidad
                norm_unidad = self.motor.extractor_magnitud._normalizar_texto(match_num_unidad.group(2))
                if self.motor.extractor_magnitud.magnitudes_normalizadas.get(norm_unidad):
                    t_clean = match_num_unidad.group(1) # Solo el número si la unidad es conocida

            if t_clean: clean_terminos_preview.append(t_clean)
        
        if not clean_terminos_preview: return
        termino_enfocar_preview_str = clean_terminos_preview[0].upper() 

        if not termino_enfocar_preview_str: return

        items_preview_ids = self.tabla_diccionario.get_children('')
        if not items_preview_ids: return
        
        logger.info(f"Intentando enfocar '{termino_enfocar_preview_str}' en vista previa del diccionario...")
        found_item_id = None
        for item_id in items_preview_ids:
            try:
                valores_fila_preview = self.tabla_diccionario.item(item_id, 'values')
                if any(termino_enfocar_preview_str in str(val_prev).upper() for val_prev in valores_fila_preview if val_prev is not None):
                    found_item_id = item_id; break
            except Exception as e_focus:
                logger.warning(f"Error procesando item {item_id} en preview (búsqueda treeview): {e_focus}"); continue
        
        if found_item_id:
            logger.info(f"Término '{termino_enfocar_preview_str}' enfocado en preview (item ID: {found_item_id}).")
            try:
                current_selection = self.tabla_diccionario.selection()
                if current_selection: self.tabla_diccionario.selection_remove(current_selection)
                self.tabla_diccionario.selection_set(found_item_id)
                self.tabla_diccionario.see(found_item_id)
                self.tabla_diccionario.focus(found_item_id) # Dar foco al item
            except Exception as e_tk_focus:
                logger.error(f"Error al enfocar item {found_item_id} en preview: {e_tk_focus}")
        else:
            logger.info(f"Término '{termino_enfocar_preview_str}' no encontrado/enfocado en vista previa del diccionario.")

    def _salvar_regla_actual(self):
        # (Lógica de 'refactor' usando las nuevas variables de estado)
        origen_nombre_actual = self.origen_principal_resultados.name
        logger.info(f"Intentando salvar regla. Origen: {origen_nombre_actual}, Último término: '{self.ultimo_termino_buscado}'")

        if not self.ultimo_termino_buscado:
            messagebox.showerror("Error", "No hay término de búsqueda para la regla.")
            return

        termino_original_regla = self.ultimo_termino_buscado
        op_regla, terminos_brutos_regla = self._parsear_termino_busqueda_inicial(termino_original_regla)
        terminos_analizados_regla = self.motor._analizar_terminos(terminos_brutos_regla)

        regla_a_guardar_base = {
            'termino_busqueda_original': termino_original_regla,
            'termino_busqueda_parseado': terminos_analizados_regla,
            'operador_principal': op_regla,
            'fuente_original_guardado': origen_nombre_actual,
            'timestamp': pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        salvo_algo = False

        if self.origen_principal_resultados.es_via_diccionario:
            decision = self._mostrar_dialogo_seleccion_salvado_via_diccionario()
            if decision['confirmed']:
                if decision['save_fcd'] and self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty:
                    regla_fcd = {**regla_a_guardar_base, 'tipo_datos_guardados': "COINCIDENCIAS_DICCIONARIO",
                                 'datos_snapshot': self.fcds_de_ultima_busqueda.to_dict(orient='records')}
                    self.reglas_guardadas.append(regla_fcd); salvo_algo = True
                    logger.info(f"Regla (FCD) salvada para: '{termino_original_regla}'")
                if decision['save_rfd'] and self.desc_finales_de_ultima_busqueda is not None and \
                   not self.desc_finales_de_ultima_busqueda.empty and \
                   self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC:
                    regla_rfd = {**regla_a_guardar_base, 'tipo_datos_guardados': "RESULTADOS_DESCRIPCION_VIA_DICCIONARIO",
                                 'datos_snapshot': self.desc_finales_de_ultima_busqueda.to_dict(orient='records')}
                    self.reglas_guardadas.append(regla_rfd); salvo_algo = True
                    logger.info(f"Regla (Resultados Desc vía Dic) salvada para: '{termino_original_regla}'")
        elif self.origen_principal_resultados.es_directo_descripcion or \
             self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA: # Búsqueda vacía
            if self.desc_finales_de_ultima_busqueda is not None: # Puede ser un DF vacío si se guardó "todo"
                tipo_datos = "TODAS_LAS_DESCRIPCIONES" if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA else "RESULTADOS_DESCRIPCION_DIRECTA"
                regla_directa = {**regla_a_guardar_base, 'tipo_datos_guardados': tipo_datos,
                                 'datos_snapshot': self.desc_finales_de_ultima_busqueda.to_dict(orient='records')}
                self.reglas_guardadas.append(regla_directa); salvo_algo = True
                logger.info(f"Regla ({tipo_datos}) salvada para: '{termino_original_regla}'")
            else: messagebox.showwarning("Sin Datos", "No hay resultados finales de descripción para salvar.")
        else: # Errores, NINGUNO, etc.
            if self.origen_principal_resultados != OrigenResultados.NINGUNO: # No mostrar si nunca se buscó
                messagebox.showerror("Error al Salvar", f"No se puede determinar qué salvar para el origen de resultados: {origen_nombre_actual}.")

        if salvo_algo: self._actualizar_estado(f"Regla(s) nueva(s) guardada(s). Total: {len(self.reglas_guardadas)}.")
        else: self._actualizar_estado("Ninguna regla fue salvada esta vez.")
        self._actualizar_botones_estado_general()

    def _mostrar_dialogo_seleccion_salvado_via_diccionario(self) -> Dict[str, bool]:
        # (Lógica de 'refactor' usando las nuevas variables de estado)
        decision = {'confirmed': False, 'save_fcd': False, 'save_rfd': False}
        dialog = tk.Toplevel(self); dialog.title("Seleccionar Datos a Salvar")
        dialog.geometry("400x200"); dialog.resizable(False, False)
        dialog.transient(self); dialog.grab_set()

        var_fcd = tk.BooleanVar(value=(self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty))
        var_rfd = tk.BooleanVar(value=(self.desc_finales_de_ultima_busqueda is not None and \
                                      not self.desc_finales_de_ultima_busqueda.empty and \
                                      self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC))

        ttk.Label(dialog, text="Búsqueda vía Diccionario. ¿Qué datos salvar?").pack(pady=10, padx=10)
        chk_fcd = ttk.Checkbutton(dialog, text="Coincidencias del Diccionario (FCDs)", variable=var_fcd)
        chk_fcd.pack(anchor=tk.W, padx=20)
        chk_fcd['state'] = 'normal' if self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty else 'disabled'
        
        chk_rfd = ttk.Checkbutton(dialog, text="Resultados Finales de Descripciones (RFDs)", variable=var_rfd)
        chk_rfd.pack(anchor=tk.W, padx=20)
        chk_rfd['state'] = 'normal' if self.desc_finales_de_ultima_busqueda is not None and \
                                      not self.desc_finales_de_ultima_busqueda.empty and \
                                      self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC else 'disabled'

        frame_botones_dialogo = ttk.Frame(dialog); frame_botones_dialogo.pack(pady=15)
        def on_confirm():
            decision['confirmed'] = True
            decision['save_fcd'] = var_fcd.get()
            decision['save_rfd'] = var_rfd.get()
            if not decision['save_fcd'] and not decision['save_rfd']:
                messagebox.showwarning("Nada Seleccionado", "No ha seleccionado datos para salvar.", parent=dialog)
                decision['confirmed'] = False; return 
            dialog.destroy()
        def on_cancel(): decision['confirmed'] = False; dialog.destroy()

        ttk.Button(frame_botones_dialogo, text="Confirmar", command=on_confirm).pack(side=tk.LEFT, padx=10)
        ttk.Button(frame_botones_dialogo, text="Cancelar", command=on_cancel).pack(side=tk.LEFT, padx=10)
        
        self.update_idletasks() # Para que winfo_width/height den valores correctos
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        self.wait_window(dialog)
        return decision

    def _exportar_resultados(self):
        # (Lógica de primordial)
        if not self.reglas_guardadas:
            messagebox.showwarning("Sin Reglas", "No hay reglas guardadas para exportar.")
            return

        timestamp_export = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        nombre_sugerido_base = f"exportacion_reglas_{timestamp_export}"
        ruta_guardar = filedialog.asksaveasfilename(
            title="Exportar Reglas Guardadas Como...",
            initialfile=f"{nombre_sugerido_base}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel (*.xlsx)", "*.xlsx")]
        )
        if not ruta_guardar:
            logger.info("Exportación cancelada."); self._actualizar_estado("Exportación cancelada."); return

        self._actualizar_estado("Exportando reglas..."); num_reglas = len(self.reglas_guardadas)
        logging.info(f"Exportando {num_reglas} regla(s) a: {ruta_guardar}")
        try:
            with pd.ExcelWriter(ruta_guardar, engine='openpyxl') as writer:
                datos_indice_export = []
                for i, regla_guardada in enumerate(self.reglas_guardadas):
                    id_regla_hoja = f"R{i+1}_{self._sanitizar_nombre_archivo(regla_guardada.get('termino_busqueda_original','S_T'),10)}"
                    datos_indice_export.append({
                        "ID_Regla_Hoja_Destino": id_regla_hoja,
                        "Termino_Busqueda_Original": regla_guardada.get('termino_busqueda_original', 'N/A'),
                        "Operador_Principal": regla_guardada.get('operador_principal', 'N/A'),
                        "Tipo_Datos_Guardados": regla_guardada.get('tipo_datos_guardados', 'N/A'),
                        "Fuente_Original_Resultados": regla_guardada.get('fuente_original_guardado', 'N/A'),
                        "Timestamp_Guardado": regla_guardada.get('timestamp', 'N/A'),
                        "Num_Filas_Snapshot": len(regla_guardada.get('datos_snapshot', []))
                    })
                    
                    df_regla_definicion = pd.DataFrame([{
                        'termino_original': regla_guardada.get('termino_busqueda_original'),
                        'operador_principal': regla_guardada.get('operador_principal'),
                        'terminos_analizados_json': json.dumps(regla_guardada.get('termino_busqueda_parseado'), ensure_ascii=False, indent=2)
                    }])
                    nombre_hoja_def = f"Def_{id_regla_hoja}"[:31] # Límite de nombre de hoja de Excel
                    df_regla_definicion.to_excel(writer, sheet_name=nombre_hoja_def, index=False)

                    datos_snapshot_list = regla_guardada.get('datos_snapshot')
                    if datos_snapshot_list: # Puede ser una lista vacía si se guardó "todo" y no había nada
                        df_snapshot = pd.DataFrame(datos_snapshot_list)
                        if not df_snapshot.empty: # Solo escribir si hay datos
                            nombre_hoja_snap = f"Snap_{id_regla_hoja}"[:31]
                            df_snapshot.to_excel(writer, sheet_name=nombre_hoja_snap, index=False)
                
                if datos_indice_export:
                    pd.DataFrame(datos_indice_export).to_excel(writer, sheet_name="Indice_Reglas_Exportadas", index=False)

            logging.info(f"Exportación de {num_reglas} regla(s) completada a {ruta_guardar}")
            messagebox.showinfo("Exportación Exitosa", f"{num_reglas} regla(s) exportadas a:\n{ruta_guardar}")
            self._actualizar_estado(f"Reglas exportadas a {Path(ruta_guardar).name}.")

            if messagebox.askyesno("Limpiar Reglas", "Exportación exitosa.\n¿Limpiar reglas guardadas internamente?"):
                self.reglas_guardadas.clear()
                self._actualizar_estado("Reglas guardadas limpiadas.")
                logging.info("Reglas guardadas limpiadas por usuario.")
            self._actualizar_botones_estado_general()

        except Exception as e:
            logging.exception("Error inesperado exportando reglas.")
            messagebox.showerror("Error Exportar", f"No se pudo exportar:\n{e}")
            self._actualizar_estado("Error exportando reglas.")

    def _sanitizar_nombre_archivo(self, texto: str, max_len: int = 50) -> str:
        # (Lógica de primordial)
        if not texto: return "resultados"
        sane = re.sub(r'[<>:"/\\|?*\s]+', '_', texto)
        sane = sane.strip('_')
        return sane[:max_len] if len(sane) > max_len else sane

    def _actualizar_estado_botones_operadores(self):
        # (Lógica de 'refactor' usando self.op_buttons)
        # Habilitar si al menos un archivo está cargado (para búsqueda directa)
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None:
            self._deshabilitar_botones_operadores(); return

        texto_completo = self.texto_busqueda_var.get(); cursor_pos = self.entrada_busqueda.index(tk.INSERT)
        segmento_antes_cursor = texto_completo[:cursor_pos]
        ultimo_op_logico_pos = -1
        for op_s in ["+", "|", "/"]:
            pos = segmento_antes_cursor.rfind(op_s)
            if pos > ultimo_op_logico_pos: ultimo_op_logico_pos = pos
        termino_actual_bruto = segmento_antes_cursor[ultimo_op_logico_pos + 1 :].lstrip()

        terminos_analizados_actual = self.motor._analizar_terminos([termino_actual_bruto])
        tipo_termino_actual = None; tiene_negacion_actual = False
        if terminos_analizados_actual:
            analisis_primero = terminos_analizados_actual[0]
            tipo_termino_actual = analisis_primero.get("tipo")
            if analisis_primero.get("negate"): tiene_negacion_actual = True
        
        # Botón NOT (#)
        # Se puede poner si no hay negación ya, y si el término actual (después de #) no está vacío.
        # O si no hay término actual (inicio de un nuevo sub-término).
        estado_not = "disabled" if tiene_negacion_actual or (termino_actual_bruto.strip() and not termino_actual_bruto.startswith("#")) else "normal"
        if termino_actual_bruto.strip() and not termino_actual_bruto.startswith("#"): # Si hay texto y no es negado, no puedo poner # delante
            estado_not = "disabled"
        elif tiene_negacion_actual and not termino_actual_bruto[1:].strip() : # Si es # y nada más
             estado_not = "normal" # Permitir escribir después del #
        elif not termino_actual_bruto.strip(): # Si el término actual está vacío (cursor al inicio o después de un operador)
            estado_not = "normal"


        self.op_buttons["#"]["state"] = estado_not

        es_comp_o_rango = tipo_termino_actual in ["gt", "lt", "ge", "le", "range"]
        # Deshabilitar si ya es comparativo/rango, o si es negado y no hay nada después del #
        estado_comparacion = "disabled" if es_comp_o_rango or (tiene_negacion_actual and not termino_actual_bruto[1:].strip()) else "normal"
        for op_key in [">", "<", ">=", "<="]: self.op_buttons[op_key]["state"] = estado_comparacion
        
        puede_poner_rango = False
        # Solo si hay un término actual, no es negado (o si es negado, hay algo después del #), y no es ya comparativo/rango
        # Y si lo que hay es plausiblemente un número o inicio de unidad.
        termino_efectivo_para_rango = termino_actual_bruto[1:].strip() if tiene_negacion_actual else termino_actual_bruto.strip()
        if termino_efectivo_para_rango and not es_comp_o_rango:
            if termino_efectivo_para_rango[-1].isdigit() or (len(termino_efectivo_para_rango) > 1 and termino_efectivo_para_rango[-1].isalpha() and termino_efectivo_para_rango[-2].isdigit()):
                 puede_poner_rango = True
        self.op_buttons["-"]["state"] = "normal" if puede_poner_rango else "disabled"

        estado_logico = "normal"
        texto_limpio_final_total = texto_completo.rstrip()
        # Deshabilitar si no hay nada, o si lo último es un operador lógico, o si estamos justo después de un operador y no hay término
        if not texto_completo.strip() or \
           (texto_limpio_final_total and texto_limpio_final_total[-1] in ["+", "|", "/"]) or \
           (not termino_actual_bruto.strip() and ultimo_op_logico_pos != -1 and ultimo_op_logico_pos == len(segmento_antes_cursor)-1): # Justo después de 'op '
            estado_logico = "disabled"
        for op_key in ["+", "|"]: self.op_buttons[op_key]["state"] = estado_logico
        # Asumo que el botón "/" no existe en self.op_buttons como tal o se maneja con "|"

    def _insertar_operador_validado(self, operador: str):
        # (Lógica de 'refactor' usando el estado del botón)
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None:
            return

        if operador in self.op_buttons and self.op_buttons[operador]["state"] == "disabled":
            logger.debug(f"Intento de insertar operador '{operador}' deshabilitado ignorado.")
            return 

        texto_a_insertar = operador
        if operador in ["+", "|", "/"]: # El botón "/" debe estar en self.op_buttons si se usa así
            texto_a_insertar = f" {operador} " 
        
        self.entrada_busqueda.insert(tk.INSERT, texto_a_insertar)
        self.entrada_busqueda.focus_set() # Devolver foco a la entrada

    def _deshabilitar_botones_operadores(self):
        for btn in self.op_buttons.values():
            btn["state"] = "disabled"

    def on_closing(self):
        logger.info("Cerrando la aplicación...")
        self._guardar_configuracion()
        self.destroy()


if __name__ == "__main__":
    log_file_name = "buscador_app_FUSION_V2.log" # Nuevo nombre de log
    logging.basicConfig(
        level=logging.DEBUG, 
        format='%(asctime)s - %(name)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file_name, encoding='utf-8', mode='a'), 
            logging.StreamHandler()
        ]
    )
    logger.info("=============================================")
    logger.info("=== Iniciando Aplicación Buscador (Fusionada V2) ===")
    
    missing_deps = []
    try: import pandas as pd; logger.info(f"Pandas versión: {pd.__version__}")
    except ImportError: missing_deps.append("pandas"); logger.critical("Dependencia faltante: pandas")
    try: import openpyxl; logger.info(f"openpyxl versión: {openpyxl.__version__}")
    except ImportError: missing_deps.append("openpyxl (para .xlsx)") 
    
    if "pandas" in missing_deps: 
        error_msg_dep = f"Faltan librerías críticas: {', '.join(missing_deps)}.\nInstale con: pip install {' '.join(missing_deps)}"
        logger.critical(error_msg_dep)
        try: 
            root_temp = tk.Tk(); root_temp.withdraw(); messagebox.showerror("Dependencias Faltantes", error_msg_dep); root_temp.destroy()
        except tk.TclError: print(f"ERROR CRÍTICO (Tkinter no disponible): {error_msg_dep}")
        except Exception as e_tk_init: print(f"ERROR CRÍTICO (Error al mostrar msgbox): {e_tk_init}\n{error_msg_dep}")
        exit(1)

    try:
        app = InterfazGrafica()
        app.mainloop()
    except Exception as main_error:
        logger.critical("¡Error fatal no capturado en la aplicación!", exc_info=True)
        try:
            root_err = tk.Tk(); root_err.withdraw()
            messagebox.showerror("Error Fatal", f"Error crítico:\n{main_error}\nConsulte '{log_file_name}'.")
            root_err.destroy()
        except Exception as fallback_error:
            logger.error(f"No se pudo mostrar el mensaje de error fatal vía Tkinter: {fallback_error}")
            print(f"ERROR FATAL: {main_error}. Consulte {log_file_name}.")
    finally:
        logger.info("=== Finalizando Aplicación Buscador (Fusionada V2) ===")
