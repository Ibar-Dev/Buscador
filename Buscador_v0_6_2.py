# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Tuple, Union, Set, Callable, Dict, Any, Literal # Literal añadido
from enum import Enum, auto
import traceback
import platform
import unicodedata
import logging
import json
import os
from pathlib import Path # Usando pathlib
import string
# import numpy as np # Mantengo comentado si no se usa directamente en esta UI

# --- Configuración del Logging ---
# Configurado globalmente al inicio del script en if __name__ == "__main__"
logger = logging.getLogger(__name__) # Crear logger para este módulo

# --- Enumeraciones (Usando la versión mejorada de PyGuru) ---
class OrigenResultados(Enum):
    NINGUNO = 0
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto()
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto()
    DIRECTO_DESCRIPCION_CON_RESULTADOS = auto()
    DIRECTO_DESCRIPCION_VACIA = auto()
    ERROR_CARGA_DICCIONARIO = auto()
    ERROR_CARGA_DESCRIPCION = auto()
    ERROR_CONFIGURACION_COLUMNAS = auto()
    ERROR_BUSQUEDA_INTERNA = auto()

    @property
    def es_via_diccionario(self) -> bool:
        return self in {
            OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
            OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
        }

    @property
    def es_directo_descripcion(self) -> bool:
        return self in {
            OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS,
            OrigenResultados.DIRECTO_DESCRIPCION_VACIA
        }

    @property
    def es_error(self) -> bool:
        return self in {
            OrigenResultados.ERROR_CARGA_DICCIONARIO,
            OrigenResultados.ERROR_CARGA_DESCRIPCION,
            OrigenResultados.ERROR_CONFIGURACION_COLUMNAS,
            OrigenResultados.ERROR_BUSQUEDA_INTERNA
        }

# --- Clases de Utilidad (Versión PyGuru) ---
class ExtractorMagnitud:
    MAGNITUDES_PREDEFINIDAS: List[str] = [
        "A","AMP","AMPS","AH","ANTENNA","BASE","BIT","ETH","FE","G","GB",
        "GBE","GE","GIGABIT","GBASE","GBASEWAN","GBIC","GBIT","GBPS","GH",
        "GHZ","HZ","KHZ","KM","KVA","KW","LINEAS","LINES","MHZ","NM","PORT",
        "PORTS","PTOS","PUERTO","PUERTOS","P","V","VA","VAC","VC","VCC",
        "VCD","VDC","W","WATTS","E","FE","GBE","GE","POTS","STM", "VOLTIOS", "VATIOS", "AMPERIOS"
    ]

    def __init__(self, magnitudes: Optional[List[str]] = None):
        self.magnitudes_normalizadas: Dict[str, str] = {
            self._normalizar_texto(m.upper()): m
            for m in (magnitudes if magnitudes is not None else self.MAGNITUDES_PREDEFINIDAS)
        }
        # logger.info(f"ExtractorMagnitud inicializado con {len(self.magnitudes_normalizadas)} magnitudes.") # Ya no es global

    @staticmethod
    def _normalizar_texto(texto: str) -> str:
        if not isinstance(texto, str) or not texto: return ""
        try:
            forma_normalizada = unicodedata.normalize('NFKD', texto.upper())
            return ''.join(c for c in forma_normalizada if not unicodedata.combining(c))
        except TypeError:
            # logger.warning(f"TypeError al normalizar texto: {texto}") # Ya no es global
            return ""

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        normalizada = self._normalizar_texto(texto_unidad)
        return self.magnitudes_normalizadas.get(normalizada)

    # La función buscar_cantidad_para_magnitud de tu script original no parece usarse
    # activamente con la nueva lógica de _generar_mascara_para_subtermino, así que la omito aquí
    # para brevedad, pero puedes reinsertarla si la necesitas.

class ManejadorExcel: # (Versión PyGuru)
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Optional[pd.DataFrame]:
        ruta = Path(ruta_archivo)
        logger.info(f"Intentando cargar archivo Excel: {ruta}")
        if not ruta.exists():
            logger.error(f"¡Archivo no encontrado! Ruta: {ruta}")
            messagebox.showerror("Error de Archivo", f"No se encontró el archivo:\n{ruta}\n\nVerifique que la ruta sea correcta.")
            return None
        try:
            engine = 'openpyxl' if ruta.suffix == '.xlsx' else None
            df = pd.read_excel(ruta, engine=engine)
            logger.info(f"Archivo '{ruta.name}' cargado ({len(df)} filas).")
            return df
        except Exception as e:
            logger.exception(f"Error inesperado al cargar archivo Excel: {ruta}")
            messagebox.showerror("Error al Cargar",
                                 f"No se pudo cargar el archivo:\n{ruta}\n\nError: {e}\n\n"
                                 "Posibles causas:\n"
                                 "- El archivo está siendo usado por otro programa.\n"
                                 "- No tiene instalado 'openpyxl' para archivos .xlsx.\n"
                                 "- El archivo está corrupto o en formato no soportado.")
            return None

# --- Clase MotorBusqueda (Versión PyGuru Mejorada) ---
class MotorBusqueda:
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[Path] = None # Usar Path
        self.archivo_descripcion_actual: Optional[Path] = None # Usar Path

        self.indices_columnas_busqueda_dic: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else []
        logger.info(f"MotorBusqueda inicializado. Índices búsqueda diccionario: {self.indices_columnas_busqueda_dic or 'Todas las de texto'}")

        self.patron_comparacion_compilado = re.compile(
            r"^([<>]=?)(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ]+)?(.*)$"
        )
        self.patron_rango_compilado = re.compile(
            r"^(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ]+)?$"
        )
        self.patron_negacion_compilado = re.compile(r"^#(.+)$")
        self.patron_num_unidad_df = re.compile(r"(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ]+)?")

        self.extractor_magnitud = ExtractorMagnitud()

    def cargar_excel_diccionario(self, ruta_str: str) -> bool:
        ruta = Path(ruta_str) # Convertir a Path
        df_cargado = ManejadorExcel.cargar_excel(ruta)
        if df_cargado is None:
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            return False
        self.datos_diccionario = df_cargado
        self.archivo_diccionario_actual = ruta # Guardar como Path
        if not self._validar_columnas_df(self.datos_diccionario, self.indices_columnas_busqueda_dic, "diccionario"):
            logger.warning("Validación de columnas del diccionario fallida. Carga invalidada.")
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            return False
        return True

    def cargar_excel_descripcion(self, ruta_str: str) -> bool:
        ruta = Path(ruta_str) # Convertir a Path
        df_cargado = ManejadorExcel.cargar_excel(ruta)
        if df_cargado is None:
            self.datos_descripcion = None
            self.archivo_descripcion_actual = None
            return False
        self.datos_descripcion = df_cargado
        self.archivo_descripcion_actual = ruta # Guardar como Path
        return True

    def _validar_columnas_df(self, df: Optional[pd.DataFrame], indices_cfg: List[int], nombre_df: str) -> bool:
        # (Implementación de PyGuru de _validar_columnas_df)
        if df is None:
            logger.error(f"DataFrame '{nombre_df}' es None, no se puede validar.")
            return False
        num_cols_df = len(df.columns)

        if not indices_cfg or indices_cfg == [-1]:
            if num_cols_df == 0:
                logger.error(f"El archivo del {nombre_df} está vacío o no contiene columnas (modo 'todas').")
                messagebox.showerror(f"Error de {nombre_df.capitalize()}", f"El archivo del {nombre_df} está vacío o no contiene columnas.")
                return False
            return True

        if not all(isinstance(idx, int) and idx >= 0 for idx in indices_cfg):
            logger.error(f"Configuración de índices para {nombre_df} inválida: {indices_cfg}. Deben ser enteros no negativos.")
            messagebox.showerror("Error de Configuración", f"Los índices de columna para {nombre_df} deben ser números enteros no negativos. Config: {indices_cfg}")
            return False

        max_indice_requerido = max(indices_cfg) if indices_cfg else -1

        if num_cols_df == 0: # Esto ya lo cubre el primer if not indices_cfg, pero por si acaso.
            logger.error(f"El {nombre_df} no tiene columnas.")
            messagebox.showerror(f"Error de {nombre_df.capitalize()}", f"El archivo del {nombre_df} está vacío o no contiene columnas.")
            return False
        elif max_indice_requerido >= num_cols_df:
            msg = (f"El {nombre_df} necesita al menos {max_indice_requerido + 1} columnas "
                   f"para los índices configurados ({indices_cfg}), pero solo tiene {num_cols_df}.")
            logger.error(msg)
            messagebox.showerror(f"Error de Columnas en {nombre_df.capitalize()}", msg)
            return False
        return True

    def _obtener_nombres_columnas_busqueda_df(self, df: Optional[pd.DataFrame], indices_cfg: List[int], nombre_df_log: str) -> Optional[List[str]]:
        # (Implementación de PyGuru de _obtener_nombres_columnas_busqueda - renombrado para evitar conflicto)
        if df is None:
            logger.error(f"Intentando obtener nombres de columnas de un DataFrame ({nombre_df_log}) que es None.")
            return None

        columnas_disponibles = df.columns
        num_cols_df = len(columnas_disponibles)

        if not indices_cfg or indices_cfg == [-1]:
            cols_texto_obj = [col for col in df.columns if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
            if cols_texto_obj:
                logger.info(f"Buscando en columnas de texto/object (detectadas) del {nombre_df_log}: {cols_texto_obj}")
                return cols_texto_obj
            elif num_cols_df > 0:
                logger.warning(f"No se encontraron columnas de texto/object en {nombre_df_log}. Se usarán todas las {num_cols_df} columnas.")
                return list(df.columns)
            else:
                logger.error(f"El DataFrame del {nombre_df_log} no tiene columnas.")
                messagebox.showerror(f"Error en {nombre_df_log.capitalize()}", f"El archivo del {nombre_df_log} no tiene columnas.")
                return None

        nombres_columnas_seleccionadas = []
        indices_validos_usados = []
        for indice in indices_cfg:
            if 0 <= indice < num_cols_df:
                nombres_columnas_seleccionadas.append(columnas_disponibles[indice])
                indices_validos_usados.append(indice)
            else:
                logger.warning(f"Índice {indice} para {nombre_df_log} es inválido o fuera de rango (0 a {num_cols_df-1}). Ignorado.")

        if not nombres_columnas_seleccionadas:
            msg = f"No se encontraron columnas válidas en {nombre_df_log} con los índices configurados: {indices_cfg}"
            logger.error(msg)
            messagebox.showerror("Error de Configuración", msg)
            return None

        logger.debug(f"Se buscará en columnas del {nombre_df_log}: {nombres_columnas_seleccionadas} (índices: {indices_validos_usados})")
        return nombres_columnas_seleccionadas

    @staticmethod
    def _parsear_termino_busqueda_general(termino_raw: str) -> Tuple[str, List[str]]:
        # (Implementación de PyGuru)
        termino_proc = termino_raw.strip().upper() # Convertir a mayúsculas una vez aquí
        operador = 'AND'

        if termino_proc.startswith('OR:'):
            operador = 'OR'
            termino_proc = termino_proc[3:].strip()
        elif termino_proc.startswith('AND:'):
            termino_proc = termino_proc[4:].strip()
        
        # Para tu UI, los operadores lógicos se manejan por los botones +, |, /
        # y la estructura de la cadena de búsqueda.
        # Aquí adaptamos para que coincida con tu UI: +, | o / separan términos.
        # El operador global (AND/OR) se decide por el primer separador encontrado.
        
        separadores_logicos = {'+': 'AND', '|': 'OR', '/': 'OR'}
        terminos_individuales = []
        
        # Regex para split manteniendo delimitadores (no necesario si tu UI los añade explícitamente)
        # Para la lógica de tu UI donde los operadores son parte del string,
        # el parseo que haces en _parsear_termino_busqueda_inicial es más adecuado.
        # Esta función es para un parseo más general si la estructura es "OPERADOR: term1 term2"
        
        # Simplificando: asumimos que tu UI _parsear_termino_busqueda_inicial ya ha hecho el split principal
        # y esta función solo refina cada "subtermino_raw"
        # Por ahora, esta función no se usa directamente por la UI, sino _analizar_terminos.
        
        # La que usas en tu UI es _parsear_termino_busqueda_inicial. Dejamos esta como referencia interna.
        terminos_individuales = [t.strip() for t in termino_proc.split() if t.strip()]
        logger.debug(f"Parseo general (interno): Operador='{operador}', Términos='{terminos_individuales}'")
        return operador, terminos_individuales


    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        # (Adaptado de la versión PyGuru _analizar_subtermino_individual)
        palabras_analizadas = []
        for term_orig_bruto in terminos_brutos:
            term_orig = str(term_orig_bruto) # Asegurar que es string
            term = term_orig.strip()
            if not term: continue

            item_analizado: Dict[str, Any] = {'original': term_orig, 'negate': False}
            match_neg = self.patron_negacion_compilado.match(term)
            if match_neg:
                item_analizado['negate'] = True # Bool
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
                else:
                    item_analizado.update({'tipo': 'str', 'valor': term.upper()})
            elif match_range:
                v1_str, v2_str, unidad_rango_str = match_range.groups()
                v1, v2 = self._parse_numero(v1_str), self._parse_numero(v2_str)
                if v1 is not None and v2 is not None:
                    item_analizado.update({
                        'tipo': 'range',
                        'valor': sorted([v1, v2]),
                        'unidad_busqueda': self.extractor_magnitud._normalizar_texto(unidad_rango_str.strip()) if unidad_rango_str else None
                    })
                else:
                    item_analizado.update({'tipo': 'str', 'valor': term.upper()})
            else:
                item_analizado.update({'tipo': 'str', 'valor': term.upper()})
            palabras_analizadas.append(item_analizado)
        logger.debug(f"Términos analizados (motor): {palabras_analizadas}")
        return palabras_analizadas


    def _parse_numero(self, num_str: Any) -> Optional[float]: # Tu UI lo pasa como string, pero Any es más seguro
        # (Implementación de PyGuru)
        if not isinstance(num_str, (str, int, float)): return None
        try:
            return float(str(num_str).replace(',', '.'))
        except ValueError:
            return None

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols_a_buscar: List[str], termino_analizado: Dict[str, Any]) -> pd.Series:
        # (Lógica de PyGuru para _generar_mascara_para_subtermino, renombrada)
        mascara_total_subtermino = pd.Series(False, index=df.index)
        tipo_sub = termino_analizado['tipo']
        valor_sub = termino_analizado['valor'] # Puede ser float, list de floats, o str
        unidad_sub_requerida = termino_analizado.get('unidad_busqueda')
        es_negado = termino_analizado.get('negate', False)


        for col_nombre in cols_a_buscar:
            if col_nombre not in df.columns:
                logger.warning(f"Columna '{col_nombre}' no encontrada en el DataFrame al generar máscara. Saltando.")
                continue
            col_series = df[col_nombre]
            mascara_col_actual = pd.Series(False, index=df.index)

            if tipo_sub in ['gt', 'lt', 'ge', 'le', 'range']:
                for idx, valor_celda in col_series.items():
                    if pd.isna(valor_celda) or str(valor_celda).strip() == "": continue
                    for match_celda in self.patron_num_unidad_df.finditer(str(valor_celda)):
                        try:
                            num_celda_val = float(match_celda.group(1).replace(',', '.'))
                            unidad_celda_val = self.extractor_magnitud._normalizar_texto(match_celda.group(2)) if match_celda.group(2) else None
                            if unidad_sub_requerida and unidad_celda_val != unidad_sub_requerida: continue
                            
                            cond_ok = False
                            if tipo_sub == 'gt' and num_celda_val > valor_sub: cond_ok = True
                            elif tipo_sub == 'lt' and num_celda_val < valor_sub: cond_ok = True
                            elif tipo_sub == 'ge' and num_celda_val >= valor_sub: cond_ok = True
                            elif tipo_sub == 'le' and num_celda_val <= valor_sub: cond_ok = True
                            elif tipo_sub == 'range' and valor_sub[0] <= num_celda_val <= valor_sub[1]: cond_ok = True
                            
                            if cond_ok:
                                mascara_col_actual.at[idx] = True
                                break
                        except ValueError: continue
                    if mascara_col_actual.at[idx]: continue
            elif tipo_sub == 'str':
                # valor_sub ya está en MAYÚSCULAS desde _analizar_terminos
                termino_regex_escapado = r"\b" + re.escape(str(valor_sub)) + r"\b"
                try:
                    if pd.api.types.is_string_dtype(col_series) or pd.api.types.is_object_dtype(col_series):
                         mascara_col_actual = col_series.astype(str).str.upper().str.contains(termino_regex_escapado, regex=True, na=False)
                    else:
                        mascara_col_actual = col_series.astype(str).str.upper().str.contains(termino_regex_escapado, regex=True, na=False)
                except Exception as e_conv:
                    logger.warning(f"No se pudo convertir/buscar string en columna '{col_nombre}': {e_conv}")

            mascara_total_subtermino |= mascara_col_actual.fillna(False)
        
        return ~mascara_total_subtermino if es_negado else mascara_total_subtermino

    def _aplicar_mascara_combinada(self, df: pd.DataFrame, cols_a_buscar: List[str], terminos_analizados: List[Dict[str, Any]], op_principal: str) -> pd.Series:
        # (Adaptado de tu _aplicar_mascara_diccionario y la lógica de PyGuru)
        if df is None or df.empty or not cols_a_buscar or not terminos_analizados:
            return pd.Series(False, index=df.index if df is not None else None)

        cols_ok_en_df = [c for c in cols_a_buscar if c in df.columns]
        if not cols_ok_en_df:
            logger.error(f"Ninguna columna de búsqueda ({cols_a_buscar}) existe en el DataFrame provisto.")
            return pd.Series(False, index=df.index)

        op_es_and = op_principal.upper() == 'AND'
        mascara_final_combinada = pd.Series(True, index=df.index) if op_es_and else pd.Series(False, index=df.index)

        if not terminos_analizados: # Si no hay términos válidos, el comportamiento depende del operador
             return mascara_final_combinada # AND: todo, OR: nada

        for termino_individual_analizado in terminos_analizados:
            mascara_este_termino = self._generar_mascara_para_un_termino(df, cols_ok_en_df, termino_individual_analizado)
            if op_es_and:
                mascara_final_combinada &= mascara_este_termino
            else: # OR
                mascara_final_combinada |= mascara_este_termino
        
        return mascara_final_combinada

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, cols_nombres: List[str]) -> Set[str]:
        # (Lógica de PyGuru _extraer_terminos_clave_de_diccionario)
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
        # (Lógica de PyGuru _buscar_terminos_clave_en_descripciones, adaptada)
        if df_desc_a_buscar is None or df_desc_a_buscar.empty:
            logger.warning("DataFrame de descripciones (para búsqueda de términos clave) no provisto o vacío.")
            return pd.DataFrame(columns=(df_desc_a_buscar.columns if df_desc_a_buscar is not None else []))
        
        if not terminos_clave:
            logger.info("No hay términos clave para buscar en las descripciones.")
            return pd.DataFrame(columns=df_desc_a_buscar.columns) # Devolver estructura vacía

        logger.info(f"Buscando {len(terminos_clave)} términos clave en {len(df_desc_a_buscar)} descripciones (require_all={require_all_terminos_clave}).")
        
        texto_filas_desc = df_desc_a_buscar.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
        terminos_escapados = [r"\b" + re.escape(t) + r"\b" for t in terminos_clave if t]
        if not terminos_escapados:
            return pd.DataFrame(columns=df_desc_a_buscar.columns)

        if require_all_terminos_clave:
            try:
                mascara_final_desc = texto_filas_desc.apply(lambda txt_fila: all(re.search(pat, txt_fila, re.IGNORECASE) for pat in terminos_escapados))
            except Exception as e:
                logger.exception(f"Error durante la búsqueda 'require_all' en descripciones: {e}")
                return pd.DataFrame(columns=df_desc_a_buscar.columns)
        else:
            patron_or_regex = '|'.join(terminos_escapados)
            try:
                mascara_final_desc = texto_filas_desc.str.contains(patron_or_regex, regex=True, case=False, na=False)
            except Exception as e:
                logger.exception(f"Error durante la búsqueda 'OR' en descripciones: {e}")
                return pd.DataFrame(columns=df_desc_a_buscar.columns)

        resultados_desc = df_desc_a_buscar[mascara_final_desc]
        logger.info(f"Búsqueda de términos clave en descripciones completada. Resultados: {len(resultados_desc)} filas.")
        return resultados_desc

    # --- MÉTODO DE BÚSQUEDA PRINCIPAL DEL MOTOR (llamado por la UI) ---
    def buscar(self, df_dic_para_busqueda: pd.DataFrame, df_desc_para_busqueda: pd.DataFrame,
               terminos_analizados_usuario: List[Dict[str, Any]], operador_principal_usuario: str,
               buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame]]:
        """
        Lógica de búsqueda principal del motor.
        Devuelve: (resultados_finales, origen_del_resultado, fcds_si_aplica)
        """
        logger.info(f"Motor.buscar: op='{operador_principal_usuario}', via_dicc={buscar_via_diccionario_flag}, terminos={terminos_analizados_usuario}")
        
        fcds_obtenidos: Optional[pd.DataFrame] = None

        if buscar_via_diccionario_flag:
            if df_dic_para_busqueda is None: return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None
            
            cols_dic = self._obtener_nombres_columnas_busqueda_df(df_dic_para_busqueda, self.indices_columnas_busqueda_dic, "diccionario (búsqueda)")
            if not cols_dic: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS, None

            mascara_en_dic = self._aplicar_mascara_combinada(df_dic_para_busqueda, cols_dic, terminos_analizados_usuario, operador_principal_usuario)
            fcds_obtenidos = df_dic_para_busqueda[mascara_en_dic].copy()
            logger.info(f"FCDs encontrados: {len(fcds_obtenidos)}")

            if fcds_obtenidos.empty:
                return pd.DataFrame(columns=df_desc_para_busqueda.columns if df_desc_para_busqueda is not None else []), OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, fcds_obtenidos

            terminos_extraidos_de_fcds = self._extraer_terminos_diccionario(fcds_obtenidos, cols_dic)
            if not terminos_extraidos_de_fcds:
                # Se encontraron FCDs pero no se pudieron extraer términos para pasar a descripciones
                # Devolver FCDs como resultado principal en este caso podría ser una opción
                # o considerarlo como un tipo de "sin resultados en descripción"
                return pd.DataFrame(columns=df_desc_para_busqueda.columns if df_desc_para_busqueda is not None else []), OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, fcds_obtenidos

            if df_desc_para_busqueda is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos
            
            # La lógica de require_all aquí es crucial. Si el op_principal era AND,
            # ¿deberían TODOS los términos extraídos estar en la descripción, o CUALQUIERA?
            # Por ahora, usamos OR (False) para los términos extraídos.
            resultados_desc_via_dic = self._buscar_terminos_en_descripciones(df_desc_para_busqueda, terminos_extraidos_de_fcds, require_all_terminos_clave=False)

            if resultados_desc_via_dic.empty:
                return pd.DataFrame(columns=df_desc_para_busqueda.columns), OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, fcds_obtenidos
            else:
                return resultados_desc_via_dic, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos
        
        # Búsqueda directa en descripciones
        else:
            if df_desc_para_busqueda is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None
            
            # Para búsqueda directa, usar todas las columnas del df de descripciones
            cols_desc = self._obtener_nombres_columnas_busqueda_df(df_desc_para_busqueda, [], "descripciones (búsqueda directa)")
            if not cols_desc: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS, None # No debería pasar si el df tiene columnas

            mascara_en_desc_directa = self._aplicar_mascara_combinada(df_desc_para_busqueda, cols_desc, terminos_analizados_usuario, operador_principal_usuario)
            resultados_directos = df_desc_para_busqueda[mascara_en_desc_directa].copy()
            logger.info(f"Resultados directos en descripción: {len(resultados_directos)}")

            if resultados_directos.empty:
                return pd.DataFrame(columns=df_desc_para_busqueda.columns), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None
            else:
                return resultados_directos, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None


# --- Clase InterfazGrafica (Tu versión, adaptada para usar el nuevo MotorBusqueda) ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE = "config_buscador.json"

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado PyGuru Mod")
        self.geometry("1250x800")

        self.config = self._cargar_configuracion()
        indices_cfg = self.config.get("indices_columnas_busqueda_dic", []) # Default a [] (todas texto)
        
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg)

        self.resultados_actuales: Optional[pd.DataFrame] = None
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        self.ultimo_termino_buscado: Optional[str] = None
        self.reglas_guardadas: List[Dict[str, Any]] = []
        self.df_candidato_diccionario: Optional[pd.DataFrame] = None # FCDs
        self.df_candidato_descripcion: Optional[pd.DataFrame] = None # Resultados finales de descripción
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO
        self.color_fila_par = "white"; self.color_fila_impar = "#f0f0f0"

        self._configurar_estilo_ttk()
        self._crear_widgets()
        self._configurar_grid()
        self._configurar_eventos()
        self._configurar_tags_treeview()
        self._configurar_orden_tabla(self.tabla_resultados)
        self._configurar_orden_tabla(self.tabla_diccionario)

        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones.")
        self._deshabilitar_botones_operadores()
        self._actualizar_botones_estado_general()
        logger.info("Interfaz Gráfica inicializada.")

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str):
        self._actualizar_estado_botones_operadores()

    def _cargar_configuracion(self) -> Dict:
        config = {}
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f: config = json.load(f)
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE}")
            except Exception as e:
                logger.error(f"Error al cargar config: {e}")
                messagebox.showwarning("Error Configuración", f"No se pudo cargar config:\n{e}")
        else:
            logger.info("Archivo de configuración no encontrado. Se creará al cerrar.")
        
        # Usar Path para las rutas si existen
        last_dic_path_str = config.get("last_dic_path")
        config["last_dic_path"] = str(Path(last_dic_path_str)) if last_dic_path_str else None
        
        last_desc_path_str = config.get("last_desc_path")
        config["last_desc_path"] = str(Path(last_desc_path_str)) if last_desc_path_str else None
        
        config.setdefault("indices_columnas_busqueda_dic", [])
        return config

    def _guardar_configuracion(self):
        # Guardar rutas como strings en JSON
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
    
    # --- Métodos de la UI (_configurar_estilo_ttk, _crear_widgets, _configurar_grid, etc.) ---
    # ... (Estos métodos se mantienen en gran medida como en tu script original de UI) ...
    # ... (Asegúrate de que los command= de los botones llamen a los métodos correctos) ...

    def _configurar_estilo_ttk(self):
        style = ttk.Style(self); themes = style.theme_names(); os_name = platform.system()
        prefs = {"Windows":["vista","xpnative","clam"],"Darwin":["aqua","clam"],"Linux":["clam","alt","default"]}
        theme_to_use = next((t for t in prefs.get(os_name, ["clam","default"]) if t in themes), None)
        if not theme_to_use:
            theme_to_use = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))
        if theme_to_use:
            logger.info(f"Aplicando tema TTK: {theme_to_use}")
            try: style.theme_use(theme_to_use)
            except tk.TclError as e: logger.warning(f"No se pudo aplicar tema '{theme_to_use}': {e}.")
        else: logger.warning("No se encontró tema TTK disponible.")

    def _crear_widgets(self):
        self.marco_controles = ttk.LabelFrame(self, text="Controles")
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.lbl_dic_cargado = ttk.Label(self.marco_controles, text="Dic: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion)
        self.lbl_desc_cargado = ttk.Label(self.marco_controles, text="Desc: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)

        self.frame_ops = ttk.Frame(self.marco_controles)
        self.btn_and = ttk.Button(self.frame_ops, text="+", width=3, command=lambda: self._insertar_operador_validado("+"))
        self.btn_or = ttk.Button(self.frame_ops, text="|", width=3, command=lambda: self._insertar_operador_validado("|"))
        self.btn_not = ttk.Button(self.frame_ops, text="#", width=3, command=lambda: self._insertar_operador_validado("#"))
        self.btn_gt = ttk.Button(self.frame_ops, text=">", width=3, command=lambda: self._insertar_operador_validado(">"))
        self.btn_lt = ttk.Button(self.frame_ops, text="<", width=3, command=lambda: self._insertar_operador_validado("<"))
        self.btn_ge = ttk.Button(self.frame_ops, text="≥", width=3, command=lambda: self._insertar_operador_validado(">="))
        self.btn_le = ttk.Button(self.frame_ops, text="≤", width=3, command=lambda: self._insertar_operador_validado("<="))
        self.btn_range = ttk.Button(self.frame_ops, text="-", width=3, command=lambda: self._insertar_operador_validado("-"))

        self.btn_and.grid(row=0, column=0, padx=1); self.btn_or.grid(row=0, column=1, padx=1)
        self.btn_not.grid(row=0, column=2, padx=1); self.btn_gt.grid(row=0, column=3, padx=1)
        self.btn_lt.grid(row=0, column=4, padx=1); self.btn_ge.grid(row=0, column=5, padx=1)
        self.btn_le.grid(row=0, column=6, padx=1); self.btn_range.grid(row=0, column=7, padx=1)

        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50, textvariable=self.texto_busqueda_var)
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda) # No necesita lambda aquí
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
        self.grid_rowconfigure(2, weight=1); self.grid_rowconfigure(4, weight=3)
        self.grid_columnconfigure(0, weight=1)
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        self.marco_controles.grid_columnconfigure(1, weight=1)
        self.marco_controles.grid_columnconfigure(3, weight=1) # Añadido para expandir lbl_desc_cargado

        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew")
        
        self.frame_ops.grid(row=1, column=0, columnspan=4, padx=5, pady=(5,0), sticky="w") # Ajustado columnspan
        self.entrada_busqueda.grid(row=2, column=0, columnspan=2, padx=5, pady=(0,5), sticky="ew")
        self.btn_buscar.grid(row=2, column=2, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_salvar_regla.grid(row=2, column=3, padx=(2,0), pady=(0,5), sticky="w") # Columna 3
        self.btn_ayuda.grid(row=2, column=4, padx=(2,0), pady=(0,5), sticky="w")     # Columna 4
        self.btn_exportar.grid(row=2, column=5, padx=(10, 5), pady=(0,5), sticky="e") # Columna 5

        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0))
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))
        
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1); self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew"); self.scrolly_diccionario.grid(row=0, column=1, sticky="ns"); self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")
        
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1); self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew"); self.scrolly_resultados.grid(row=0, column=1, sticky="ns"); self.scrollx_resultados.grid(row=1, column=0, sticky="ew")
        
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0))

    def _configurar_eventos(self):
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        # self.entrada_busqueda.bind("<KeyRelease>", self._on_key_release_busqueda_dic) # Mantenerlo si se usa

    def _actualizar_estado(self, mensaje: str):
        self.barra_estado.config(text=mensaje)
        logger.info(f"Estado UI: {mensaje}")
        self.update_idletasks()

    def _mostrar_ayuda(self):
        # (Tu ayuda es completa, sin cambios)
        ayuda = """Sintaxis de Búsqueda en Diccionario:
-------------------------------------
- Texto simple: Busca la palabra o frase exacta (insensible a mayús/minús).
  Ej: `router cisco`

- Operadores Lógicos:
  * `término1 + término2`: Busca filas que contengan AMBOS términos (AND).
    Ej: `tarjeta + 16 puertos`
  * `término1 | término2`: Busca filas que contengan AL MENOS UNO (OR).
    Ej: `modulo | SFP`
  * `término1 / término2`: Alternativa para OR.

- Comparaciones numéricas (con o sin unidad, la unidad debe coincidir si se especifica):
  * `>numero[UNIDAD]`: Mayor que. Ej: `>1000` o `>1000w`
  * `<numero[UNIDAD]`: Menor que. Ej: `<50` o `<50v`
  * `>=numero[UNIDAD]` o `≥numero[UNIDAD]`: Mayor o igual. Ej: `>=48a`
  * `<=numero[UNIDAD]` o `≤numero[UNIDAD]`: Menor o igual. Ej: `<=10.5w`
  * La unidad (ej: W, V, A, VATIOS) es opcional. Si se usa, debe coincidir en la celda.

- Rangos numéricos (con o sin unidad opcional, ambos extremos incluidos):
  * `num1-num2[UNIDAD]`: Entre num1 y num2. Ej: `10-20` o `50-100V`
  * Si se usa unidad, debe coincidir en la celda.

- Negación (excluir filas que coincidan con el término negado):
  * `#término`: 'término' puede ser texto, comparación o rango.
    Ej: `switch + #gestionable`
    Ej: `tarjeta + #>8 puertos` (negación aplica a ">8 puertos")

Búsqueda Directa (si el término no está en diccionario o si el usuario elige):
--------------------------------------------------------------------------
Se busca directamente en las descripciones usando la misma lógica de operadores,
comparaciones, rangos y negación que para el diccionario.
- `texto`: Busca el texto.
- `t1 + t2`: Descripciones con AMBOS términos.
- `t1 | t2` o `t1 / t2`: Descripciones con AL MENOS UNO.
- `#termino_directo` / `>100AMPS_directo` / `10-20V_directo`

Notas Generales:
- Búsqueda insensible a mayúsculas/minúsculas y acentos.
- Términos de texto se buscan como palabras completas.
- La negación (#) aplica al término que le sigue inmediatamente.
- Los operadores lógicos (+, |) tienen menor precedencia que la negación.
  Ej: `SAI + #SOBREMESA` busca SAI y que NO sea SOBREMESA.
"""
        messagebox.showinfo("Ayuda - Sintaxis de Búsqueda", ayuda)

    def _configurar_tags_treeview(self):
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par)
            tabla.tag_configure('impar', background=self.color_fila_impar)

    def _configurar_orden_tabla(self, tabla: ttk.Treeview):
        cols = tabla["columns"]
        if cols:
            for col in cols:
                tabla.heading(col, text=str(col), anchor=tk.W,
                              command=lambda c=col, t=tabla: self._ordenar_columna(t, c, False))

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool):
        # (Adaptado de tu lógica, mejorado para DataFrames)
        df_para_ordenar = None
        if tabla == self.tabla_diccionario:
            df_para_ordenar = self.motor.datos_diccionario
        elif tabla == self.tabla_resultados:
            df_para_ordenar = self.resultados_actuales
        
        if df_para_ordenar is None or df_para_ordenar.empty or col not in df_para_ordenar.columns:
            logging.debug(f"No se puede ordenar la tabla por columna '{col}'. Datos no disponibles o columna inexistente.")
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, not reverse)) # Invertir para siguiente click
            return

        logging.info(f"Ordenando tabla por columna '{col}', descendente={reverse}")
        try:
            # Convertir a numérico si es posible para una ordenación más natural
            if pd.api.types.is_numeric_dtype(df_para_ordenar[col].infer_objects().dtype):
                df_ordenado = df_para_ordenar.sort_values(by=col, ascending=not reverse, na_position='last')
            else: # Ordenar como string (case-insensitive)
                df_ordenado = df_para_ordenar.sort_values(by=col, ascending=not reverse, na_position='last', key=lambda x: x.astype(str).str.lower())
            
            if tabla == self.tabla_diccionario:
                self.motor.datos_diccionario = df_ordenado # Actualiza el DataFrame del motor
                self._actualizar_tabla(tabla, df_ordenado, limite_filas=100) # Mostrar preview actualizada
            elif tabla == self.tabla_resultados:
                self.resultados_actuales = df_ordenado
                self._actualizar_tabla(tabla, df_ordenado)
            
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, not reverse))
            self._actualizar_estado(f"Tabla ordenada por '{col}' ({'Asc' if not reverse else 'Desc'}).")
        except Exception as e:
            logging.exception(f"Error al intentar ordenar por columna '{col}'")
            messagebox.showerror("Error al Ordenar", f"No se pudo ordenar por '{col}':\n{e}")
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, False)) # Reset

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None, columnas_a_mostrar: Optional[List[str]] = None):
        # (Lógica de tu UI, con pequeñas adaptaciones)
        is_diccionario = tabla == self.tabla_diccionario
        logger.debug(f"Actualizando tabla {'Diccionario' if is_diccionario else 'Resultados'}.")
        try:
            for i in tabla.get_children(): tabla.delete(i)
        except tk.TclError as e: logger.warning(f"Error Tcl al limpiar tabla: {e}"); pass
        tabla["columns"] = ()

        if datos is None or datos.empty:
            logger.debug("No hay datos para mostrar en la tabla.")
            return

        datos_a_usar = datos
        cols_df = list(datos_a_usar.columns)

        if columnas_a_mostrar:
            cols_finales = [c for c in columnas_a_mostrar if c in cols_df]
            if not cols_finales: cols_finales = cols_df # Fallback a todas si ninguna de las especificadas existe
        else:
            cols_finales = cols_df
        
        if not cols_finales:
            logger.warning("DataFrame no tiene columnas para mostrar o columnas seleccionadas no existen.")
            return

        df_para_mostrar_vista = datos_a_usar[cols_finales]
        tabla["columns"] = tuple(cols_finales)

        for col in cols_finales:
            tabla.heading(col, text=str(col), anchor=tk.W)
            try:
                col_as_str = df_para_mostrar_vista[col].dropna().astype(str)
                ancho_contenido = col_as_str.str.len().max() if not col_as_str.empty else 0
                ancho_cabecera = len(str(col))
                ancho = max(70, min(int(max(ancho_cabecera * 8, ancho_contenido * 6.5) + 25), 400))
                tabla.column(col, anchor=tk.W, width=ancho, minwidth=70)
            except Exception as e:
                logger.warning(f"Error calculando ancho para col '{col}': {e}. Ancho default.")
                tabla.column(col, anchor=tk.W, width=100, minwidth=50)

        num_filas_a_iterar = limite_filas if is_diccionario and limite_filas is not None else len(df_para_mostrar_vista)
        df_iterar = df_para_mostrar_vista.head(num_filas_a_iterar)

        for i, (_, row) in enumerate(df_iterar.iterrows()):
            vals = [str(v) if pd.notna(v) else "" for v in row.values]
            tag = 'par' if i % 2 == 0 else 'impar'
            try:
                tabla.insert("", "end", values=vals, tags=(tag,))
            except tk.TclError as e:
                logger.warning(f"Error Tcl insertando fila {i}: {e}. Intentando con ASCII.")
                try:
                    vals_ascii = [v.encode('ascii', 'ignore').decode('ascii') for v in vals]
                    tabla.insert("", "end", values=vals_ascii, tags=(tag,))
                except Exception as e_inner:
                    logger.error(f"Fallo el fallback ASCII para fila {i}: {e_inner}")
        # No es necesario reconfigurar orden aquí si las columnas son las mismas.

    def _actualizar_etiquetas_archivos(self):
        dic_path = self.motor.archivo_diccionario_actual
        desc_path = self.motor.archivo_descripcion_actual
        dic_name = dic_path.name if dic_path else "Ninguno"
        desc_name = desc_path.name if desc_path else "Ninguno"
        
        max_len_label = 25 # Ajusta según tu Label width
        dic_display = f"Dic: {dic_name}" if len(dic_name) <= max_len_label else f"Dic: ...{dic_name[-(max_len_label-4):]}"
        desc_display = f"Desc: {desc_name}" if len(desc_name) <= max_len_label else f"Desc: ...{desc_name[-(max_len_label-4):]}"
        
        self.lbl_dic_cargado.config(text=dic_display, foreground="green" if dic_path else "red")
        self.lbl_desc_cargado.config(text=desc_display, foreground="green" if desc_path else "red")

    def _actualizar_botones_estado_general(self):
        # (Tu lógica, adaptada ligeramente)
        dic_cargado = self.motor.datos_diccionario is not None #and not self.motor.datos_diccionario.empty (puede estar vacío y ser válido)
        desc_cargado = self.motor.datos_descripcion is not None #and not self.motor.datos_descripcion.empty

        if dic_cargado: self._actualizar_estado_botones_operadores()
        else: self._deshabilitar_botones_operadores()

        self.btn_buscar['state'] = 'normal' if dic_cargado and desc_cargado else 'disabled'
        
        puede_salvar_algo = False
        if self.ultimo_termino_buscado and self.origen_principal_resultados != OrigenResultados.NINGUNO:
            if self.origen_principal_resultados.es_via_diccionario:
                if (self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty) or \
                   (self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty and \
                    self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC):
                    puede_salvar_algo = True
            elif self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
                if self.df_candidato_descripcion is not None: # Puede estar vacío si se guardó "todo"
                    puede_salvar_algo = True
        
        self.btn_salvar_regla['state'] = 'normal' if puede_salvar_algo else 'disabled'
        self.btn_exportar['state'] = 'normal' if self.reglas_guardadas else 'disabled'

    def _cargar_diccionario(self):
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
        self._actualizar_tabla(self.tabla_resultados, None) # Limpiar resultados también
        self.resultados_actuales = None; self.df_candidato_diccionario = None; self.df_candidato_descripcion = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        if self.motor.cargar_excel_diccionario(ruta):
            self.config["last_dic_path"] = ruta # Guardar la nueva ruta
            self._guardar_configuracion() # Guardar inmediatamente
            df_dic = self.motor.datos_diccionario
            if df_dic is not None:
                num_filas = len(df_dic)
                cols_busqueda_nombres = self.motor._obtener_nombres_columnas_busqueda_df(df_dic, self.motor.indices_columnas_busqueda_dic, "diccionario (preview)")
                
                indices_str = ', '.join(map(str, self.motor.indices_columnas_busqueda_dic)) if self.motor.indices_columnas_busqueda_dic and self.motor.indices_columnas_busqueda_dic != [-1] else "Todas Texto"
                lbl_text_dic_preview = f"Vista Previa Dic ({num_filas} filas)"
                if cols_busqueda_nombres:
                    lbl_text_dic_preview = f"Vista Previa Diccionario (Cols: {', '.join(cols_busqueda_nombres)} - Índices: {indices_str})"
                self.lbl_tabla_diccionario.config(text=lbl_text_dic_preview)
                
                self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols_busqueda_nombres)
                self.title(f"Buscador - Dic: {nombre_archivo}")
                self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado.")
        else:
            self._actualizar_estado("Error al cargar el diccionario.")
            self.title("Buscador Avanzado PyGuru Mod")
        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _cargar_excel_descripcion(self):
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
        self.resultados_actuales = None; self.df_candidato_descripcion = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        if self.motor.cargar_excel_descripcion(ruta):
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
            self._actualizar_estado("Error al cargar las descripciones.")
        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _parsear_termino_busqueda_inicial(self, termino_raw: str) -> Tuple[str, List[str]]:
        # (Tu lógica de parseo inicial de la UI)
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
            op_principal = 'AND' # Si es un solo término, se trata como AND (debe cumplir ese término)
            terminos_brutos = [termino_limpio]

        if not any(terminos_brutos):
            logger.warning(f"Término '{termino_raw}' vacío tras parseo inicial."); return 'OR', []
        return op_principal, terminos_brutos

    def _ejecutar_busqueda(self):
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", "Cargue Diccionario y Descripciones.")
            return

        termino_busqueda_actual = self.texto_busqueda_var.get()
        self.ultimo_termino_buscado = termino_busqueda_actual

        self.resultados_actuales = None
        self._actualizar_tabla(self.tabla_resultados, None)
        self.df_candidato_diccionario = None
        self.df_candidato_descripcion = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        if not termino_busqueda_actual.strip():
            logger.info("Búsqueda vacía. Mostrando todas las descripciones.")
            df_desc_all = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc_all)
            self.resultados_actuales = df_desc_all.copy() if df_desc_all is not None else pd.DataFrame()
            self.df_candidato_descripcion = self.resultados_actuales
            self.origen_principal_resultados = OrigenResultados.DIRECTO_DESCRIPCION_VACIA # Indica que se mostraron todas
            self._actualizar_estado(f"Mostrando todas las {len(df_desc_all) if df_desc_all is not None else 0} descripciones.")
            self._actualizar_botones_estado_general()
            return

        self._actualizar_estado(f"Buscando '{termino_busqueda_actual}'...")
        
        # Usar copias para la búsqueda para no alterar los DFs cargados en el motor
        df_dic_copia = self.motor.datos_diccionario.copy() if self.motor.datos_diccionario is not None else None
        df_desc_copia = self.motor.datos_descripcion.copy() if self.motor.datos_descripcion is not None else None

        if df_dic_copia is None or df_desc_copia is None : # Doble check
             messagebox.showerror("Error Interno", "Los DataFrames base no están disponibles para búsqueda.")
             self._actualizar_estado("Error interno en DataFrames."); return


        op_principal_usuario, terminos_brutos_usuario = self._parsear_termino_busqueda_inicial(termino_busqueda_actual)
        if not terminos_brutos_usuario:
            messagebox.showwarning("Término Inválido", "El término de búsqueda está vacío o malformado.")
            self._actualizar_estado("Término de búsqueda inválido."); self._actualizar_botones_estado_general(); return
            
        terminos_analizados_usuario = self.motor._analizar_terminos(terminos_brutos_usuario)
        if not terminos_analizados_usuario:
            messagebox.showwarning("Término Inválido", f"No se pudieron analizar los términos en '{termino_busqueda_actual}'. Verifique sintaxis.")
            self._actualizar_estado(f"Análisis de '{termino_busqueda_actual}' fallido."); self._actualizar_botones_estado_general(); return

        # Decidir si la búsqueda es vía diccionario
        # Aquí tu lógica original era un poco diferente.
        # Vamos a asumir que si el término NO produce FCDs, entonces se pregunta al usuario si quiere búsqueda directa.
        
        resultados_busqueda, origen_res, fcds = self.motor.buscar(
            df_dic_copia, df_desc_copia, terminos_analizados_usuario, op_principal_usuario, buscar_via_diccionario_flag=True
        )
        
        self.df_candidato_diccionario = fcds # Guardar FCDs si los hubo
        self.origen_principal_resultados = origen_res

        if origen_res == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC:
            self.resultados_actuales = resultados_busqueda
            self.df_candidato_descripcion = resultados_busqueda.copy() if resultados_busqueda is not None else pd.DataFrame()
            self._actualizar_estado(f"'{termino_busqueda_actual}': {len(fcds) if fcds is not None else 0} en Dic, {len(resultados_busqueda) if resultados_busqueda is not None else 0} en Desc.")
        elif origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC:
            self.resultados_actuales = pd.DataFrame(columns=df_desc_copia.columns) # Vacío pero con estructura
            self.df_candidato_descripcion = self.resultados_actuales.copy()
            self._actualizar_estado(f"'{termino_busqueda_actual}': {len(fcds) if fcds is not None else 0} en Dic, 0 en Desc.")
            messagebox.showinfo("Información", f"Se encontraron {len(fcds) if fcds is not None else 0} filas en Diccionario para '{termino_busqueda_actual}', pero ninguna coincidencia en Descripciones.")
        elif origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS:
            self.resultados_actuales = pd.DataFrame(columns=df_desc_copia.columns)
            self.df_candidato_descripcion = self.resultados_actuales.copy()
            self._actualizar_estado(f"'{termino_busqueda_actual}': {len(fcds) if fcds is not None else 0} en Dic, sin términos válidos para Desc.")
            messagebox.showinfo("Información", f"Se encontraron {len(fcds) if fcds is not None else 0} filas en Diccionario, pero no se extrajeron términos válidos para buscar en descripciones.")

        # Fallback a búsqueda directa si la vía diccionario no dio resultados EN DESCRIPCIONES
        # o si no se encontraron FCDs en absoluto.
        if (origen_res.es_via_diccionario and (resultados_busqueda is None or resultados_busqueda.empty)) or \
           (fcds is None or fcds.empty and origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC): # Si no hubo FCDs
            
            mensaje_fallback = f"'{termino_busqueda_actual}' no produjo resultados finales vía diccionario."
            if fcds is None or fcds.empty :
                 mensaje_fallback = f"'{termino_busqueda_actual}' no se encontró en el Diccionario."
            
            if messagebox.askyesno("Búsqueda Alternativa",
                                   f"{mensaje_fallback}\n\n"
                                   f"¿Desea buscar '{termino_busqueda_actual}' directamente en las Descripciones?"):
                
                resultados_directos, origen_directo, _ = self.motor.buscar(
                    df_dic_copia, df_desc_copia, terminos_analizados_usuario, op_principal_usuario, buscar_via_diccionario_flag=False
                )
                self.resultados_actuales = resultados_directos
                self.df_candidato_descripcion = resultados_directos.copy() if resultados_directos is not None else pd.DataFrame()
                self.origen_principal_resultados = origen_directo # Actualizar origen principal
                self.df_candidato_diccionario = None # Ya no aplica si la búsqueda fue directa

                num_rdd = len(self.resultados_actuales) if self.resultados_actuales is not None else 0
                self._actualizar_estado(f"Búsqueda directa '{termino_busqueda_actual}': {num_rdd} resultados.")
                if num_rdd == 0:
                    messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_busqueda_actual}' en búsqueda directa.")
            else:
                # Si el usuario dice NO a la búsqueda directa, y no hubo resultados vía diccionario
                if self.resultados_actuales is None or self.resultados_actuales.empty:
                     self._actualizar_estado(f"Búsqueda de '{termino_busqueda_actual}' cancelada/sin resultados.")
                     self.resultados_actuales = pd.DataFrame(columns=df_desc_copia.columns) # Mostrar tabla vacía
                     self.df_candidato_descripcion = self.resultados_actuales.copy()
                     self.origen_principal_resultados = OrigenResultados.NINGUNO # O el último origen de error/sin_resultados
        
        self._actualizar_tabla(self.tabla_resultados, self.resultados_actuales)
        self._actualizar_botones_estado_general()

        if self.motor.datos_diccionario is not None and not self.motor.datos_diccionario.empty:
            self._buscar_y_enfocar_en_preview()

    def _buscar_y_enfocar_en_preview(self):
        # (Lógica de tu UI, con pequeñas adaptaciones)
        termino_buscar_raw = self.texto_busqueda_var.get()
        if not termino_buscar_raw: return
        
        op_principal, terminos_brutos = self._parsear_termino_busqueda_inicial(termino_buscar_raw)
        if not terminos_brutos: return
        termino_enfocar = terminos_brutos[0] 
        
        analizado_enfocar_lista = self.motor._analizar_terminos([termino_enfocar])
        termino_enfocar_valor_str = termino_enfocar # Fallback
        if analizado_enfocar_lista:
            # Si es numérico o rango, el valor es float o list. Convertir a str para búsqueda en Treeview.
            val_analizado = analizado_enfocar_lista[0].get('valor')
            if isinstance(val_analizado, (float, int)):
                termino_enfocar_valor_str = str(val_analizado)
            elif isinstance(val_analizado, list) and len(val_analizado) == 2: # Rango
                termino_enfocar_valor_str = f"{val_analizado[0]}-{val_analizado[1]}"
            elif isinstance(val_analizado, str):
                 termino_enfocar_valor_str = val_analizado # Ya es string y upper
            else: # Si es otro tipo o None, usar el original
                termino_enfocar_valor_str = termino_enfocar.upper() # Asegurar upper para comparación
        else:
             termino_enfocar_valor_str = termino_enfocar.upper()


        if not termino_enfocar_valor_str: return

        items_preview = self.tabla_diccionario.get_children('')
        if not items_preview or self.motor.datos_diccionario is None: return

        termino_upper_para_treeview = termino_enfocar_valor_str # Ya debería estar en upper
        logger.info(f"Intentando enfocar '{termino_upper_para_treeview}' en vista previa del diccionario...")

        found_item_id = None
        for item_id in items_preview:
            try:
                valores_fila = self.tabla_diccionario.item(item_id, 'values')
                if any(termino_upper_para_treeview in str(val).upper() for val in valores_fila if val is not None):
                    found_item_id = item_id; break
            except Exception as e:
                logger.warning(f"Error procesando item {item_id} en preview (búsqueda treeview): {e}"); continue
        
        # (La lógica de buscar en el DF completo y recargar la tabla es compleja y se omite aquí
        # para mantener la fusión más directa con tu UI original. Requeriría paginación.)

        if found_item_id:
            logger.info(f"Término '{termino_upper_para_treeview}' enfocado en preview (item ID: {found_item_id}).")
            try:
                current_selection = self.tabla_diccionario.selection()
                if current_selection: self.tabla_diccionario.selection_remove(current_selection)
                self.tabla_diccionario.selection_set(found_item_id)
                self.tabla_diccionario.see(found_item_id)
                self.tabla_diccionario.focus(found_item_id)
            except Exception as e:
                logger.error(f"Error al enfocar item {found_item_id} en preview: {e}")
        else:
            logger.info(f"Término '{termino_upper_para_treeview}' no encontrado/enfocado en vista previa del diccionario.")

    def _salvar_regla_actual(self):
        # (Tu lógica de salvar regla, adaptada para usar el nuevo origen y datos)
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
            'termino_busqueda_parseado': terminos_analizados_regla, # Guardar la estructura analizada
            'operador_principal': op_regla,
            'fuente_original_guardado': origen_nombre_actual,
            'timestamp': pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        salvo_algo = False

        if self.origen_principal_resultados.es_via_diccionario:
            decision = self._mostrar_dialogo_seleccion_salvado_via_diccionario()
            if decision['confirmed']:
                if decision['save_fcd'] and self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty:
                    regla_fcd = {**regla_a_guardar_base, 'tipo_datos_guardados': "COINCIDENCIAS_DICCIONARIO",
                                 'datos_snapshot': self.df_candidato_diccionario.to_dict(orient='records')}
                    self.reglas_guardadas.append(regla_fcd); salvo_algo = True
                    logger.info(f"Regla (FCD) salvada para: '{termino_original_regla}'")
                if decision['save_rfd'] and self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty and \
                   self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC:
                    regla_rfd = {**regla_a_guardar_base, 'tipo_datos_guardados': "RESULTADOS_DESCRIPCION_VIA_DICCIONARIO",
                                 'datos_snapshot': self.df_candidato_descripcion.to_dict(orient='records')}
                    self.reglas_guardadas.append(regla_rfd); salvo_algo = True
                    logger.info(f"Regla (Resultados Desc vía Dic) salvada para: '{termino_original_regla}'")
        elif self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            if self.df_candidato_descripcion is not None: # Puede ser un DF vacío si se guardó "todo"
                tipo_datos = "TODAS_LAS_DESCRIPCIONES" if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA else "RESULTADOS_DESCRIPCION_DIRECTA"
                regla_directa = {**regla_a_guardar_base, 'tipo_datos_guardados': tipo_datos,
                                 'datos_snapshot': self.df_candidato_descripcion.to_dict(orient='records')}
                self.reglas_guardadas.append(regla_directa); salvo_algo = True
                logger.info(f"Regla ({tipo_datos}) salvada para: '{termino_original_regla}'")
            else: messagebox.showwarning("Sin Datos", "No hay resultados para salvar.")
        else:
            if self.origen_principal_resultados != OrigenResultados.NINGUNO:
                messagebox.showerror("Error", f"No se puede determinar qué salvar para origen: {origen_nombre_actual}.")

        if salvo_algo: self._actualizar_estado(f"Regla(s) nueva(s) guardada(s). Total: {len(self.reglas_guardadas)}.")
        else: self._actualizar_estado("Ninguna regla fue salvada.")
        self._actualizar_botones_estado_general()

    def _mostrar_dialogo_seleccion_salvado_via_diccionario(self) -> Dict[str, bool]:
        # (Tu lógica, sin cambios)
        decision = {'confirmed': False, 'save_fcd': False, 'save_rfd': False}
        # Crear un TopLevel para el diálogo personalizado
        dialog = tk.Toplevel(self)
        dialog.title("Seleccionar Datos a Salvar")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        dialog.transient(self) # Hacerlo modal sobre la ventana principal
        dialog.grab_set()      # Capturar eventos

        var_fcd = tk.BooleanVar(value=True if self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty else False)
        var_rfd = tk.BooleanVar(value=True if self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty and self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC else False)

        lbl_msg = ttk.Label(dialog, text="La búsqueda fue vía Diccionario. ¿Qué datos desea salvar para esta regla?")
        lbl_msg.pack(pady=10, padx=10)

        chk_fcd = ttk.Checkbutton(dialog, text="Coincidencias del Diccionario (FCDs)", variable=var_fcd)
        chk_fcd.pack(anchor=tk.W, padx=20)
        chk_fcd['state'] = 'normal' if self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty else 'disabled'


        chk_rfd = ttk.Checkbutton(dialog, text="Resultados Finales de Descripciones (RFDs)", variable=var_rfd)
        chk_rfd.pack(anchor=tk.W, padx=20)
        chk_rfd['state'] = 'normal' if self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty and self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC else 'disabled'
        
        frame_botones_dialogo = ttk.Frame(dialog)
        frame_botones_dialogo.pack(pady=15)

        def on_confirm():
            decision['confirmed'] = True
            decision['save_fcd'] = var_fcd.get()
            decision['save_rfd'] = var_rfd.get()
            if not decision['save_fcd'] and not decision['save_rfd']:
                 messagebox.showwarning("Nada Seleccionado", "No ha seleccionado ningún conjunto de datos para salvar.", parent=dialog)
                 decision['confirmed'] = False # No cerrar si no seleccionó nada útil
                 return # No cerrar
            dialog.destroy()

        def on_cancel():
            decision['confirmed'] = False
            dialog.destroy()

        btn_confirm = ttk.Button(frame_botones_dialogo, text="Confirmar", command=on_confirm)
        btn_confirm.pack(side=tk.LEFT, padx=10)
        btn_cancel = ttk.Button(frame_botones_dialogo, text="Cancelar", command=on_cancel)
        btn_cancel.pack(side=tk.LEFT, padx=10)
        
        # Centrar el diálogo
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        self.wait_window(dialog) # Esperar a que el diálogo se cierre
        return decision

    def _exportar_resultados(self):
        # (Tu lógica, adaptada para la nueva estructura de reglas)
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
                    nombre_hoja_def = f"Def_{id_regla_hoja}"[:31]
                    df_regla_definicion.to_excel(writer, sheet_name=nombre_hoja_def, index=False)

                    datos_snapshot_list = regla_guardada.get('datos_snapshot')
                    if datos_snapshot_list:
                        df_snapshot = pd.DataFrame(datos_snapshot_list)
                        if not df_snapshot.empty:
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
        # (Tu lógica, sin cambios)
        if not texto: return "resultados"
        sane = re.sub(r'[<>:"/\\|?*\s]+', '_', texto)
        sane = sane.strip('_')
        return sane[:max_len] if len(sane) > max_len else sane

    def _actualizar_estado_botones_operadores(self):
        # (Tu lógica, sin cambios)
        if self.motor.datos_diccionario is None:
            self._deshabilitar_botones_operadores(); return

        texto_completo = self.texto_busqueda_var.get()
        cursor_pos = self.entrada_busqueda.index(tk.INSERT)
        segmento_antes_cursor = texto_completo[:cursor_pos]
        ultimo_op_logico_pos = -1
        for op_s in ['+', '|', '/']:
            pos = segmento_antes_cursor.rfind(op_s)
            if pos > ultimo_op_logico_pos: ultimo_op_logico_pos = pos
        termino_actual_bruto = segmento_antes_cursor[ultimo_op_logico_pos+1:].lstrip()

        terminos_analizados_actual = self.motor._analizar_terminos([termino_actual_bruto]) # Usa el motor
        tipo_termino_actual = None; tiene_negacion_actual = False
        if terminos_analizados_actual:
            analisis_primero = terminos_analizados_actual[0]
            tipo_termino_actual = analisis_primero.get('tipo')
            if analisis_primero.get('negate', False): tiene_negacion_actual = True

        self.btn_not['state'] = 'disabled' if tiene_negacion_actual or not termino_actual_bruto.strip() else 'normal'
        es_comparativo_o_rango_actual = tipo_termino_actual in ['gt', 'lt', 'ge', 'le', 'range']
        estado_comparacion = 'disabled' if es_comparativo_o_rango_actual or (tiene_negacion_actual and not termino_actual_bruto[1:].strip()) else 'normal'
        for btn in [self.btn_gt, self.btn_lt, self.btn_ge, self.btn_le]: btn.config(state=estado_comparacion)

        puede_poner_rango = False
        if termino_actual_bruto.strip() and not tiene_negacion_actual and not es_comparativo_o_rango_actual:
            if termino_actual_bruto[-1].isdigit() or termino_actual_bruto[-1].isalpha() : puede_poner_rango = True
        self.btn_range.config(state='normal' if puede_poner_rango else 'disabled')

        estado_logico = 'normal'; texto_limpio_final_total = texto_completo.rstrip()
        if not texto_completo.strip(): estado_logico = 'disabled'
        elif texto_limpio_final_total and texto_limpio_final_total[-1] in ['+', '|', '/']: estado_logico = 'disabled'
        elif not termino_actual_bruto.strip() and ultimo_op_logico_pos != -1 : estado_logico = 'disabled'
        self.btn_and.config(state=estado_logico); self.btn_or.config(state=estado_logico)

    def _insertar_operador_validado(self, operador: str):
        # (Tu lógica, sin cambios)
        if self.motor.datos_diccionario is None: return
        texto_actual = self.texto_busqueda_var.get()
        cursor_pos = self.entrada_busqueda.index(tk.INSERT)
        char_antes_cursor = texto_actual[cursor_pos-1:cursor_pos] if cursor_pos > 0 else ""
        puede_insertar = True
        
        # Esta validación es simple; la activación/desactivación de botones es la principal.
        if operador in ['+', '|', '/']:
            if not texto_actual.strip() or char_antes_cursor in ['+', '|', '/',' ']: puede_insertar = False
            operador = f" {operador} "
        elif operador == '#':
            if char_antes_cursor.strip() and char_antes_cursor not in ['+', '|', '/']: puede_insertar = False
        
        if puede_insertar: # Dejar que la lógica de _actualizar_estado_botones_operadores decida si está realmente activo
            self.entrada_busqueda.insert(tk.INSERT, operador)

    def _deshabilitar_botones_operadores(self):
        for btn_op in [self.btn_and, self.btn_or, self.btn_not, self.btn_gt, self.btn_lt, self.btn_ge, self.btn_le, self.btn_range]:
            btn_op['state'] = 'disabled'

    def on_closing(self):
        logger.info("Cerrando la aplicación...")
        self._guardar_configuracion()
        self.destroy()


if __name__ == "__main__":
    log_file = 'buscador_app_fUSION.log' # Nuevo nombre de log
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8', mode='a'), # 'a' para append
            logging.StreamHandler()
        ]
    )
    logger.info("=============================================")
    logger.info("=== Iniciando Aplicación Buscador (Fusionada) ===")
    
    missing_deps = []
    try: import pandas as pd; logger.info(f"Pandas versión: {pd.__version__}")
    except ImportError: missing_deps.append("pandas"); logger.critical("Dependencia faltante: pandas")
    try: import openpyxl; logger.info(f"openpyxl versión: {openpyxl.__version__}")
    except ImportError: missing_deps.append("openpyxl (para .xlsx)"); # No es crítico si solo usas .xls
    # numpy no se importa explícitamente aquí a menos que lo necesites directamente.

    if "pandas" in missing_deps: # Pandas es crítico
        error_msg = f"Faltan librerías críticas: {', '.join(missing_deps)}.\nInstale con: pip install {' '.join(missing_deps)}"
        logger.critical(error_msg)
        try:
            root_temp = tk.Tk(); root_temp.withdraw(); messagebox.showerror("Dependencias Faltantes", error_msg); root_temp.destroy()
        except tk.TclError: print(f"ERROR CRÍTICO: {error_msg}")
        exit(1)

    try:
        app = InterfazGrafica()
        app.mainloop()
    except Exception as main_error:
        logger.critical("¡Error fatal no capturado en la aplicación!", exc_info=True)
        try:
            root_err = tk.Tk(); root_err.withdraw()
            messagebox.showerror("Error Fatal", f"Error crítico:\n{main_error}\nConsulte '{log_file}'.")
            root_err.destroy()
        except Exception as fallback_error:
            logger.error(f"No se pudo mostrar el mensaje de error fatal: {fallback_error}")
            print(f"ERROR FATAL: {main_error}. Consulte {log_file}.")
    finally:
        logger.info("=== Finalizando Aplicación Buscador ===")
