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
    Dict,
    Any,
)
from enum import Enum, auto
import platform
import unicodedata
import logging
import json
import os
from pathlib import Path
import string
import numpy as np

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

    @property
    def es_via_diccionario(self) -> bool:
        return self in {
            OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
            OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC,
            OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS,
        }
    @property
    def es_directo_descripcion(self) -> bool:
        return self in {OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, OrigenResultados.DIRECTO_DESCRIPCION_VACIA,}
    @property
    def es_error_carga(self) -> bool:
        return self in {OrigenResultados.ERROR_CARGA_DICCIONARIO, OrigenResultados.ERROR_CARGA_DESCRIPCION,}
    @property
    def es_error_configuracion(self) -> bool:
        return self in {OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC,}
    @property
    def es_error_operacional(self) -> bool: return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR
    @property
    def es_termino_invalido(self) -> bool: return self == OrigenResultados.TERMINO_INVALIDO

class ExtractorMagnitud:
    MAPEO_MAGNITUDES_PREDEFINIDO = {} 
    def __init__(self, mapeo_magnitudes: Optional[Dict[str, List[str]]] = None):
        self.sinonimo_a_canonico_normalizado: Dict[str, str] = {}
        mapeo_a_usar = mapeo_magnitudes if mapeo_magnitudes is not None else self.MAPEO_MAGNITUDES_PREDEFINIDO
        for forma_canonica, lista_sinonimos in mapeo_a_usar.items():
            canonico_norm = self._normalizar_texto(forma_canonica)
            if not canonico_norm: logger.warning(f"Forma canónica '{forma_canonica}' vacía tras normalizar."); continue
            self.sinonimo_a_canonico_normalizado[canonico_norm] = canonico_norm
            for sinonimo in lista_sinonimos:
                sinonimo_norm = self._normalizar_texto(sinonimo)
                if sinonimo_norm: self.sinonimo_a_canonico_normalizado[sinonimo_norm] = canonico_norm
    @staticmethod
    def _normalizar_texto(texto: str) -> str: 
        if not isinstance(texto, str) or not texto: return ""
        try:
            texto_upper = texto.upper()
            forma_normalizada = unicodedata.normalize("NFKD", texto_upper)
            res = "".join(c for c in forma_normalizada if not unicodedata.combining(c) and (c.isalnum() or c.isspace() or c in ['.', '-', '_', '/']))
            return ' '.join(res.split())
        except TypeError: logger.error(f"TypeError en _normalizar_texto (ExtractorMagnitud) con: {texto}"); return ""
    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        if not texto_unidad: return None
        normalizada = self._normalizar_texto(texto_unidad)
        return self.sinonimo_a_canonico_normalizado.get(normalizada) if normalizada else None

class ManejadorExcel:
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        ruta = Path(ruta_archivo)
        if not ruta.exists(): logger.error(f"ManejadorExcel: Archivo no encontrado: {ruta}"); return None, f"¡Archivo no encontrado! Ruta: {ruta}"
        try:
            engine = "openpyxl" if ruta.suffix.lower() == ".xlsx" else None 
            logger.info(f"ManejadorExcel: Cargando '{ruta.name}' con engine='{engine or 'auto (intentará xlrd para .xls)'}'...")
            df = pd.read_excel(ruta, engine=engine)
            logger.info(f"ManejadorExcel: Archivo '{ruta.name}' ({len(df)} filas) cargado.")
            return df, None
        except ImportError as ie: 
            logger.exception(f"ManejadorExcel: Falta dependencia para leer '{ruta.name}'. Error: {ie}")
            return None, (f"Error al cargar '{ruta.name}': Falta librería.\n"
                          f"Para .xlsx: pip install openpyxl\nPara .xls: pip install xlrd\nDetalle: {ie}")
        except Exception as e: 
            logger.exception(f"ManejadorExcel: Error genérico al cargar '{ruta.name}'.")
            return None, (f"No se pudo cargar '{ruta.name}': {e}\nVerifique formato, permisos y si está en uso.")

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
        self.extractor_magnitud = ExtractorMagnitud() 

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)
        if df_cargado is None:
            self.datos_diccionario = None; self.archivo_diccionario_actual = None
            return False, error_msg_carga 
        col0_vals = df_cargado.iloc[:, 0].dropna().astype(str).unique() if df_cargado.shape[1] > 0 else []
        mapeo_dinamico = {val.strip(): [val.strip()] for val in col0_vals if val.strip()}
        self.extractor_magnitud = ExtractorMagnitud(mapeo_magnitudes=mapeo_dinamico)
        logger.info(f"Extractor de magnitudes actualizado desde '{ruta.name}'.")
        self.datos_diccionario = df_cargado; self.archivo_diccionario_actual = ruta
        if logger.isEnabledFor(logging.DEBUG) and self.datos_diccionario is not None:
             logger.debug(f"Diccionario '{ruta.name}' cargado (primeras 3 filas):\n{self.datos_diccionario.head(3).to_string()}")
        return True, None

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)
        if df_cargado is None:
            self.datos_descripcion = None; self.archivo_descripcion_actual = None
            return False, error_msg_carga 
        self.datos_descripcion = df_cargado; self.archivo_descripcion_actual = ruta
        logger.info(f"Archivo de descripciones '{ruta.name}' cargado.")
        return True, None
    
    def _obtener_nombres_columnas_busqueda_df(self, df: pd.DataFrame, indices_cfg: List[int], tipo_busqueda: str) -> Tuple[Optional[List[str]], Optional[str]]:
        if df is None or df.empty: return None, f"DF para '{tipo_busqueda}' vacío."
        columnas_disponibles = list(df.columns); num_cols_df = len(columnas_disponibles)
        if num_cols_df == 0: return None, f"DF '{tipo_busqueda}' sin columnas."
        if tipo_busqueda == "diccionario_fcds_inicial": 
            logger.debug(f"'{tipo_busqueda}': Usando todas las {num_cols_df} columnas: {columnas_disponibles}")
            return columnas_disponibles, None
        if not indices_cfg or indices_cfg == [-1]: 
            cols_texto_obj = [col for col in columnas_disponibles if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
            if cols_texto_obj: logger.debug(f"'{tipo_busqueda}': Usando cols texto/obj: {cols_texto_obj}"); return cols_texto_obj, None
            logger.warning(f"No hay cols texto/obj en '{tipo_busqueda}', usando todas: {columnas_disponibles}"); return columnas_disponibles, None
        nombres_seleccionadas = [] 
        for i in indices_cfg:
            if not (isinstance(i, int) and 0 <= i < num_cols_df): return None, f"Índice {i} inválido para '{tipo_busqueda}'."
            nombres_seleccionadas.append(columnas_disponibles[i])
        if not nombres_seleccionadas and indices_cfg: return None, f"Índices {indices_cfg} no válidos para '{tipo_busqueda}'."
        logger.debug(f"'{tipo_busqueda}': Usando columnas por índice {indices_cfg}: {nombres_seleccionadas}")
        return nombres_seleccionadas, None
    
    def _normalizar_para_busqueda(self, texto: str) -> str:
        if not isinstance(texto, str) or not texto: return ""
        try:
            texto_upper = texto.upper()
            texto_norm_nfkd = unicodedata.normalize('NFKD', texto_upper)
            texto_sin_acentos = "".join([c for c in texto_norm_nfkd if not unicodedata.combining(c)])
            return ' '.join(texto_sin_acentos.split()).strip()
        except Exception as e: logger.error(f"Error normalizando '{texto[:50]}...': {e}"); return str(texto).upper().strip()

    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, cols: List[str], texto: str) -> Tuple[pd.DataFrame, str]:
        texto_limpio = texto.strip()
        if not texto_limpio: return df_original.copy() if df_original is not None else pd.DataFrame(), ""
        if df_original is None or df_original.empty: return pd.DataFrame(columns=df_original.columns if df_original is not None else []), texto_limpio
        negados, positivos_parts, last_end = [], [], 0
        for m in self.patron_termino_negado.finditer(texto_limpio):
            positivos_parts.append(texto_limpio[last_end:m.start()]); last_end = m.end()
            term_raw = m.group(1) or m.group(2)
            if term_raw: 
                norm = self._normalizar_para_busqueda(term_raw.strip('"')) # Quitar comillas de frases negadas
                if norm and norm not in negados: negados.append(norm) 
        positivos_parts.append(texto_limpio[last_end:]); positivos_str = ' '.join("".join(positivos_parts).split())
        df_filt = df_original.copy()
        if not negados: return df_filt, positivos_str
        logger.debug(f"Aplicando negación con términos: {negados}")
        mascara_excluir_total = pd.Series(False, index=df_filt.index)
        for term_neg in negados:
            if not term_neg: continue
            mascara_term_actual = pd.Series(False, index=df_filt.index)
            for col_nombre in cols:
                if col_nombre not in df_filt.columns: continue
                try:
                    serie_norm_df = df_filt[col_nombre].astype(str).map(self._normalizar_para_busqueda)
                    pat_regex = r"\b" + re.escape(term_neg) + r"\b"
                    mascara_term_actual |= serie_norm_df.str.contains(pat_regex, regex=True, na=False)
                except Exception as e_neg: logger.error(f"Error negación col '{col_nombre}', término '{term_neg}': {e_neg}")
            mascara_excluir_total |= mascara_term_actual
        df_resultado_final = df_filt[~mascara_excluir_total]
        logger.info(f"Negación: {len(df_filt)} -> {len(df_resultado_final)} filas. Negados: {negados}. Positivos: '{positivos_str}'")
        return df_resultado_final, positivos_str

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        texto_limpio = texto_complejo.strip(); 
        if not texto_limpio: return "OR", [] 
        for sep_regex, sep_char in [(r"\s*\|\s*", "|"), (r"\s*/\s*", "/")]:
            if sep_char in texto_limpio : 
                 segmentos = [s.strip() for s in re.split(sep_regex, texto_limpio) if s.strip()]
                 if len(segmentos) > 1 or (len(segmentos) == 1 and texto_limpio != segmentos[0] and texto_limpio.count(sep_char) > 0) :
                    return "OR", segmentos
        return "AND", [texto_limpio]

    def _descomponer_nivel2_and(self, termino_segmento_n1: str) -> Tuple[str, List[str]]:
        termino_limpio = termino_segmento_n1.strip(); 
        if not termino_limpio: return "AND", []
        partes = re.split(r'\s+\+\s+', termino_limpio)
        return "AND", [p.strip() for p in partes if p.strip()]

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        analizados = []
        for term_orig_bruto in terminos_brutos:
            term_orig_procesado = str(term_orig_bruto).strip()
            # CORRECCIÓN: Quitar comillas si el término original estaba entrecomillado
            # (esto sucede con los términos de la query OR construida).
            if len(term_orig_procesado) > 1 and term_orig_procesado.startswith('"') and term_orig_procesado.endswith('"'):
                term_orig_final = term_orig_procesado[1:-1]
            else:
                term_orig_final = term_orig_procesado
            
            if not term_orig_final: continue

            item_analizado: Dict[str, Any] = {"original": term_orig_final} # Usar el término sin comillas para 'original' también
            
            mc, mr = self.patron_comparacion.match(term_orig_final), self.patron_rango.match(term_orig_final)
            if mc:
                op, v_str, unidad_str = mc.groups(); v_num = self._parse_numero(v_str)
                if v_num is not None:
                    op_map = {">": "gt", "<": "lt", ">=": "ge", "<=": "le", "=": "eq"}
                    u_canon = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str.strip()) if unidad_str and unidad_str.strip() else None
                    item_analizado.update({"tipo": op_map.get(op), "valor": v_num, "unidad_busqueda": u_canon})
                else: item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(term_orig_final)})
            elif mr:
                v1_s, v2_s, u_s = mr.groups(); v1,v2 = self._parse_numero(v1_s), self._parse_numero(v2_s)
                if v1 is not None and v2 is not None:
                    u_canon = self.extractor_magnitud.obtener_magnitud_normalizada(u_s.strip()) if u_s and u_s.strip() else None
                    item_analizado.update({"tipo": "range", "valor": sorted([v1, v2]), "unidad_busqueda": u_canon}) 
                else: item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(term_orig_final)})
            else: 
                item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(term_orig_final)})
            analizados.append(item_analizado)
        logger.debug(f"Términos analizados (post-AND, comillas quitadas): {analizados}")
        return analizados

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        if isinstance(num_str, (int, float)): return float(num_str)
        if not isinstance(num_str, str): return None
        try: return float(num_str.replace(",", ".")) 
        except ValueError: return None 

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols: List[str], term_an: Dict[str, Any]) -> pd.Series:
        tipo, valor, unidad_req_canon = term_an["tipo"], term_an["valor"], term_an.get("unidad_busqueda")
        mascara_total = pd.Series(False, index=df.index)
        logger.debug(f"Generando máscara para término: tipo='{tipo}', valor='{valor}', unidad_req='{unidad_req_canon}'")
        for col_n in cols:
            if col_n not in df.columns: continue
            col_s = df[col_n]; mascara_col_actual_num = pd.Series(False, index=df.index)
            if tipo in ["gt", "lt", "ge", "le", "range", "eq"]:
                for idx, val_celda_raw in col_s.items(): 
                    if pd.isna(val_celda_raw) or str(val_celda_raw).strip() == "": continue
                    for match_c in self.patron_num_unidad_df.finditer(str(val_celda_raw)):
                        try:
                            num_c_val = self._parse_numero(match_c.group(1)); u_c_raw = match_c.group(2)
                            if num_c_val is None: continue
                            u_c_canon = self.extractor_magnitud.obtener_magnitud_normalizada(u_c_raw.strip()) if u_c_raw and u_c_raw.strip() else None
                            u_ok = (unidad_req_canon is None) or (u_c_canon is not None and u_c_canon == unidad_req_canon) or \
                                   (u_c_raw and unidad_req_canon and self.extractor_magnitud._normalizar_texto(u_c_raw.strip()) == unidad_req_canon)
                            if not u_ok: continue
                            cond = False
                            if tipo == "eq" and np.isclose(num_c_val, valor): cond = True
                            elif tipo == "gt" and num_c_val > valor and not np.isclose(num_c_val, valor): cond = True
                            elif tipo == "lt" and num_c_val < valor and not np.isclose(num_c_val, valor): cond = True
                            elif tipo == "ge" and (num_c_val >= valor or np.isclose(num_c_val, valor)): cond = True
                            elif tipo == "le" and (num_c_val <= valor or np.isclose(num_c_val, valor)): cond = True
                            elif tipo == "range" and ((valor[0] <= num_c_val or np.isclose(num_c_val, valor[0])) and (num_c_val <= valor[1] or np.isclose(num_c_val, valor[1]))): cond = True
                            if cond: mascara_col_actual_num.at[idx] = True; break
                        except ValueError: continue
                mascara_total |= mascara_col_actual_num
            elif tipo == "str":
                try:
                    val_norm_busq = str(valor); 
                    if not val_norm_busq: continue
                    serie_norm_df_col = col_s.astype(str).map(self._normalizar_para_busqueda)
                    pat_regex = r"\b" + re.escape(val_norm_busq) + r"\b"
                    logger.debug(f"BUSQ_STR: Term='{val_norm_busq}' (Regex='{pat_regex}') en Col='{col_n}'")
                    mascara_col_actual = serie_norm_df_col.str.contains(pat_regex, regex=True, na=False)
                    if mascara_col_actual.any():
                         logger.debug(f"HIT_STR: Term='{val_norm_busq}' en Col='{col_n}'. Índices: {df.index[mascara_col_actual].tolist()}")
                    mascara_total |= mascara_col_actual
                except Exception as e: logger.warning(f"Error búsqueda STR col '{col_n}' para '{valor}': {e}")
        return mascara_total

    def _aplicar_mascara_combinada_para_segmento_and(self, df: pd.DataFrame, cols: List[str], term_an_seg: List[Dict[str, Any]]) -> pd.Series:
        if df is None or df.empty or not cols: return pd.Series(False, index=df.index if df is not None else None) 
        if not term_an_seg: return pd.Series(False, index=df.index) 
        mascara_final = pd.Series(True, index=df.index) 
        for term_ind_an in term_an_seg:
            mascara_este_term = self._generar_mascara_para_un_termino(df, cols, term_ind_an)
            mascara_final &= mascara_este_term 
            if not mascara_final.any(): break
        return mascara_final

    def _combinar_mascaras_de_segmentos_or(self, lista_mascaras: List[pd.Series], df_idx_ref: Optional[pd.Index] = None) -> pd.Series:
        if not lista_mascaras: return pd.Series(False, index=df_idx_ref) if df_idx_ref is not None else pd.Series(dtype=bool) 
        idx_usar = df_idx_ref if df_idx_ref is not None else lista_mascaras[0].index
        mascara_final = pd.Series(False, index=idx_usar) 
        for masc_seg in lista_mascaras:
            if not masc_seg.index.equals(idx_usar) and not masc_seg.empty:
                try: masc_seg = masc_seg.reindex(idx_usar, fill_value=False)
                except Exception as e_reidx: logger.error(f"Fallo reindex máscara OR: {e_reidx}"); continue 
            mascara_final |= masc_seg 
        return mascara_final

    def _procesar_busqueda_en_df_objetivo(self, df_obj: pd.DataFrame, cols_obj: List[str], term_busq_orig: str) -> Tuple[pd.DataFrame, Optional[str]]:
        logger.debug(f"Procesando búsqueda: '{term_busq_orig}' en {len(cols_obj)} cols del DF de {len(df_obj)} filas.")
        df_post_neg, term_pos_q = self._aplicar_negaciones_y_extraer_positivos(df_obj, cols_obj, term_busq_orig)
        if df_post_neg.empty and not term_pos_q.strip(): return df_post_neg.copy(), None 
        if not term_pos_q.strip(): return df_post_neg.copy(), None
        op_n1, segs_n1 = self._descomponer_nivel1_or(term_pos_q)
        if not segs_n1: return (pd.DataFrame(columns=df_post_neg.columns), "Término positivo inválido post-OR.") if term_busq_orig.strip() or term_pos_q.strip() else (df_post_neg.copy(), None)
        lista_mascaras_or = []
        for seg_n1 in segs_n1: 
            op_n2, terms_brutos_n2 = self._descomponer_nivel2_and(seg_n1) 
            terms_atom_an = self._analizar_terminos(terms_brutos_n2) 
            mascara_seg_n1 = self._aplicar_mascara_combinada_para_segmento_and(df_post_neg, cols_obj, terms_atom_an) if terms_atom_an else pd.Series(False, index=df_post_neg.index)
            lista_mascaras_or.append(mascara_seg_n1)
        if not lista_mascaras_or: return pd.DataFrame(columns=df_post_neg.columns), "Error interno: no se generaron máscaras OR."
        mascara_final_df_objetivo = self._combinar_mascaras_de_segmentos_or(lista_mascaras_or, df_post_neg.index)
        df_resultado = df_post_neg[mascara_final_df_objetivo].copy()
        logger.debug(f"Resultado de _procesar_busqueda_en_df_objetivo: {len(df_resultado)} filas.")
        return df_resultado, None

    def _extraer_terminos_de_fila_completa(self, fila_df: pd.Series) -> Set[str]:
        terminos_extraidos: Set[str] = set()
        if fila_df is None or fila_df.empty: 
            logger.debug(f"Fila vacía o None pasada a _extraer_terminos_de_fila_completa.")
            return terminos_extraidos
        for valor_celda in fila_df.values: 
            if pd.notna(valor_celda): 
                texto_celda_str = str(valor_celda).strip()
                if texto_celda_str: 
                    texto_celda_norm = self._normalizar_para_busqueda(texto_celda_str)
                    palabras_celda = [palabra for palabra in texto_celda_norm.split() if len(palabra) > 1 and not palabra.isdigit()]
                    if palabras_celda: terminos_extraidos.update(palabras_celda)
                    elif texto_celda_norm and len(texto_celda_norm) > 1 and not texto_celda_norm.isdigit(): terminos_extraidos.add(texto_celda_norm)
        if terminos_extraidos: logger.debug(f"Términos extraídos de la fila ({fila_df.name if hasattr(fila_df, 'name') else 'N/A'}): {sorted(list(terminos_extraidos))[:10]}...")
        return terminos_extraidos

    def buscar(self, termino_busqueda_original: str, buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[List[int]], Optional[str]]:
        logger.info(f"Motor.buscar INICIO: termino='{termino_busqueda_original}', via_dicc={buscar_via_diccionario_flag}")
        df_vacio_desc = pd.DataFrame(columns=(self.datos_descripcion.columns if self.datos_descripcion is not None else []))
        fcds_obtenidos: Optional[pd.DataFrame] = None; indices_fcds_a_resaltar: Optional[List[int]] = None

        if not termino_busqueda_original.strip():
            return (self.datos_descripcion.copy() if self.datos_descripcion is not None else df_vacio_desc), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, (None if self.datos_descripcion is not None else "Descripciones no cargadas.")

        if buscar_via_diccionario_flag:
            if self.datos_diccionario is None: return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, None, "Diccionario no cargado."
            
            cols_dic_fcds, err_cols_dic = self._obtener_nombres_columnas_busqueda_df(self.datos_diccionario, [], "diccionario_fcds_inicial")
            if not cols_dic_fcds: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, None, err_cols_dic or "No se pudo det. cols para FCDs."
            
            logger.info(f"BUSCAR EN DICC (FCDs): Query='{termino_busqueda_original}' en {len(cols_dic_fcds)} cols: {cols_dic_fcds[:5]}...")
            try:
                fcds_obtenidos, error_procesamiento_dic = self._procesar_busqueda_en_df_objetivo(self.datos_diccionario, cols_dic_fcds, termino_busqueda_original)
                if error_procesamiento_dic: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_procesamiento_dic
                
                if fcds_obtenidos is not None and not fcds_obtenidos.empty:
                    indices_fcds_a_resaltar = fcds_obtenidos.index.tolist()
                    logger.info(f"FCDs obtenidos del diccionario: {len(fcds_obtenidos)} filas. Índices para resaltar: {indices_fcds_a_resaltar}")
                    if logger.isEnabledFor(logging.DEBUG): logger.debug(f"DataFrame FCDs (muestra):\n{fcds_obtenidos.head(min(3, len(fcds_obtenidos))).to_string()}")
                else:
                    logger.info(f"No se encontraron FCDs en el diccionario para '{termino_busqueda_original}'.")
                    return df_vacio_desc, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None
            except Exception as e_dic: logger.exception("Excepción en búsqueda en diccionario."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc): {e_dic}"

            if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos, indices_fcds_a_resaltar, "Descripciones no cargadas."
            
            terminos_para_descripcion: Set[str] = set()
            logger.info(f"Extrayendo TODOS los términos de las {len(fcds_obtenidos)} fila(s) FCD para buscar en descripciones.")
            for idx_fcd_loop, fila_fcd in fcds_obtenidos.iterrows():
                terminos_de_esta_fila = self._extraer_terminos_de_fila_completa(fila_fcd)
                terminos_para_descripcion.update(terminos_de_esta_fila)
            
            if not terminos_para_descripcion:
                logger.info("FCDs encontrados, pero _extraer_terminos_de_fila_completa no produjo términos para buscar en descripciones.")
                return df_vacio_desc, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, fcds_obtenidos, indices_fcds_a_resaltar, None
            
            logger.info(f"TÉRMINOS FINALES PARA DESC ({len(terminos_para_descripcion)} únicos, muestra): {sorted(list(terminos_para_descripcion))[:15]}...")
            try:
                # La query OR se construye con los términos YA LIMPIOS (sin comillas extras).
                # _analizar_terminos se encargará de procesar cada uno como un string individual.
                termino_or_para_desc = " | ".join(terminos_para_descripcion) # No añadir comillas aquí, _analizar_terminos los tratará como literales.
                                                                          # Si un término tiene espacios, _analizar_terminos lo normalizará.
                                                                          # Para que un término con espacios se busque como frase, la sintaxis de query lo debe indicar (ej. "palabra A" + "palabra B")
                                                                          # o la UI debe permitir encapsular frases. Aquí, son términos individuales ORed.
                
                if not termino_or_para_desc: return df_vacio_desc, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, fcds_obtenidos, indices_fcds_a_resaltar, "Query OR para descripciones vacía."
                
                cols_desc_para_fcds, err_cols_desc_fcd = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion")
                if not cols_desc_para_fcds: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos, indices_fcds_a_resaltar, err_cols_desc_fcd
                
                logger.info(f"BUSCAR EN DESC (vía FCD): Query OR (muestra): '{termino_or_para_desc[:200]}...'")
                resultados_desc_via_dic, error_proc_desc_fcd = self._procesar_busqueda_en_df_objetivo(self.datos_descripcion, cols_desc_para_fcds, termino_or_para_desc)
                if error_proc_desc_fcd: return df_vacio_desc, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos, indices_fcds_a_resaltar, error_proc_desc_fcd
                
                if resultados_desc_via_dic is None or resultados_desc_via_dic.empty:
                    return df_vacio_desc, OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, fcds_obtenidos, indices_fcds_a_resaltar, None
                else:
                    return resultados_desc_via_dic, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos, indices_fcds_a_resaltar, None
            except Exception as e_desc_fcd: logger.exception("Excepción búsqueda en descripciones vía FCDs."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos, indices_fcds_a_resaltar, f"Error motor (desc vía FCD): {e_desc_fcd}"
        else: 
            if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas."
            cols_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion")
            if not cols_desc_directo: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_directo
            try:
                logger.info(f"BUSCAR EN DESC (DIRECTO): Query '{termino_busqueda_original}' en descripciones.")
                resultados_directos, error_proc_desc_directo = self._procesar_busqueda_en_df_objetivo(self.datos_descripcion, cols_desc_directo, termino_busqueda_original)
                if error_proc_desc_directo: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_proc_desc_directo
                if resultados_directos is None or resultados_directos.empty: return df_vacio_desc, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
                else: return resultados_directos, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None, None
            except Exception as e_desc_dir: logger.exception("Excepción búsqueda directa en descripciones."); return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (desc directa): {e_desc_dir}"

# --- Interfaz Gráfica ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE = "config_buscador_definitivo_v1.3.json" 
    def __init__(self):
        super().__init__()
        self.title("Buscador Definitivo v1.3") 
        self.geometry("1250x800")
        self.config = self._cargar_configuracion()
        indices_cfg_preview = self.config.get("indices_columnas_busqueda_dic_preview", []) 
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg_preview)
        self.resultados_actuales: Optional[pd.DataFrame] = None
        self.texto_busqueda_var = tk.StringVar(self); self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        self.ultimo_termino_buscado: Optional[str] = None; self.reglas_guardadas: List[Dict[str, Any]] = []
        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.indices_fcds_resaltados: Optional[List[int]] = None 
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO
        self.color_fila_par = "white"; self.color_fila_impar = "#f0f0f0"; self.color_resaltado_dic = "light sky blue" 
        self.op_buttons: Dict[str, ttk.Button] = {}
        self._configurar_estilo_ttk(); self._crear_widgets(); self._configurar_grid(); self._configurar_eventos()
        self._configurar_tags_treeview(); self._configurar_orden_tabla(self.tabla_resultados); self._configurar_orden_tabla(self.tabla_diccionario)
        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones."); self._deshabilitar_botones_operadores(); self._actualizar_botones_estado_general()
        logger.info(f"Interfaz Gráfica (Definitiva v1.3) inicializada.")

    def _on_texto_busqueda_change(self,v,i,m): self._actualizar_estado_botones_operadores()
    def _cargar_configuracion(self) -> Dict: 
        config = {}; cfg_path = Path(self.CONFIG_FILE)
        if cfg_path.exists():
            try:
                with cfg_path.open("r", encoding="utf-8") as f: config = json.load(f)
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE}")
            except Exception as e: logger.error(f"Error al cargar config: {e}")
        else: logger.info(f"Archivo config '{self.CONFIG_FILE}' no encontrado.")
        for key in ["last_dic_path", "last_desc_path"]: config[key] = str(Path(config[key])) if config.get(key) else None
        config.setdefault("indices_columnas_busqueda_dic_preview", []) 
        return config
    def _guardar_configuracion(self): 
        self.config["last_dic_path"] = str(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else None
        self.config["last_desc_path"] = str(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else None
        self.config["indices_columnas_busqueda_dic_preview"] = self.motor.indices_columnas_busqueda_dic_preview
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(self.config, f, indent=4)
            logger.info(f"Configuración guardada en: {self.CONFIG_FILE}")
        except Exception as e: logger.error(f"Error al guardar config: {e}")
    def _configurar_estilo_ttk(self): 
        s = ttk.Style(self); os_n = platform.system(); prefs = {"Windows":["vista","xpnative"],"Darwin":["aqua"],"Linux":["clam","alt"]}
        theme = next((t for t in prefs.get(os_n,["clam"]) if t in s.theme_names()), s.theme_use() or "default")
        try: s.theme_use(theme); s.configure("Operator.TButton",padding=(2,1),font=("TkDefaultFont",9)); logger.info(f"Tema TTK: {theme}")
        except: logger.warning(f"Fallo al aplicar tema {theme}")
    def _crear_widgets(self):
        self.marco_controles=ttk.LabelFrame(self,text="Controles")
        self.btn_cargar_diccionario=ttk.Button(self.marco_controles,text="Cargar Diccionario",command=self._cargar_diccionario)
        self.lbl_dic_cargado=ttk.Label(self.marco_controles,text="Dic: Ninguno",width=20,anchor=tk.W,relief=tk.SUNKEN,borderwidth=1)
        self.btn_cargar_descripciones=ttk.Button(self.marco_controles,text="Cargar Descripciones",command=self._cargar_excel_descripcion)
        self.lbl_desc_cargado=ttk.Label(self.marco_controles,text="Desc: Ninguno",width=20,anchor=tk.W,relief=tk.SUNKEN,borderwidth=1)
        self.frame_ops=ttk.Frame(self.marco_controles)
        op_buttons_defs = [("+","+"),("|","|"),("#","#"),("> ",">"),("< ","<"),("≥ ",">="),("≤ ","<="),("-","-")]
        for i, (text, op_val_with_space) in enumerate(op_buttons_defs):
            op_val_clean = op_val_with_space.strip() 
            btn = ttk.Button(self.frame_ops,text=text,command=lambda op=op_val_clean: self._insertar_operador_validado(op),style="Operator.TButton",width=3)
            btn.grid(row=0,column=i,padx=1,pady=1,sticky="nsew")
            self.op_buttons[op_val_clean] = btn 
        self.entrada_busqueda=ttk.Entry(self.marco_controles,width=60,textvariable=self.texto_busqueda_var)
        self.btn_buscar=ttk.Button(self.marco_controles,text="Buscar",command=self._ejecutar_busqueda)
        self.btn_salvar_regla=ttk.Button(self.marco_controles,text="Salvar Regla",command=self._salvar_regla_actual,state="disabled")
        self.btn_ayuda=ttk.Button(self.marco_controles,text="?",command=self._mostrar_ayuda,width=3)
        self.btn_exportar=ttk.Button(self.marco_controles,text="Exportar",command=self._exportar_resultados,state="disabled")
        self.lbl_tabla_diccionario=ttk.Label(self,text="Vista Previa Diccionario:")
        self.frame_tabla_diccionario=ttk.Frame(self);self.tabla_diccionario=ttk.Treeview(self.frame_tabla_diccionario,show="headings",height=8);self.scrolly_diccionario=ttk.Scrollbar(self.frame_tabla_diccionario,orient="vertical",command=self.tabla_diccionario.yview);self.scrollx_diccionario=ttk.Scrollbar(self.frame_tabla_diccionario,orient="horizontal",command=self.tabla_diccionario.xview);self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set,xscrollcommand=self.scrollx_diccionario.set)
        self.lbl_tabla_resultados=ttk.Label(self,text="Resultados / Descripciones:");self.frame_tabla_resultados=ttk.Frame(self);self.tabla_resultados=ttk.Treeview(self.frame_tabla_resultados,show="headings");self.scrolly_resultados=ttk.Scrollbar(self.frame_tabla_resultados,orient="vertical",command=self.tabla_resultados.yview);self.scrollx_resultados=ttk.Scrollbar(self.frame_tabla_resultados,orient="horizontal",command=self.tabla_resultados.xview);self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set,xscrollcommand=self.scrollx_resultados.set)
        self.barra_estado=ttk.Label(self,text="Listo.",relief=tk.SUNKEN,anchor=tk.W,borderwidth=1);self._actualizar_etiquetas_archivos()
    def _configurar_grid(self): 
        self.grid_rowconfigure(2,weight=1);self.grid_rowconfigure(4,weight=3);self.grid_columnconfigure(0,weight=1);self.marco_controles.grid(row=0,column=0,sticky="new",padx=10,pady=(10,5));self.marco_controles.grid_columnconfigure(1,weight=1);self.marco_controles.grid_columnconfigure(3,weight=1);self.btn_cargar_diccionario.grid(row=0,column=0,padx=(5,0),pady=5,sticky="w");self.lbl_dic_cargado.grid(row=0,column=1,padx=(2,10),pady=5,sticky="ew");self.btn_cargar_descripciones.grid(row=0,column=2,padx=(5,0),pady=5,sticky="w");self.lbl_desc_cargado.grid(row=0,column=3,padx=(2,5),pady=5,sticky="ew");self.frame_ops.grid(row=1,column=0,columnspan=6,padx=5,pady=(5,0),sticky="ew");[self.frame_ops.grid_columnconfigure(i,weight=1) for i in range(len(self.op_buttons))];self.entrada_busqueda.grid(row=2,column=0,columnspan=2,padx=5,pady=(0,5),sticky="ew");self.btn_buscar.grid(row=2,column=2,padx=(2,0),pady=(0,5),sticky="w");self.btn_salvar_regla.grid(row=2,column=3,padx=(2,0),pady=(0,5),sticky="w");self.btn_ayuda.grid(row=2,column=4,padx=(2,0),pady=(0,5),sticky="w");self.btn_exportar.grid(row=2,column=5,padx=(10,5),pady=(0,5),sticky="e");self.lbl_tabla_diccionario.grid(row=1,column=0,sticky="sw",padx=10,pady=(10,0));self.frame_tabla_diccionario.grid(row=2,column=0,sticky="nsew",padx=10,pady=(0,10));self.frame_tabla_diccionario.grid_rowconfigure(0,weight=1);self.frame_tabla_diccionario.grid_columnconfigure(0,weight=1);self.tabla_diccionario.grid(row=0,column=0,sticky="nsew");self.scrolly_diccionario.grid(row=0,column=1,sticky="ns");self.scrollx_diccionario.grid(row=1,column=0,sticky="ew");self.lbl_tabla_resultados.grid(row=3,column=0,sticky="sw",padx=10,pady=(0,0));self.frame_tabla_resultados.grid(row=4,column=0,sticky="nsew",padx=10,pady=(0,10));self.frame_tabla_resultados.grid_rowconfigure(0,weight=1);self.frame_tabla_resultados.grid_columnconfigure(0,weight=1);self.tabla_resultados.grid(row=0,column=0,sticky="nsew");self.scrolly_resultados.grid(row=0,column=1,sticky="ns");self.scrollx_resultados.grid(row=1,column=0,sticky="ew");self.barra_estado.grid(row=5,column=0,sticky="sew",padx=0,pady=(5,0))
    def _configurar_eventos(self): self.entrada_busqueda.bind("<Return>",lambda e:self._ejecutar_busqueda());self.protocol("WM_DELETE_WINDOW",self.on_closing)
    def _actualizar_estado(self,m): self.barra_estado.config(text=m);logger.info(f"UI: {m}");self.update_idletasks()
    def _mostrar_ayuda(self): 
        ayuda = ("Sintaxis:\n- Texto: `router cisco`\n- AND: `tarjeta + 16 puertos`\n- OR: `modulo | SFP` o `modulo / SFP`\n"
                 "- Numérico: `>1000W`, `<50V`, `>=48A`, `<=10.5W`\n- Rango: `10-20V`\n- Negación: `#palabra` o `# \"frase completa\"`\n\n"
                 "Flujo de Búsqueda:\n1. Término se busca en todas las columnas del Diccionario.\n"
                 "2. Filas FCD del Diccionario se resaltan.\n"
                 "3. TODAS las palabras de esas filas FCD se extraen.\n"
                 "4. Esas palabras (unidas por OR) se buscan en Descripciones.\n"
                 "5. Si no hay FCDs, se ofrece búsqueda directa del término original en Descripciones.\n"
                 "6. Búsqueda vacía muestra todas las Descripciones.")
        messagebox.showinfo("Ayuda - Sintaxis y Flujo", ayuda)
    def _configurar_tags_treeview(self): 
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure("par", background=self.color_fila_par)
            tabla.tag_configure("impar", background=self.color_fila_impar)
        self.tabla_diccionario.tag_configure("resaltado_azul", background=self.color_resaltado_dic, foreground="black")
    def _configurar_orden_tabla(self,tabla): 
        cols = tabla["columns"]
        if cols: [tabla.heading(c,text=str(c),anchor=tk.W,command=lambda col=c,tbl=tabla:self._ordenar_columna(tbl,col,False)) for c in cols]
    def _ordenar_columna(self,tabla,col,rev): 
        df_copia=None;idx_resaltar=None
        if tabla==self.tabla_diccionario and self.motor.datos_diccionario is not None:df_copia=self.motor.datos_diccionario.copy();idx_resaltar=self.indices_fcds_resaltados
        elif tabla==self.tabla_resultados and self.resultados_actuales is not None:df_copia=self.resultados_actuales.copy()
        else: tabla.heading(col,command=lambda c=col,t=tabla:self._ordenar_columna(t,c,not rev));return
        if df_copia.empty or col not in df_copia.columns: tabla.heading(col,command=lambda c=col,t=tabla:self._ordenar_columna(t,c,not rev));return
        try:
            df_num=pd.to_numeric(df_copia[col],errors='coerce')
            df_ord=df_copia.sort_values(by=col,ascending=not rev,na_position='last',key=(lambda x:pd.to_numeric(x,errors='coerce')) if not df_num.isna().all() else (lambda x:x.astype(str).str.lower()))
            if tabla==self.tabla_diccionario:self._actualizar_tabla(tabla,df_ord,limite_filas=None,indices_a_resaltar=idx_resaltar)
            elif tabla==self.tabla_resultados:self.resultados_actuales=df_ord;self._actualizar_tabla(tabla,self.resultados_actuales)
            tabla.heading(col,command=lambda c=col,t=tabla:self._ordenar_columna(t,c,not rev));self._actualizar_estado(f"Ordenado por '{col}'.")
        except Exception as e:logging.exception(f"Error ordenando '{col}'");messagebox.showerror("Error Ordenar",f"Fallo: {e}");tabla.heading(col,command=lambda c=col,t=tabla:self._ordenar_columna(t,c,False))
    def _actualizar_tabla(self,tabla,datos,limite_filas=None,columnas_a_mostrar=None,indices_a_resaltar=None): 
        is_dicc=tabla==self.tabla_diccionario;[tabla.delete(i) for i in tabla.get_children()];tabla["columns"]=()
        if datos is None or datos.empty:self._configurar_orden_tabla(tabla);return
        cols_orig=list(datos.columns);cols_usar=cols_orig
        if columnas_a_mostrar:cols_usar=[c for c in columnas_a_mostrar if c in cols_orig] or cols_orig
        if not cols_usar:self._configurar_orden_tabla(tabla);return
        tabla["columns"]=tuple(cols_usar)
        for c in cols_usar:tabla.heading(c,text=str(c),anchor=tk.W);tabla.column(c,anchor=tk.W,width=max(70,min(int(max(len(str(c))*8,(datos[c].astype(str).str.len().max() if not datos[c].empty else 0)*6.5)+25),400)),minwidth=70)
        df_iterar=datos[cols_usar];num_filas_original=len(df_iterar)
        if not (is_dicc and indices_a_resaltar and num_filas_original > 0) and limite_filas and num_filas_original > limite_filas: df_iterar=df_iterar.head(limite_filas); logger.debug(f"Mostrando {limite_filas} de {num_filas_original} en tabla.")
        elif is_dicc and indices_a_resaltar: logger.debug(f"Mostrando todas las {num_filas_original} filas del diccionario para resaltar.")
        for i,(idx,row) in enumerate(df_iterar.iterrows()):
            vals=[str(v) if pd.notna(v) else "" for v in row.values];tags=["par" if i%2==0 else "impar"]
            if is_dicc and indices_a_resaltar and idx in indices_a_resaltar:tags.append("resaltado_azul")
            try:tabla.insert("","end",values=vals,tags=tuple(tags),iid=f"row_{idx}")
            except Exception as e_ins: logger.warning(f"Error insertando fila {idx} en treeview: {e_ins}") 
        self._configurar_orden_tabla(tabla)
    def _actualizar_etiquetas_archivos(self): 
        max_l=25;dic_p=self.motor.archivo_diccionario_actual;desc_p=self.motor.archivo_descripcion_actual
        dic_n=dic_p.name if dic_p else "Ninguno";desc_n=desc_p.name if desc_p else "Ninguno"
        dic_d=f"Dic: {dic_n}" if len(dic_n)<=max_l else f"Dic: ...{dic_n[-(max_l-4):]}";desc_d=f"Desc: {desc_n}" if len(desc_n)<=max_l else f"Desc: ...{desc_n[-(max_l-4):]}"
        self.lbl_dic_cargado.config(text=dic_d,foreground="green" if dic_p else "red");self.lbl_desc_cargado.config(text=desc_d,foreground="green" if desc_p else "red")
    def _actualizar_botones_estado_general(self): 
        dic_ok=self.motor.datos_diccionario is not None;desc_ok=self.motor.datos_descripcion is not None
        self._actualizar_estado_botones_operadores() if dic_ok or desc_ok else self._deshabilitar_botones_operadores()
        self.btn_buscar["state"]="normal" if dic_ok and desc_ok else "disabled";salvar_ok=False
        if self.ultimo_termino_buscado and self.origen_principal_resultados!=OrigenResultados.NINGUNO:
            if self.origen_principal_resultados.es_via_diccionario and ((self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty)or(self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and self.origen_principal_resultados==OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC)):salvar_ok=True
            elif (self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados==OrigenResultados.DIRECTO_DESCRIPCION_VACIA) and self.desc_finales_de_ultima_busqueda is not None:salvar_ok=True
        self.btn_salvar_regla["state"]="normal" if salvar_ok else "disabled";self.btn_exportar["state"]="normal" if (self.resultados_actuales is not None and not self.resultados_actuales.empty) else "disabled"
    def _cargar_diccionario(self): 
        cfg_path=self.config.get("last_dic_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada=filedialog.askopenfilename(title="Cargar Diccionario",filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada: return
        nombre_archivo = Path(ruta_seleccionada).name # CORRECCIÓN: Definir antes del if/else
        self._actualizar_estado(f"Cargando dicc: {nombre_archivo}...")
        self._actualizar_tabla(self.tabla_diccionario,None);self._actualizar_tabla(self.tabla_resultados,None);self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        ok,msg=self.motor.cargar_excel_diccionario(ruta_seleccionada)
        desc_n_title=Path(self.motor.archivo_descripcion_actual).name if self.motor.archivo_descripcion_actual else "N/A" # CORRECCIÓN: Definir antes
        if ok and self.motor.datos_diccionario is not None:
            self.config["last_dic_path"]=ruta_seleccionada;self._guardar_configuracion();df_d=self.motor.datos_diccionario;n_filas=len(df_d)
            cols_prev,_=self.motor._obtener_nombres_columnas_busqueda_df(df_d,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
            self.lbl_tabla_diccionario.config(text=f"Diccionario ({n_filas} filas)");self._actualizar_tabla(self.tabla_diccionario,df_d,limite_filas=100,columnas_a_mostrar=cols_prev)
            self.title(f"Buscador - Dic: {nombre_archivo} | Desc: {desc_n_title}");self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({n_filas}) cargado.")
        else:
            self._actualizar_estado(f"Error cargando diccionario: {msg or 'Desconocido'}");messagebox.showerror("Error Carga Dicc",msg or "Error desconocido")
            self.title(f"Buscador - Dic: N/A (Error) | Desc: {desc_n_title}") 
        self._actualizar_etiquetas_archivos();self._actualizar_botones_estado_general()
        
    def _cargar_excel_descripcion(self): 
        cfg_path=self.config.get("last_desc_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada_str=filedialog.askopenfilename(title="Cargar Descripciones",filetypes=[("Archivos Excel","*.xlsx *.xls"),("Todos los archivos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada_str: logger.info("Carga de descripciones cancelada."); return
        nombre_archivo = Path(ruta_seleccionada_str).name; 
        self._actualizar_estado(f"Cargando descripciones: {nombre_archivo}...")
        self.resultados_actuales=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self._actualizar_tabla(self.tabla_resultados,None) 
        ok, msg_error = self.motor.cargar_excel_descripcion(ruta_seleccionada_str)
        dic_n_title=Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A" 
        if ok and self.motor.datos_descripcion is not None:
            self.config["last_desc_path"] = ruta_seleccionada_str; self._guardar_configuracion()
            df_desc = self.motor.datos_descripcion; num_filas = len(df_desc)
            self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando vista previa...")
            self._actualizar_tabla(self.tabla_resultados, df_desc, limite_filas=200)
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: {nombre_archivo}")
        else:
            error_a_mostrar = msg_error or "Ocurrió un error desconocido al cargar el archivo de descripciones."
            self._actualizar_estado(f"Error cargando descripciones: {error_a_mostrar}"); messagebox.showerror("Error al Cargar Archivo de Descripciones", error_a_mostrar)
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: N/A (Error)")
        self._actualizar_etiquetas_archivos();self._actualizar_botones_estado_general()

    def _ejecutar_busqueda(self): 
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:messagebox.showwarning("Archivos Faltantes","Cargue Diccionario y Descripciones.");return
        term_ui=self.texto_busqueda_var.get();self.ultimo_termino_buscado=term_ui
        self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        self._actualizar_tabla(self.tabla_resultados,None);self._actualizar_estado(f"Buscando '{term_ui}'...")
        
        res_df,origen,fcds,idx_res,err_msg = self.motor.buscar(
            termino_busqueda_original=term_ui, # CORREGIDO
            buscar_via_diccionario_flag=True
        )
        
        self.fcds_de_ultima_busqueda=fcds;self.origen_principal_resultados=origen;self.indices_fcds_resaltados=idx_res
        df_desc_cols=self.motor.datos_descripcion.columns if self.motor.datos_descripcion is not None else []
        if self.motor.datos_diccionario is not None: 
            num_fcds_actual=len(self.indices_fcds_resaltados) if self.indices_fcds_resaltados else 0
            dicc_lbl=f"Diccionario ({len(self.motor.datos_diccionario)} filas)" + (f" - {num_fcds_actual} FCDs resaltados" if num_fcds_actual>0 and origen.es_via_diccionario and origen!=OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS else "")
            self.lbl_tabla_diccionario.config(text=dicc_lbl)
            cols_prev_dic,_=self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
            self._actualizar_tabla(self.tabla_diccionario,self.motor.datos_diccionario,limite_filas=None if self.indices_fcds_resaltados else 100,columnas_a_mostrar=cols_prev_dic,indices_a_resaltar=self.indices_fcds_resaltados)
        if err_msg and origen.es_error_operacional:messagebox.showerror("Error Motor",f"Error interno: {err_msg}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        elif origen.es_error_carga or origen.es_error_configuracion or origen.es_termino_invalido:messagebox.showerror("Error Búsqueda",err_msg or f"Error: {origen.name}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        elif origen==OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC:self.resultados_actuales=res_df;self._actualizar_estado(f"'{term_ui}': {len(fcds) if fcds is not None else 0} en Dic, {len(res_df) if res_df is not None else 0} en Desc.")
        elif origen==OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS:self.resultados_actuales=res_df;self._actualizar_estado(f"'{term_ui}': No en Diccionario.");messagebox.showinfo("Sin coincidencias",f"'{term_ui}' no en Diccionario.")
        elif origen in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC,OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS]:
            self.resultados_actuales=res_df;num_fcds_i=len(fcds) if fcds is not None else 0;msg_fcd_i=f"{num_fcds_i} en Diccionario"
            msg_desc_i="pero no se extrajeron términos válidos para Desc." if origen==OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS else "pero 0 resultados en Desc."
            self._actualizar_estado(f"'{term_ui}': {msg_fcd_i}, {msg_desc_i.replace('.','')} en Desc.")
            if messagebox.askyesno("Búsqueda Alternativa",f"{msg_fcd_i} para '{term_ui}', {msg_desc_i}\n\nBuscar '{term_ui}' directo en Descripciones?"):
                self._actualizar_estado(f"Buscando directo '{term_ui}' en desc...");self.indices_fcds_resaltados=None
                if self.motor.datos_diccionario is not None: 
                    cols_prev_dic_alt,_=self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario,self.motor.indices_columnas_busqueda_dic_preview,"diccionario_preview")
                    self.lbl_tabla_diccionario.config(text=f"Vista Previa Diccionario ({len(self.motor.datos_diccionario)} filas)");self._actualizar_tabla(self.tabla_diccionario,self.motor.datos_diccionario,limite_filas=100,columnas_a_mostrar=cols_prev_dic_alt,indices_a_resaltar=None)
                res_df_dir,orig_dir,_,_,err_msg_dir=self.motor.buscar(termino_busqueda_original=term_ui,buscar_via_diccionario_flag=False) # CORREGIDO
                self.origen_principal_resultados=orig_dir;self.fcds_de_ultima_busqueda=None
                if err_msg_dir and orig_dir.es_error_operacional:messagebox.showerror("Error Búsqueda Directa",f"Error: {err_msg_dir}");self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
                else:self.resultados_actuales=res_df_dir
                num_rdd=len(self.resultados_actuales) if self.resultados_actuales is not None else 0;self._actualizar_estado(f"Búsqueda directa '{term_ui}': {num_rdd} resultados.")
                if num_rdd==0 and orig_dir==OrigenResultados.DIRECTO_DESCRIPCION_VACIA and term_ui.strip():messagebox.showinfo("Info",f"No resultados para '{term_ui}' en búsqueda directa.")
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS:self.resultados_actuales=res_df;self._actualizar_estado(f"Búsqueda directa '{term_ui}': {len(res_df) if res_df is not None else 0} resultados.")
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            self.resultados_actuales=res_df;num_r=len(res_df) if res_df is not None else 0
            self._actualizar_estado(f"Mostrando todas las desc ({num_r})." if not term_ui.strip() else f"Búsqueda directa '{term_ui}': 0 resultados.")
            if term_ui.strip() and num_r==0 :messagebox.showinfo("Info",f"No resultados para '{term_ui}' en búsqueda directa.")
        if self.resultados_actuales is None:self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        self.desc_finales_de_ultima_busqueda=self.resultados_actuales.copy();self._actualizar_tabla(self.tabla_resultados,self.resultados_actuales);self._actualizar_botones_estado_general()

    def _salvar_regla_actual(self): 
        origen_nombre = self.origen_principal_resultados.name
        if not self.ultimo_termino_buscado and not (self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and self.desc_finales_de_ultima_busqueda is not None): messagebox.showerror("Error Salvar", "No hay búsqueda para salvar."); return
        df_salvar = None; tipo_datos = "DESCONOCIDO"
        if self.origen_principal_resultados.es_via_diccionario:
            if self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty: df_salvar = self.desc_finales_de_ultima_busqueda; tipo_datos = "DESC_VIA_DICC"
            elif self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty: df_salvar = self.fcds_de_ultima_busqueda; tipo_datos = "FCDS_DICC"
        elif self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            if self.desc_finales_de_ultima_busqueda is not None: df_salvar = self.desc_finales_de_ultima_busqueda; tipo_datos = "DESC_DIRECTA"; 
            if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and not self.ultimo_termino_buscado.strip(): tipo_datos = "TODAS_DESC"
        if df_salvar is not None:
            regla = {"termino": self.ultimo_termino_buscado or "N/A", "origen": origen_nombre, "tipo": tipo_datos, "filas": len(df_salvar), "ts": pd.Timestamp.now().isoformat()}
            self.reglas_guardadas.append(regla); self._actualizar_estado(f"Búsqueda '{self.ultimo_termino_buscado}' registrada."); messagebox.showinfo("Regla Salvada", f"Metadatos de '{self.ultimo_termino_buscado}' guardados.")
        else: messagebox.showwarning("Nada que Salvar", "No hay datos claros para salvar.")
        self._actualizar_botones_estado_general()
    def _exportar_resultados(self): 
        if self.resultados_actuales is None or self.resultados_actuales.empty: messagebox.showinfo("Exportar", "No hay resultados para exportar."); return
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")], title="Guardar resultados", initialfile=f"resultados_{pd.Timestamp.now():%Y%m%d_%H%M%S}")
        if not ruta: return
        try:
            if ruta.endswith(".xlsx"): self.resultados_actuales.to_excel(ruta, index=False)
            elif ruta.endswith(".csv"): self.resultados_actuales.to_csv(ruta, index=False, encoding='utf-8-sig')
            else: messagebox.showerror("Error Formato", "Usar .xlsx o .csv."); return
            messagebox.showinfo("Exportado", f"Resultados exportados a:\n{ruta}"); self._actualizar_estado(f"Exportado a {Path(ruta).name}")
        except Exception as e: logger.exception("Error exportando"); messagebox.showerror("Error Exportar", f"Fallo al exportar:\n{e}")
    def _actualizar_estado_botones_operadores(self): 
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None: self._deshabilitar_botones_operadores(); return
        [btn.config(state="normal") for btn in self.op_buttons.values()]
        txt=self.texto_busqueda_var.get();cur_pos=self.entrada_busqueda.index(tk.INSERT)
        last_char=txt[:cur_pos].strip()[-1:] if txt[:cur_pos].strip() else ""
        op_keys_logicos = ["+", "|"]; op_keys_comparacion_prefijo = [">", "<", ">=", "<=", "-"]
        if not last_char or last_char in op_keys_logicos + ["/","#","<",">"," ","="]: 
            if self.op_buttons.get("+"): self.op_buttons["+"]["state"]="disabled"
            if self.op_buttons.get("|"): self.op_buttons["|"]["state"]="disabled"
        if last_char and last_char not in op_keys_logicos + ["/"," ","="]: 
            if self.op_buttons.get("#"): self.op_buttons["#"]["state"]="disabled"
        if last_char in [">","<","="]: 
            for opk in op_keys_comparacion_prefijo + ["="]: 
                 if self.op_buttons.get(opk): self.op_buttons[opk]["state"]="disabled"
    def _insertar_operador_validado(self,op_limpio): 
        insert_txt=f" {op_limpio} " if op_limpio in ["+","|","/","<",">",">=","<=","-","="] else f"{op_limpio} " 
        self.entrada_busqueda.insert(tk.INSERT,insert_txt);self.entrada_busqueda.focus_set()
    def _deshabilitar_botones_operadores(self): [btn.config(state="disabled") for btn in self.op_buttons.values()]
    def on_closing(self): logger.info("Cerrando...");self._guardar_configuracion();self.destroy()

if __name__ == "__main__":
    log_file_name = "Buscador_Definitivo_v1.3_Corregido.log" # Nombre de log actualizado
    logging.basicConfig(
        level=logging.DEBUG, 
        format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s",
        handlers=[
            logging.FileHandler(log_file_name, encoding="utf-8", mode="w"), 
            logging.StreamHandler(), 
        ],
    )
    logger.info(f"--- Iniciando Buscador Definitivo v1.3 (Corregido) ({Path(__file__).name}) ---")
    
    missing_deps=[]
    try:import pandas as pd;logger.info(f"Pandas: {pd.__version__}")
    except ImportError:missing_deps.append("pandas")
    try:import openpyxl;logger.info(f"openpyxl: {openpyxl.__version__}")
    except ImportError:missing_deps.append("openpyxl")
    try:import xlrd; logger.info(f"xlrd: {xlrd.__version__}") 
    except ImportError: logger.warning("xlrd no encontrado. Carga de .xls antiguos podría fallar.")
    try:import numpy as np;logger.info(f"Numpy: {np.__version__}")
    except ImportError:missing_deps.append("numpy")
    
    critical_deps_missing = [dep for dep in ["pandas", "numpy", "openpyxl"] if dep in missing_deps]
    if critical_deps_missing:
        err_msg=f"Faltan dependencias críticas: {', '.join(critical_deps_missing)}. Instale con: pip install pandas openpyxl numpy xlrd"
        logger.critical(err_msg)
        try:r=tk.Tk();r.withdraw();messagebox.showerror("Dependencias Faltantes",err_msg);r.destroy()
        except:print(f"ERROR CRITICO: {err_msg}")
        exit(1)
    
    try:app=InterfazGrafica();app.mainloop()
    except Exception as e_main:logger.critical("Error fatal en app!",exc_info=True)
    finally:logger.info(f"--- Finalizando Buscador ---")
