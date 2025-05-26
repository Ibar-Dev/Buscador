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
# import string # No se usa directamente, pero puede ser útil para futuras extensiones.
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
    # Nuevo origen para búsquedas puramente negativas vía diccionario
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
        
        # Para 'diccionario_fcds_inicial' y 'descripcion', usar todas las columnas
        # si no se especifica lo contrario (o si indices_cfg está vacío/[-1] y no hay texto/obj).
        if tipo_busqueda in ["diccionario_fcds_inicial", "descripcion_fcds", "descripcion"]: 
             if not indices_cfg or indices_cfg == [-1]: # Si no hay config específica o es "todas texto/obj"
                cols_texto_obj = [col for col in columnas_disponibles if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
                if cols_texto_obj:
                    logger.debug(f"'{tipo_busqueda}': Usando cols texto/obj (config por defecto): {cols_texto_obj}")
                    return cols_texto_obj, None
                else: # Si no hay columnas de texto/objeto, usar todas como fallback
                    logger.warning(f"'{tipo_busqueda}': No hay cols texto/obj, usando todas las {num_cols_df} columnas: {columnas_disponibles}")
                    return columnas_disponibles, None

        # Para 'diccionario_preview' u otros tipos que usan 'indices_cfg' explícitamente
        if not indices_cfg or indices_cfg == [-1]: # Si es preview y no hay config específica, usa texto/obj o todas
            cols_texto_obj = [col for col in columnas_disponibles if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
            if cols_texto_obj: 
                logger.debug(f"'{tipo_busqueda}': Usando cols texto/obj (config por defecto para preview): {cols_texto_obj}")
                return cols_texto_obj, None
            logger.warning(f"'{tipo_busqueda}': No hay cols texto/obj para preview, usando todas: {columnas_disponibles}"); 
            return columnas_disponibles, None
        
        nombres_seleccionadas = [] 
        for i in indices_cfg:
            if not (isinstance(i, int) and 0 <= i < num_cols_df): 
                return None, f"Índice {i} inválido para '{tipo_busqueda}' (total cols: {num_cols_df})."
            nombres_seleccionadas.append(columnas_disponibles[i])
        
        if not nombres_seleccionadas and indices_cfg: 
            return None, f"Índices {indices_cfg} no resultaron en columnas válidas para '{tipo_busqueda}'."
        
        logger.debug(f"'{tipo_busqueda}': Usando columnas por índice {indices_cfg}: {nombres_seleccionadas}")
        return nombres_seleccionadas, None
    
    def _normalizar_para_busqueda(self, texto: str) -> str:
        if not isinstance(texto, str) or not texto: return ""
        try:
            texto_upper = texto.upper()
            texto_norm_nfkd = unicodedata.normalize('NFKD', texto_upper)
            texto_sin_acentos = "".join([c for c in texto_norm_nfkd if not unicodedata.combining(c)])
            # Permitir caracteres alfanuméricos, espacios y algunos especiales que pueden ser parte de términos técnicos
            texto_limpio_final = re.sub(r'[^\w\s\.\-\/\_]', '', texto_sin_acentos) 
            return ' '.join(texto_limpio_final.split()).strip()
        except Exception as e: 
            logger.error(f"Error normalizando '{texto[:50]}...': {e}")
            return str(texto).upper().strip()

    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, cols: List[str], texto: str) -> Tuple[pd.DataFrame, str, List[str]]:
        texto_limpio = texto.strip()
        negados_lista: List[str] = []
        
        df_a_procesar = df_original.copy() if df_original is not None else pd.DataFrame()

        if not texto_limpio: 
            return df_a_procesar, "", negados_lista

        positivos_parts, last_end = [], 0
        for m in self.patron_termino_negado.finditer(texto_limpio):
            positivos_parts.append(texto_limpio[last_end:m.start()])
            last_end = m.end()
            term_raw = m.group(1) or m.group(2) 
            if term_raw: 
                norm_negado = self._normalizar_para_busqueda(term_raw.strip('"')) 
                if norm_negado and norm_negado not in negados_lista: 
                    negados_lista.append(norm_negado) 
        
        positivos_parts.append(texto_limpio[last_end:])
        positivos_str = ' '.join("".join(positivos_parts).split()).strip()
        
        # logger.debug(f"Parseo de negaciones: Original='{texto_limpio}', Positivos='{positivos_str}', Negados={negados_lista}") # Log más conciso abajo

        if df_a_procesar.empty or not negados_lista or not cols:
            return df_a_procesar, positivos_str, negados_lista
            
        # logger.debug(f"Aplicando negación con términos: {negados_lista} en {len(df_a_procesar)} filas.") # Log más conciso abajo
        mascara_excluir_total = pd.Series(False, index=df_a_procesar.index)
        
        for term_neg in negados_lista:
            if not term_neg: continue
            mascara_term_actual = pd.Series(False, index=df_a_procesar.index)
            for col_nombre in cols:
                if col_nombre not in df_a_procesar.columns: continue
                try:
                    serie_norm_df = df_a_procesar[col_nombre].astype(str).map(self._normalizar_para_busqueda)
                    pat_regex = r"\b" + re.escape(term_neg) + r"\b"
                    mascara_term_actual |= serie_norm_df.str.contains(pat_regex, regex=True, na=False)
                except Exception as e_neg: 
                    logger.error(f"Error aplicando negación en columna '{col_nombre}', término '{term_neg}': {e_neg}")
            mascara_excluir_total |= mascara_term_actual
            
        df_resultado_final = df_a_procesar[~mascara_excluir_total]
        logger.info(f"Filtrado por negación (original='{texto_limpio}'): {len(df_a_procesar)} -> {len(df_resultado_final)} filas. Negados aplicados: {negados_lista}. Positivos para query: '{positivos_str}'")
        return df_resultado_final, positivos_str, negados_lista

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        texto_limpio = texto_complejo.strip(); 
        if not texto_limpio: return "OR", [] 
        separadores_or = [
            (r"\s*\|\s*", "|"), 
            (r"\s*/\s*", "/")  
        ]
        for sep_regex, sep_char in separadores_or:
            if sep_char in texto_limpio:
                segmentos = [s.strip() for s in re.split(sep_regex, texto_limpio) if s.strip()]
                if len(segmentos) > 1 or (len(segmentos) == 1 and texto_limpio != segmentos[0]):
                    # logger.debug(f"Descomposición Nivel 1 (OR) para '{texto_complejo}': Op=OR, Segmentos={segmentos}")
                    return "OR", segmentos
        
        # logger.debug(f"Descomposición Nivel 1 (OR) para '{texto_complejo}': Op=AND (sin OR explícito), Segmento=['{texto_limpio}']")
        return "AND", [texto_limpio] 

    def _descomponer_nivel2_and(self, termino_segmento_n1: str) -> Tuple[str, List[str]]:
        termino_limpio = termino_segmento_n1.strip(); 
        if not termino_limpio: return "AND", []
        partes = re.split(r'\s+\+\s+', termino_limpio)
        partes_limpias = [p.strip() for p in partes if p.strip()]
        # logger.debug(f"Descomposición Nivel 2 (AND) para '{termino_segmento_n1}': Partes={partes_limpias}")
        return "AND", partes_limpias 

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        analizados = []
        for term_orig_bruto in terminos_brutos:
            term_orig_procesado = str(term_orig_bruto).strip()
            
            if len(term_orig_procesado) >= 2 and term_orig_procesado.startswith('"') and term_orig_procesado.endswith('"'):
                term_orig_final = term_orig_procesado[1:-1]
            else:
                term_orig_final = term_orig_procesado
            
            if not term_orig_final: continue

            item_analizado: Dict[str, Any] = {"original": term_orig_final} 
            
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
        logger.debug(f"Términos (post-AND) analizados para búsqueda detallada: {analizados}")
        return analizados

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        if isinstance(num_str, (int, float)): return float(num_str)
        if not isinstance(num_str, str): return None
        try: return float(num_str.replace(",", ".")) 
        except ValueError: return None 

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols: List[str], term_an: Dict[str, Any]) -> pd.Series:
        tipo, valor, unidad_req_canon = term_an["tipo"], term_an["valor"], term_an.get("unidad_busqueda")
        mascara_total = pd.Series(False, index=df.index)
        # logger.debug(f"Generando máscara para término: tipo='{tipo}', valor='{valor}', unidad_req='{unidad_req_canon}'") # Verboso
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
                            u_ok = (unidad_req_canon is None) or \
                                   (u_c_canon is not None and u_c_canon == unidad_req_canon) or \
                                   (u_c_raw and unidad_req_canon and self.extractor_magnitud._normalizar_texto(u_c_raw.strip()) == unidad_req_canon) 
                            if not u_ok: continue
                            cond = False
                            if tipo == "eq" and np.isclose(num_c_val, valor): cond = True
                            elif tipo == "gt" and num_c_val > valor and not np.isclose(num_c_val, valor): cond = True
                            elif tipo == "lt" and num_c_val < valor and not np.isclose(num_c_val, valor): cond = True
                            elif tipo == "ge" and (num_c_val >= valor or np.isclose(num_c_val, valor)): cond = True
                            elif tipo == "le" and (num_c_val <= valor or np.isclose(num_c_val, valor)): cond = True
                            elif tipo == "range" and ((valor[0] <= num_c_val or np.isclose(num_c_val, valor[0])) and \
                                                      (num_c_val <= valor[1] or np.isclose(num_c_val, valor[1]))): cond = True
                            if cond: mascara_col_actual_num.at[idx] = True; break 
                        except ValueError: continue 
                mascara_total |= mascara_col_actual_num
            elif tipo == "str":
                try:
                    val_norm_busq = str(valor); 
                    if not val_norm_busq: continue 
                    serie_norm_df_col = col_s.astype(str).map(self._normalizar_para_busqueda)
                    pat_regex = r"\b" + re.escape(val_norm_busq) + r"\b"
                    mascara_col_actual = serie_norm_df_col.str.contains(pat_regex, regex=True, na=False)
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
        if not lista_mascaras: 
            return pd.Series(False, index=df_idx_ref) if df_idx_ref is not None else pd.Series(dtype=bool) 
        
        idx_usar = df_idx_ref if df_idx_ref is not None else lista_mascaras[0].index
        if idx_usar.empty and not lista_mascaras[0].empty: 
            idx_usar = lista_mascaras[0].index

        mascara_final = pd.Series(False, index=idx_usar) 
        for masc_seg in lista_mascaras:
            if masc_seg.empty: continue 
            
            if not masc_seg.index.equals(idx_usar):
                try: 
                    masc_seg = masc_seg.reindex(idx_usar, fill_value=False)
                except Exception as e_reidx: 
                    logger.error(f"Fallo reindex máscara OR: {e_reidx}. Máscara ignorada."); continue 
            mascara_final |= masc_seg 
        return mascara_final

    def _procesar_busqueda_en_df_objetivo(self, df_obj: pd.DataFrame, cols_obj: List[str], term_busq_orig: str, terminos_negativos_adicionales: Optional[List[str]] = None) -> Tuple[pd.DataFrame, Optional[str]]:
        logger.debug(f"Procesando búsqueda: Query='{term_busq_orig}' en {len(cols_obj)} cols del DF ({len(df_obj)} filas). Neg. Adicionales: {terminos_negativos_adicionales}")
        
        df_post_neg_inicial, term_pos_q_de_query, negados_de_query = self._aplicar_negaciones_y_extraer_positivos(df_obj, cols_obj, term_busq_orig)
        
        df_final_post_neg = df_post_neg_inicial # Empezamos con el DF ya filtrado por negaciones de la query
        
        # Aplicar negativos adicionales si existen y no fueron ya cubiertos por negados_de_query
        if terminos_negativos_adicionales and not df_final_post_neg.empty:
            negativos_solo_adicionales_norm = []
            negados_de_query_norm_set = set(negados_de_query) # Ya están normalizados
            for n_adic_raw in terminos_negativos_adicionales:
                n_adic_norm = self._normalizar_para_busqueda(n_adic_raw)
                if n_adic_norm and n_adic_norm not in negados_de_query_norm_set:
                    negativos_solo_adicionales_norm.append(n_adic_norm)
            
            if negativos_solo_adicionales_norm:
                logger.debug(f"Aplicando negativos ADICIONALES (no en query original): {negativos_solo_adicionales_norm} a {len(df_final_post_neg)} filas.")
                mascara_excluir_adicional = pd.Series(False, index=df_final_post_neg.index)
                for term_neg_adic_norm in negativos_solo_adicionales_norm:
                    mascara_term_actual_adic = pd.Series(False, index=df_final_post_neg.index)
                    for col_nombre in cols_obj:
                        if col_nombre not in df_final_post_neg.columns: continue
                        try:
                            serie_norm_df = df_final_post_neg[col_nombre].astype(str).map(self._normalizar_para_busqueda)
                            pat_regex = r"\b" + re.escape(term_neg_adic_norm) + r"\b"
                            mascara_term_actual_adic |= serie_norm_df.str.contains(pat_regex, regex=True, na=False)
                        except Exception as e_neg_adic: logger.error(f"Error negación adicional col '{col_nombre}', término '{term_neg_adic_norm}': {e_neg_adic}")
                    mascara_excluir_adicional |= mascara_term_actual_adic
                df_final_post_neg = df_final_post_neg[~mascara_excluir_adicional]
                logger.info(f"Filtrado por neg. ADICIONAL: {len(df_post_neg_inicial)} -> {len(df_final_post_neg)} filas.")

        term_pos_q_final_para_parseo = term_pos_q_de_query # La parte positiva de la query original no cambia

        if df_final_post_neg.empty and not term_pos_q_final_para_parseo.strip(): 
            logger.debug("DF vacío después de todas las negaciones y sin términos positivos en query.")
            return df_final_post_neg.copy(), None 
        
        if not term_pos_q_final_para_parseo.strip():
            logger.debug(f"Sin términos positivos en query ('{term_pos_q_final_para_parseo}'). Devolviendo DF post-negaciones.")
            return df_final_post_neg.copy(), None

        op_n1, segs_n1 = self._descomponer_nivel1_or(term_pos_q_final_para_parseo)
        
        if not segs_n1: 
            if term_busq_orig.strip() or term_pos_q_final_para_parseo.strip():
                logger.warning(f"Término positivo '{term_pos_q_final_para_parseo}' (de '{term_busq_orig}') inválido post-OR. Ningún segmento.")
                return pd.DataFrame(columns=df_final_post_neg.columns), "Término positivo inválido (sin segmentos OR)."
            else: 
                logger.debug("Query original y positiva post-negación vacías. Devolviendo df post-negaciones.")
                return df_final_post_neg.copy(), None
                
        lista_mascaras_or = []
        for seg_n1 in segs_n1: 
            op_n2, terms_brutos_n2 = self._descomponer_nivel2_and(seg_n1) 
            terms_atom_an = self._analizar_terminos(terms_brutos_n2) 
            
            if not terms_atom_an: 
                if op_n1 == "AND": 
                    logger.warning(f"Segmento AND '{seg_n1}' sin términos atómicos válidos. Segmento falla.")
                    return pd.DataFrame(columns=df_final_post_neg.columns), f"Segmento AND '{seg_n1}' inválido."
                logger.debug(f"Segmento OR '{seg_n1}' sin términos atómicos válidos. Se ignora para el OR.")
                mascara_seg_n1 = pd.Series(False, index=df_final_post_neg.index)
            else:
                 mascara_seg_n1 = self._aplicar_mascara_combinada_para_segmento_and(df_final_post_neg, cols_obj, terms_atom_an)
            
            lista_mascaras_or.append(mascara_seg_n1)
            
        if not lista_mascaras_or: 
             logger.error("Error interno: no se generaron máscaras OR a pesar de haber segmentos N1.")
             return pd.DataFrame(columns=df_final_post_neg.columns), "Error interno: no se generaron máscaras OR."
             
        mascara_final_df_objetivo = self._combinar_mascaras_de_segmentos_or(lista_mascaras_or, df_final_post_neg.index)
        df_resultado = df_final_post_neg[mascara_final_df_objetivo].copy()
        
        logger.debug(f"Resultado de _procesar_busqueda_en_df_objetivo para '{term_busq_orig}': {len(df_resultado)} filas.")
        return df_resultado, None

    def _extraer_terminos_de_fila_completa(self, fila_df: pd.Series) -> Set[str]:
        terminos_extraidos: Set[str] = set()
        if fila_df is None or fila_df.empty: 
            return terminos_extraidos
        for valor_celda in fila_df.values: 
            if pd.notna(valor_celda): 
                texto_celda_str = str(valor_celda).strip()
                if texto_celda_str: 
                    texto_celda_norm = self._normalizar_para_busqueda(texto_celda_str)
                    palabras_celda = [palabra for palabra in texto_celda_norm.split() if len(palabra) > 1 and not palabra.isdigit()]
                    if palabras_celda: 
                        terminos_extraidos.update(palabras_celda)
                    elif texto_celda_norm and len(texto_celda_norm) > 1 and not texto_celda_norm.isdigit(): 
                        terminos_extraidos.add(texto_celda_norm)
        return terminos_extraidos

    def buscar(self, termino_busqueda_original: str, buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[List[int]], Optional[str]]:
        logger.info(f"Motor.buscar INICIO: termino='{termino_busqueda_original}', via_dicc={buscar_via_diccionario_flag}")
        df_vacio_desc = pd.DataFrame(columns=(self.datos_descripcion.columns if self.datos_descripcion is not None else []))
        fcds_obtenidos_final: Optional[pd.DataFrame] = None
        indices_fcds_a_resaltar: Optional[List[int]] = None

        if not termino_busqueda_original.strip():
            return (self.datos_descripcion.copy() if self.datos_descripcion is not None else df_vacio_desc), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, (None if self.datos_descripcion is not None else "Descripciones no cargadas.")

        if buscar_via_diccionario_flag:
            if self.datos_diccionario is None: return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, None, "Diccionario no cargado."
            cols_dic_fcds, err_cols_dic = self._obtener_nombres_columnas_busqueda_df(self.datos_diccionario, [], "diccionario_fcds_inicial")
            if not cols_dic_fcds: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, None, err_cols_dic or "No se pudo det. cols para FCDs."

            df_dummy_para_parseo = pd.DataFrame()
            _, term_busq_positivos_globales_str, term_busq_negativos_globales_list = \
                self._aplicar_negaciones_y_extraer_positivos(
                    df_dummy_para_parseo, [], termino_busqueda_original
                )
            logger.info(f"Parseo global inicial: Positivos='{term_busq_positivos_globales_str}', Negativos Globales={term_busq_negativos_globales_list}")

            origen_propuesto_inicial = OrigenResultados.NINGUNO # Para determinar el flujo

            if term_busq_positivos_globales_str.strip():
                logger.info(f"BUSCAR EN DICC (FCDs) - Caso Positivos: Query para FCDs='{term_busq_positivos_globales_str}'")
                origen_propuesto_inicial = OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC # Asume éxito con positivos
                try:
                    fcds_obtenidos_para_positivos, error_dic = self._procesar_busqueda_en_df_objetivo(
                        self.datos_diccionario, cols_dic_fcds, term_busq_positivos_globales_str,
                        terminos_negativos_adicionales=None # Neg. globales se aplican a descripciones
                    )
                    if error_dic: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic
                    fcds_obtenidos_final = fcds_obtenidos_para_positivos
                except Exception as e_dic: 
                    logger.exception("Excepción en búsqueda en diccionario (caso positivos)."); 
                    return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-positivos): {e_dic}"
            
            elif term_busq_negativos_globales_list: # Positivos vacíos, pero hay negativos globales
                logger.info(f"BUSCAR EN DICC (FCDs) - Caso Puramente Negativo: Negativos Globales para FCDs={term_busq_negativos_globales_list}")
                origen_propuesto_inicial = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC
                try:
                    fcds_filtradas_por_negativos, error_dic_neg = self._procesar_busqueda_en_df_objetivo(
                        self.datos_diccionario, cols_dic_fcds, "", # Sin términos positivos para el parseo interno
                        terminos_negativos_adicionales=term_busq_negativos_globales_list # Estos son los que filtran
                    )
                    if error_dic_neg: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_neg
                    fcds_obtenidos_final = fcds_filtradas_por_negativos
                    term_busq_negativos_globales_list = [] # Ya se usaron para FCDs, no reaplicar a desc.
                except Exception as e_dic_neg:
                    logger.exception("Excepción en búsqueda en diccionario (caso puramente negativo)."); 
                    return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-negativo): {e_dic_neg}"
            else: # Query original era vacía o solo espacios.
                logger.debug("Query original vacía o solo espacios, ya manejado.") # Debería haber sido capturado antes.
                return df_vacio_desc, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None


            if fcds_obtenidos_final is not None and not fcds_obtenidos_final.empty:
                indices_fcds_a_resaltar = fcds_obtenidos_final.index.tolist()
                logger.info(f"FCDs obtenidas del diccionario: {len(fcds_obtenidos_final)} filas.")
            else:
                logger.info(f"No se encontraron FCDs en el diccionario para la parte relevante de '{termino_busqueda_original}'.")
                return df_vacio_desc, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None

            if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final, indices_fcds_a_resaltar, "Descripciones no cargadas."
            
            terminos_para_descripcion: Set[str] = set()
            logger.info(f"Extrayendo términos de {len(fcds_obtenidos_final)} FCDs para buscar en descripciones.")
            for _, fila_fcd in fcds_obtenidos_final.iterrows():
                terminos_para_descripcion.update(self._extraer_terminos_de_fila_completa(fila_fcd))
            
            if not terminos_para_descripcion:
                logger.info("FCDs encontrados, pero no se extrajeron términos para buscar en descripciones.")
                origen_final = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                if origen_propuesto_inicial == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                    origen_final = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                return df_vacio_desc, origen_final, fcds_obtenidos_final, indices_fcds_a_resaltar, None
            
            logger.info(f"Términos para descripciones ({len(terminos_para_descripcion)} únicos, muestra): {sorted(list(terminos_para_descripcion))[:10]}...")
            try:
                terminos_con_comillas_si_necesario = [f'"{t}"' if " " in t else t for t in terminos_para_descripcion if t]
                termino_or_para_desc = " | ".join(terminos_con_comillas_si_necesario)
                
                if not termino_or_para_desc: 
                    origen_final = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                    if origen_propuesto_inicial == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                        origen_final = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                    return df_vacio_desc, origen_final, fcds_obtenidos_final, indices_fcds_a_resaltar, "Query OR para descripciones vacía."
                
                cols_desc_para_fcds, err_cols_desc_fcd = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                if not cols_desc_para_fcds: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final, indices_fcds_a_resaltar, err_cols_desc_fcd
                
                logger.info(f"BUSCAR EN DESC (vía FCD): Query OR (muestra): '{termino_or_para_desc[:100]}...'. Neg. Globales a aplicar: {term_busq_negativos_globales_list}")
                
                resultados_desc_via_dic, error_proc_desc_fcd = self._procesar_busqueda_en_df_objetivo(
                    self.datos_descripcion, cols_desc_para_fcds, termino_or_para_desc,
                    terminos_negativos_adicionales=term_busq_negativos_globales_list 
                )
                
                if error_proc_desc_fcd: return df_vacio_desc, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final, indices_fcds_a_resaltar, error_proc_desc_fcd
                
                if resultados_desc_via_dic is None or resultados_desc_via_dic.empty:
                    origen_final = OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                    if origen_propuesto_inicial == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                        origen_final = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                    return df_vacio_desc, origen_final, fcds_obtenidos_final, indices_fcds_a_resaltar, None
                else:
                    return resultados_desc_via_dic, origen_propuesto_inicial, fcds_obtenidos_final, indices_fcds_a_resaltar, None
            
            except Exception as e_desc_fcd: 
                logger.exception("Excepción búsqueda en descripciones vía FCDs."); 
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos_final, indices_fcds_a_resaltar, f"Error motor (desc vía FCD): {e_desc_fcd}"

        else: # Búsqueda directa en descripciones
            if self.datos_descripcion is None: return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas."
            cols_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion")
            if not cols_desc_directo: return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_directo
            try:
                logger.info(f"BUSCAR EN DESC (DIRECTO): Query '{termino_busqueda_original}' en descripciones.")
                resultados_directos, error_proc_desc_directo = self._procesar_busqueda_en_df_objetivo(
                    self.datos_descripcion, cols_desc_directo, termino_busqueda_original,
                    terminos_negativos_adicionales=None # En modo directo, todos los '#' son de la query original
                )
                if error_proc_desc_directo: return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_proc_desc_directo
                
                if resultados_directos is None or resultados_directos.empty: 
                    return df_vacio_desc, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
                else: 
                    return resultados_directos, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None, None
            except Exception as e_desc_dir: 
                logger.exception("Excepción búsqueda directa en descripciones."); 
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (desc directa): {e_desc_dir}"

# --- Interfaz Gráfica ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE = "config_buscador_definitivo_v1.5_refactorizado.json" # Actualizado
    def __init__(self):
        super().__init__()
        self.title("Buscador Definitivo v1.5 (Refactorizado)") # Actualizado
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
        self.color_fila_par = "white"; self.color_fila_impar = "#f0f0f0"; self.color_resaltado_dic = "sky blue" 
        self.op_buttons: Dict[str, ttk.Button] = {}
        self._configurar_estilo_ttk(); self._crear_widgets(); self._configurar_grid(); self._configurar_eventos()
        self._configurar_tags_treeview(); self._configurar_orden_tabla(self.tabla_resultados); self._configurar_orden_tabla(self.tabla_diccionario)
        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones."); self._deshabilitar_botones_operadores(); self._actualizar_botones_estado_general()
        logger.info(f"Interfaz Gráfica (Definitiva v1.5 Refactorizado) inicializada.")

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
                    "Flujo de Búsqueda Vía Diccionario:\n"
                    "1. Query 'A + #B': Parte '#B' se separa. Parte 'A' se busca en Diccionario (FCDs).\n"
                    "2. Query '#B': Se buscan FCDs que NO contengan 'B'.\n"
                    "3. Filas FCD del Diccionario se resaltan.\n"
                    "4. TODAS las palabras de esas filas FCD se extraen (forman Query_OR_Desc).\n"
                    "5. Query_OR_Desc se busca en Descripciones.\n"
                    "6. Si query era 'A + #B', ahora '#B' filtra los resultados de Descripciones.\n"
                    "7. Si no hay FCDs, se ofrece búsqueda directa del término original en Descripciones.\n"
                    "8. Búsqueda vacía muestra todas las Descripciones.")
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
        is_dicc=tabla==self.tabla_diccionario; tabla_nombre = "Diccionario" if is_dicc else "Resultados"
        # logger.debug(f"Actualizando tabla {tabla_nombre} con {len(datos) if datos is not None else 0} filas. Resaltar: {indices_a_resaltar is not None}") # Verboso
        [tabla.delete(i) for i in tabla.get_children()];tabla["columns"]=()
        if datos is None or datos.empty:self._configurar_orden_tabla(tabla); logger.debug(f"Tabla {tabla_nombre} vacía."); return
        
        cols_orig=list(datos.columns);cols_usar=cols_orig
        if columnas_a_mostrar:cols_usar=[c for c in columnas_a_mostrar if c in cols_orig] or cols_orig
        if not cols_usar:self._configurar_orden_tabla(tabla); logger.debug(f"Tabla {tabla_nombre} sin columnas usables."); return
        
        tabla["columns"]=tuple(cols_usar)
        for c in cols_usar:
            tabla.heading(c,text=str(c),anchor=tk.W)
            try:
                ancho_contenido = datos[c].astype(str).str.len().quantile(0.95) if not datos[c].empty else 0
                ancho_cabecera = len(str(c))
                ancho = max(70, min(int(max(ancho_cabecera * 7, ancho_contenido * 5.5) + 15), 350)) 
            except: ancho = 100 
            tabla.column(c,anchor=tk.W,width=ancho,minwidth=50) 

        df_iterar=datos[cols_usar];num_filas_original=len(df_iterar)
        
        mostrar_todo_por_resaltado = is_dicc and indices_a_resaltar and num_filas_original > 0
        if not mostrar_todo_por_resaltado and limite_filas and num_filas_original > limite_filas: 
            df_iterar=df_iterar.head(limite_filas); 
            # logger.debug(f"Mostrando {limite_filas} de {num_filas_original} en tabla {tabla_nombre}.") # Verboso
        elif mostrar_todo_por_resaltado: 
            logger.debug(f"Mostrando todas las {num_filas_original} filas del {tabla_nombre} para asegurar visibilidad de resaltados.")
        
        for i,(idx,row) in enumerate(df_iterar.iterrows()):
            vals=[str(v) if pd.notna(v) else "" for v in row.values];tags=["par" if i%2==0 else "impar"]
            if is_dicc and indices_a_resaltar and idx in indices_a_resaltar:tags.append("resaltado_azul")
            try:
                tabla.insert("","end",values=vals,tags=tuple(tags),iid=f"row_{idx}")
            except Exception as e_ins: logger.warning(f"Error insertando fila {idx} en treeview {tabla_nombre}: {e_ins}") 
        self._configurar_orden_tabla(tabla)
        logger.debug(f"Tabla {tabla_nombre} actualizada con {len(tabla.get_children())} filas visibles.")

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
            if self.origen_principal_resultados.es_via_diccionario and ((self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty)or(self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and self.origen_principal_resultados in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC] )):salvar_ok=True
            elif (self.origen_principal_resultados.es_directo_descripcion or self.origen_principal_resultados==OrigenResultados.DIRECTO_DESCRIPCION_VACIA) and self.desc_finales_de_ultima_busqueda is not None:salvar_ok=True
        self.btn_salvar_regla["state"]="normal" if salvar_ok else "disabled";self.btn_exportar["state"]="normal" if (self.resultados_actuales is not None and not self.resultados_actuales.empty) else "disabled"
    def _cargar_diccionario(self): 
        cfg_path=self.config.get("last_dic_path");init_dir=str(Path(cfg_path).parent) if cfg_path and Path(cfg_path).exists() else os.getcwd()
        ruta_seleccionada=filedialog.askopenfilename(title="Cargar Diccionario",filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")],initialdir=init_dir)
        if not ruta_seleccionada: return
        nombre_archivo = Path(ruta_seleccionada).name 
        self._actualizar_estado(f"Cargando dicc: {nombre_archivo}...")
        self._actualizar_tabla(self.tabla_diccionario,None);self._actualizar_tabla(self.tabla_resultados,None);self.resultados_actuales=None;self.fcds_de_ultima_busqueda=None;self.desc_finales_de_ultima_busqueda=None;self.origen_principal_resultados=OrigenResultados.NINGUNO;self.indices_fcds_resaltados=None
        ok,msg=self.motor.cargar_excel_diccionario(ruta_seleccionada)
        desc_n_title=Path(self.motor.archivo_descripcion_actual).name if self.motor.archivo_descripcion_actual else "N/A" 
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
            termino_busqueda_original=term_ui, 
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
        elif origen in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC]:
             self.resultados_actuales=res_df;self._actualizar_estado(f"'{term_ui}': {len(fcds) if fcds is not None else 0} en Dic, {len(res_df) if res_df is not None else 0} en Desc.")
        elif origen==OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS:
            self.resultados_actuales=res_df 
            self._actualizar_estado(f"'{term_ui}': No en Diccionario.");
            if messagebox.askyesno("Búsqueda Alternativa",f"'{term_ui}' no encontrado en Diccionario.\n\n¿Buscar '{term_ui}' directamente en Descripciones?"):
                self._buscar_directo_en_descripciones(term_ui, df_desc_cols) 
            else:
                self._actualizar_botones_estado_general() 
        elif origen in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC]:
            self.resultados_actuales=res_df;num_fcds_i=len(fcds) if fcds is not None else 0;msg_fcd_i=f"{num_fcds_i} en Diccionario"
            msg_desc_i="pero no se extrajeron términos válidos para Desc." if origen in [OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS] else "pero 0 resultados en Desc."
            self._actualizar_estado(f"'{term_ui}': {msg_fcd_i}, {msg_desc_i.replace('.','')} en Desc.")
            if messagebox.askyesno("Búsqueda Alternativa",f"{msg_fcd_i} para '{term_ui}', {msg_desc_i}\n\n¿Buscar '{term_ui}' directamente en Descripciones?"):
                 self._buscar_directo_en_descripciones(term_ui, df_desc_cols) 
            else:
                self._actualizar_botones_estado_general()
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS:self.resultados_actuales=res_df;self._actualizar_estado(f"Búsqueda directa '{term_ui}': {len(res_df) if res_df is not None else 0} resultados.")
        elif origen==OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            self.resultados_actuales=res_df;num_r=len(res_df) if res_df is not None else 0
            self._actualizar_estado(f"Mostrando todas las desc ({num_r})." if not term_ui.strip() else f"Búsqueda directa '{term_ui}': 0 resultados.")
            if term_ui.strip() and num_r==0 :messagebox.showinfo("Info",f"No resultados para '{term_ui}' en búsqueda directa.")
        
        if self.resultados_actuales is None:self.resultados_actuales=pd.DataFrame(columns=df_desc_cols)
        self.desc_finales_de_ultima_busqueda=self.resultados_actuales.copy();self._actualizar_tabla(self.tabla_resultados,self.resultados_actuales);self._actualizar_botones_estado_general()

    def _buscar_directo_en_descripciones(self, term_ui_original: str, df_desc_cols_ref: List[str]):
        """Función auxiliar para realizar una búsqueda directa en descripciones."""
        self._actualizar_estado(f"Buscando directo '{term_ui_original}' en descripciones...")
        self.indices_fcds_resaltados = None 
        if self.motor.datos_diccionario is not None: 
            cols_prev_dic_alt,_ = self.motor._obtener_nombres_columnas_busqueda_df(self.motor.datos_diccionario, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview")
            self.lbl_tabla_diccionario.config(text=f"Vista Previa Diccionario ({len(self.motor.datos_diccionario)} filas)")
            self._actualizar_tabla(self.tabla_diccionario, self.motor.datos_diccionario, limite_filas=100, columnas_a_mostrar=cols_prev_dic_alt, indices_a_resaltar=None)

        res_df_dir, orig_dir, _, _, err_msg_dir = self.motor.buscar(
            termino_busqueda_original=term_ui_original, 
            buscar_via_diccionario_flag=False
        )
        
        self.origen_principal_resultados = orig_dir
        self.fcds_de_ultima_busqueda = None 
        
        if err_msg_dir and orig_dir.es_error_operacional:
            messagebox.showerror("Error Búsqueda Directa", f"Error: {err_msg_dir}")
            self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref)
        elif err_msg_dir and orig_dir.es_termino_invalido:
            messagebox.showerror("Error Búsqueda Directa", f"Término inválido: {err_msg_dir}")
            self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref)
        else:
            self.resultados_actuales = res_df_dir

        num_rdd = len(self.resultados_actuales) if self.resultados_actuales is not None else 0
        self._actualizar_estado(f"Búsqueda directa '{term_ui_original}': {num_rdd} resultados.")
        if num_rdd == 0 and orig_dir == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and term_ui_original.strip():
            messagebox.showinfo("Info", f"No resultados para '{term_ui_original}' en búsqueda directa.")
        # La actualización de la tabla de resultados y botones se hará en el flujo principal de _ejecutar_busqueda

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
    log_file_name = "Buscador_Definitivo_v1.5_Refactorizado.log" 
    logging.basicConfig(
        level=logging.DEBUG, 
        format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s",
        handlers=[
            logging.FileHandler(log_file_name, encoding="utf-8", mode="w"), 
            logging.StreamHandler(), 
        ],
    )
    logger.info(f"--- Iniciando Buscador Definitivo v1.5 (Refactorizado) ({Path(__file__).name}) ---")
    
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