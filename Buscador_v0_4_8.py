# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Tuple, Union, Set, Callable, Dict, Any
import traceback
import platform
import unicodedata
import logging
import json
import os
import string # Para sanitizar nombres de archivo

# --- Configuración del Logging ---
# (Se configura en el bloque __main__)

# --- Clases de Lógica ---

class ManejadorExcel:
    """Clase estática para manejar la carga de archivos Excel."""
    @staticmethod
    def cargar_excel(ruta: str) -> Optional[pd.DataFrame]:
        logging.info(f"Intentando cargar archivo Excel: {ruta}")
        try:
            engine = 'openpyxl' if ruta.endswith('.xlsx') else None
            df = pd.read_excel(ruta, engine=engine)
            logging.info(f"Archivo '{os.path.basename(ruta)}' cargado ({len(df)} filas).")
            return df
        except FileNotFoundError:
            logging.error(f"Archivo no encontrado: {ruta}")
            messagebox.showerror("Error", f"Archivo no encontrado:\n{ruta}")
            return None
        except Exception as e:
            logging.exception(f"Error inesperado al cargar archivo: {ruta}")
            messagebox.showerror("Error al Cargar", f"No se pudo cargar:\n{ruta}\nError: {e}\nVerifique archivo y 'openpyxl'.")
            return None

class MotorBusqueda:
    """Gestiona datos y lógica de búsqueda."""
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[str] = None
        self.archivo_descripcion_actual: Optional[str] = None
        self.indices_columnas_busqueda_dic: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else [0, 3] # Default robusto
        logging.info(f"MotorBusqueda inicializado. Índices búsqueda diccionario: {self.indices_columnas_busqueda_dic}")
        self.patron_comparacion_compilado = re.compile(r"^([<>]=?)(\d+([.,]\d+)?).*$")
        self.patron_rango_compilado = re.compile(r"^(\d+([.,]\d+)?)-(\d+([.,]\d+)?)$")
        self.patron_negacion_compilado = re.compile(r"^#(.+)$")

    def cargar_excel_diccionario(self, ruta: str) -> bool:
        self.datos_diccionario = ManejadorExcel.cargar_excel(ruta)
        if self.datos_diccionario is None: self.archivo_diccionario_actual = None; return False
        self.archivo_diccionario_actual = ruta
        if not self._validar_columnas_diccionario():
            logging.warning("Validación columnas diccionario fallida. Invalidando carga.")
            self.datos_diccionario = None; self.archivo_diccionario_actual = None; return False
        return True

    def cargar_excel_descripcion(self, ruta: str) -> bool:
        self.datos_descripcion = ManejadorExcel.cargar_excel(ruta)
        if self.datos_descripcion is None: self.archivo_descripcion_actual = None; return False
        self.archivo_descripcion_actual = ruta
        return True

    def _validar_columnas_diccionario(self) -> bool:
        if self.datos_diccionario is None: return False
        num_cols = len(self.datos_diccionario.columns)
        if not self.indices_columnas_busqueda_dic: # Si la lista está vacía
            logging.error("La lista de índices de columnas de búsqueda está vacía en la configuración.")
            messagebox.showerror("Error Configuración", "No hay índices de columna definidos para la búsqueda en el diccionario.")
            return False
        max_indice_requerido = max(self.indices_columnas_busqueda_dic)

        if num_cols == 0:
            logging.error("Diccionario sin columnas.")
            messagebox.showerror("Error Diccionario", "Archivo diccionario vacío o sin columnas.")
            return False
        elif num_cols <= max_indice_requerido:
            logging.error(f"Diccionario tiene {num_cols} cols, necesita índice {max_indice_requerido}.")
            messagebox.showerror("Error Diccionario", f"Diccionario necesita {max_indice_requerido + 1} cols (índices {self.indices_columnas_busqueda_dic}), tiene {num_cols}.")
            return False
        return True

    def _obtener_nombres_columnas_busqueda(self, df: pd.DataFrame) -> Optional[List[str]]:
        if df is None: logging.error("Intento obtener cols de DataFrame nulo."); return None
        columnas_disponibles = df.columns; cols_encontradas_nombres = []; num_cols_df = len(columnas_disponibles)
        indices_validos = []
        for indice in self.indices_columnas_busqueda_dic:
            if isinstance(indice, int) and 0 <= indice < num_cols_df: # Validar índice
                cols_encontradas_nombres.append(columnas_disponibles[indice])
                indices_validos.append(indice)
            else: logging.warning(f"Índice {indice} inválido o fuera de rango (0-{num_cols_df-1}). Se omitirá.")
        if not cols_encontradas_nombres:
            logging.error(f"No se encontraron columnas válidas para índices: {self.indices_columnas_busqueda_dic}")
            messagebox.showerror("Error Diccionario", f"No hay columnas válidas para los índices configurados: {self.indices_columnas_busqueda_dic}")
            return None
        logging.debug(f"Columnas búsqueda diccionario: {cols_encontradas_nombres} (Índices: {indices_validos})")
        return cols_encontradas_nombres

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, cols_nombres: List[str]) -> Set[str]:
        # (Sin cambios funcionales, solo logging)
        terms: Set[str] = set()
        if df_coincidencias is None or df_coincidencias.empty or not cols_nombres: return terms
        cols_validas = [c for c in cols_nombres if c in df_coincidencias.columns]
        if not cols_validas: logging.warning(f"Ninguna col {cols_nombres} en coincidencias."); return terms
        for col in cols_validas:
            try: terms.update(df_coincidencias[col].dropna().astype(str).str.upper().unique())
            except Exception as e: logging.warning(f"Error extrayendo términos de '{col}': {e}")
        terms = {t for t in terms if t and not t.isspace()}
        logging.debug(f"Términos extraídos diccionario: {len(terms)} únicos.")
        return terms

    def _buscar_terminos_en_descripciones(self, df_desc: pd.DataFrame, terms: Set[str], require_all: bool = False) -> pd.DataFrame:
        # (Sin cambios funcionales, solo logging)
        cols_orig = list(df_desc.columns) if df_desc is not None else []
        if df_desc is None or df_desc.empty or not terms: return pd.DataFrame(columns=cols_orig)
        logging.info(f"Buscando {len(terms)} térms en {len(df_desc)} descrips (require_all={require_all}).")
        try:
            txt_filas = df_desc.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            terms_ok = {t for t in terms if t};
            if not terms_ok: logging.warning("No hay términos válidos."); return pd.DataFrame(columns=cols_orig)
            terms_esc = [r"\b" + re.escape(t) + r"\b" for t in terms_ok]
            if require_all: mask = txt_filas.apply(lambda txt: all(re.search(p, txt, re.I) for p in terms_esc))
            else: mask = txt_filas.str.contains('|'.join(terms_esc), regex=True, na=False, case=False)
            res = df_desc[mask]; logging.info(f"Búsqueda descrips OK. Resultados: {len(res)}."); return res
        except Exception as e: logging.exception("Error búsqueda descrips."); messagebox.showerror("Error", f"{e}"); return pd.DataFrame(columns=cols_orig)

    def _parse_numero(self, num_str: str) -> Union[int, float, None]: # (Sin cambios)
        if not isinstance(num_str, str): return None
        try: return float(num_str.replace(',', '.'))
        except ValueError: return None

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        # (Sin cambios funcionales desde la última versión con todos los operadores)
        palabras_analizadas = []; patron_comp = self.patron_comparacion_compilado
        patron_rango = self.patron_rango_compilado; patron_neg = self.patron_negacion_compilado
        for term_orig in terminos_brutos:
            term = term_orig.strip(); negate = False; item = {'original': term_orig}
            if not term: continue
            match_neg = patron_neg.match(term)
            if match_neg: negate = True; term = match_neg.group(1).strip()
            if not term: continue # Solo era '#'
            item['negate'] = negate
            match_comp = patron_comp.match(term); match_range = patron_rango.match(term)
            if match_comp:
                op, v_str = match_comp.group(1), match_comp.group(2); v_num = self._parse_numero(v_str)
                if v_num is not None: op_map = {'>':'gt', '<':'lt', '>=':'ge', '<=':'le'}; item.update({'tipo': op_map[op], 'valor': v_num})
                else: logging.warning(f"Num inválido '{v_str}' en '{term}'. Texto."); item.update({'tipo': 'str', 'valor': term})
            elif match_range:
                v1_str, v2_str = match_range.group(1), match_range.group(3); v1, v2 = self._parse_numero(v1_str), self._parse_numero(v2_str)
                if v1 is not None and v2 is not None: item.update({'tipo': 'range', 'valor': sorted([v1, v2])})
                else: logging.warning(f"Rango inválido '{term}'. Texto."); item.update({'tipo': 'str', 'valor': term})
            else: item.update({'tipo': 'str', 'valor': term})
            palabras_analizadas.append(item)
        logging.debug(f"Términos analizados: {palabras_analizadas}")
        return palabras_analizadas

    def _aplicar_mascara_diccionario(self, df: pd.DataFrame, cols_nombres: List[str], terms_analizados: List[Dict[str, Any]], op_principal: str) -> pd.Series:
        # (Lógica sin cambios, solo comentarios añadidos)
        if df is None or df.empty or not cols_nombres or not terms_analizados: return pd.Series(False, index=df.index if df is not None else None)
        cols_ok = [c for c in cols_nombres if c in df.columns];
        if not cols_ok: logging.error(f"Ninguna col {cols_nombres} válida."); return pd.Series(False, index=df.index)

        # Separar términos positivos (sin #) y negativos (con #)
        terms_pos = [item for item in terms_analizados if not item['negate']]
        terms_neg = [item for item in terms_analizados if item['negate']]

        # --- Máscara para términos positivos ---
        # Inicializamos: Si es AND, empezamos con todo True (deben cumplirse todos). Si es OR, empezamos con todo False (basta que se cumpla uno).
        mask_pos = pd.Series(op_principal.upper() == 'AND', index=df.index)

        if terms_pos:
            # Procesamos cada término positivo
            for item in terms_pos:
                mask_item = pd.Series(False, index=df.index) # Máscara para este término específico (inicia en False)
                tipo, valor = item['tipo'], item['valor']
                # Buscamos el término en CUALQUIERA de las columnas de búsqueda válidas
                for col_n in cols_ok:
                    col = df[col_n]; mask_col_item = pd.Series(False, index=df.index) # Máscara para esta columna específica
                    try:
                        # Aplicar la lógica de comparación/búsqueda según el tipo de término
                        if tipo == 'str':
                            mask_col_item = col.astype(str).str.contains(re.escape(str(valor)), case=False, na=False, regex=True)
                        elif tipo in ['gt', 'lt', 'ge', 'le']:
                            col_num = pd.to_numeric(col, errors='coerce') # Intentar convertir a número, si falla -> NaN
                            if tipo == 'gt': mask_col_item = col_num > valor
                            elif tipo == 'lt': mask_col_item = col_num < valor
                            elif tipo == 'ge': mask_col_item = col_num >= valor
                            else: mask_col_item = col_num <= valor # le
                            mask_col_item = mask_col_item.fillna(False) # Tratar NaN como no coincidencia
                        elif tipo == 'range':
                            min_v, max_v = valor
                            col_num = pd.to_numeric(col, errors='coerce')
                            mask_col_item = (col_num >= min_v) & (col_num <= max_v)
                            mask_col_item = mask_col_item.fillna(False)

                        # Combinamos la máscara de esta columna con la máscara general del término (OR)
                        # El término se cumple si coincide en AL MENOS UNA columna
                        mask_item |= mask_col_item
                    except Exception as e:
                        logging.warning(f"Error procesando término positivo '{item['original']}' en columna '{col_n}': {e}")

                # Combinamos la máscara de este término con la máscara positiva general
                if op_principal.upper() == 'AND':
                    mask_pos &= mask_item # Si es AND, todas las mask_item deben ser True
                else: # OR
                    mask_pos |= mask_item # Si es OR, basta con que una mask_item sea True
        # else: mask_pos ya está inicializada correctamente (True para AND si no hay positivos, False para OR)

        # --- Máscara para términos negativos ---
        # Queremos identificar las filas que coinciden con CUALQUIERA de los términos negados.
        mask_neg_comb = pd.Series(False, index=df.index) # Empezamos en False

        if terms_neg:
            # Procesamos cada término negativo
            for item in terms_neg:
                mask_item_neg = pd.Series(False, index=df.index) # Máscara para este término negado (inicia en False)
                tipo, valor = item['tipo'], item['valor']
                # Buscamos el término negado en CUALQUIERA de las columnas de búsqueda válidas
                for col_n in cols_ok:
                    col = df[col_n]; mask_col_item_neg = pd.Series(False, index=df.index)
                    try:
                        # Aplicar la misma lógica de comparación/búsqueda que antes
                        if tipo == 'str':
                            mask_col_item_neg = col.astype(str).str.contains(re.escape(str(valor)), case=False, na=False, regex=True)
                        elif tipo in ['gt', 'lt', 'ge', 'le']:
                            col_num = pd.to_numeric(col, errors='coerce')
                            if tipo == 'gt': mask_col_item_neg = col_num > valor
                            elif tipo == 'lt': mask_col_item_neg = col_num < valor
                            elif tipo == 'ge': mask_col_item_neg = col_num >= valor
                            else: mask_col_item_neg = col_num <= valor # le
                            mask_col_item_neg = mask_col_item_neg.fillna(False)
                        elif tipo == 'range':
                            min_v, max_v = valor
                            col_num = pd.to_numeric(col, errors='coerce')
                            mask_col_item_neg = (col_num >= min_v) & (col_num <= max_v)
                            mask_col_item_neg = mask_col_item_neg.fillna(False)

                        # Combinamos la máscara de esta columna con la máscara del término negado (OR)
                        mask_item_neg |= mask_col_item_neg
                    except Exception as e:
                        logging.warning(f"Error procesando término negativo '{item['original']}' en columna '{col_n}': {e}")

                # Combinamos la máscara de este término negado con la máscara negativa combinada (OR)
                # Marcamos una fila si coincide con CUALQUIER término negado.
                mask_neg_comb |= mask_item_neg

        # --- Combinación Final ---
        # Una fila es válida si cumple las condiciones positivas Y NO cumple NINGUNA de las condiciones negativas.
        mask_final = mask_pos & (~mask_neg_comb)

        logging.debug(f"Mask Pos: {mask_pos.sum()}, Mask Neg: {mask_neg_comb.sum()}, Mask Final: {mask_final.sum()}")
        return mask_final

    def _busqueda_simple(self, df_dic: pd.DataFrame, df_desc: pd.DataFrame, term: str) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        # (Sin cambios funcionales, usa lógica unificada)
        logging.info(f"Búsqueda Simple: '{term}'"); cols_n = self._obtener_nombres_columnas_busqueda(df_dic)
        if cols_n is None: return (df_dic, df_desc) # Devuelve originales si hay error de columnas
        terms_b = [term.strip()]; terms_a = self._analizar_terminos(terms_b)
        if not terms_a: logging.warning(f"Término simple inválido: '{term}'"); messagebox.showwarning("Inválido", f"'{term}'"); return (df_dic, df_desc) # Devuelve originales
        # En búsqueda simple, aplicamos máscara con 'OR' (aunque solo hay un término, es consistente)
        mask = self._aplicar_mascara_diccionario(df_dic, cols_n, terms_a, 'OR')
        if not mask.any(): logging.info(f"'{term}' no encontrado/negado en diccionario."); return (df_dic, df_desc) # Indicador de no encontrado
        logging.info(f"'{term}' encontrado en diccionario. Extrayendo términos..."); coinc = df_dic[mask]; terms_d = self._extraer_terminos_diccionario(coinc, cols_n)
        if not terms_d: logging.warning(f"Sin términos válidos extraídos de {len(coinc)} fila(s) coincidentes."); messagebox.showinfo("Aviso", f"Se encontraron {len(coinc)} fila(s) en el diccionario para '{term}', pero no se pudieron extraer términos válidos de ellas para buscar en descripciones."); return pd.DataFrame(columns=df_desc.columns) # Devuelve DF vacío
        logging.info(f"Buscando {len(terms_d)} términos extraídos en descripciones..."); return self._buscar_terminos_en_descripciones(df_desc, terms_d)

    def _busqueda_compuesta(self, df_dic: pd.DataFrame, df_desc: pd.DataFrame, term: str, sep: str, op: str, req_all: bool) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        # (Sin cambios funcionales, usa lógica unificada)
        logging.info(f"Búsqueda Compuesta ({op}, sep='{sep}'): '{term}'"); cols_n = self._obtener_nombres_columnas_busqueda(df_dic)
        if cols_n is None: return (df_dic, df_desc) # Devuelve originales
        terms_b = [p.strip() for p in term.split(sep) if p.strip()]
        if not terms_b: logging.warning(f"Sin términos válidos al separar '{term}' con '{sep}'."); messagebox.showwarning("Inválido", f"La entrada '{term}' no contiene términos válidos usando el separador '{sep}'."); return (df_dic, df_desc) # Devuelve originales
        terms_a = self._analizar_terminos(terms_b)
        if not terms_a: logging.warning(f"Términos inválidos tras análisis: '{term}'."); messagebox.showwarning("Inválido", f"No se pudieron analizar los términos en '{term}'. Revise la sintaxis."); return (df_dic, df_desc) # Devuelve originales
        mask = self._aplicar_mascara_diccionario(df_dic, cols_n, terms_a, op)
        if not mask.any(): logging.info(f"Combinación '{term}' ({op}) no encontrada/negada en diccionario."); return (df_dic, df_desc) # Indicador no encontrado
        logging.info(f"Combinación '{term}' ({op}) encontrada en diccionario. Extrayendo..."); coinc = df_dic[mask]; terms_d = self._extraer_terminos_diccionario(coinc, cols_n)
        if not terms_d: logging.warning(f"Sin términos válidos extraídos."); messagebox.showinfo("Aviso", f"Se encontraron {len(coinc)} fila(s) en el diccionario para '{term}' ({op}), pero sin términos válidos extraíbles."); return pd.DataFrame(columns=df_desc.columns) # Devuelve DF vacío
        logging.info(f"Buscando {len(terms_d)} términos en descripciones (require_all={req_all})..."); return self._buscar_terminos_en_descripciones(df_desc, terms_d, require_all=req_all)

    def buscar(self, term_buscar: str) -> Union[None, pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        # (Sin cambios funcionales, usa lógica unificada)
        logging.info(f"--- Nueva Búsqueda --- Término: '{term_buscar}'")
        if self.datos_diccionario is None: logging.error("Diccionario no cargado."); messagebox.showwarning("Falta Archivo", "Por favor, cargue primero el archivo Diccionario."); return None
        if self.datos_descripcion is None: logging.error("Descripciones no cargadas."); messagebox.showwarning("Falta Archivo", "Por favor, cargue primero el archivo de Descripciones."); return None

        term_proc = term_buscar.strip();
        # Si la búsqueda está vacía, devolvemos todas las descripciones
        if not term_proc:
            logging.info("Término de búsqueda vacío. Devolviendo todas las descripciones.")
            return self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame()

        # Hacemos copias para no modificar los DataFrames originales
        df_dic, df_desc = self.datos_diccionario.copy(), self.datos_descripcion.copy()

        if df_dic.empty: logging.error("El DataFrame del Diccionario está vacío."); messagebox.showerror("Error Datos", "El archivo Diccionario cargado está vacío."); return None
        if df_desc.empty: logging.error("El DataFrame de Descripciones está vacío."); messagebox.showerror("Error Datos", "El archivo de Descripciones cargado está vacío."); return pd.DataFrame(columns=df_desc.columns) # Devolvemos vacío con columnas

        try:
            # Decidimos qué tipo de búsqueda hacer según los operadores
            if '+' in term_proc:
                # Búsqueda Compuesta AND (todos los términos separados por +)
                # El último parámetro 'False' indica que no se requiere que todos los términos extraídos aparezcan en la descripción final
                return self._busqueda_compuesta(df_dic, df_desc, term_proc, '+', 'AND', False)
            elif '|' in term_proc:
                 # Búsqueda Compuesta OR (cualquier término separado por |)
                return self._busqueda_compuesta(df_dic, df_desc, term_proc, '|', 'OR', False)
            elif '/' in term_proc:
                 # Búsqueda Compuesta OR (alternativa con /)
                return self._busqueda_compuesta(df_dic, df_desc, term_proc, '/', 'OR', False)
            else:
                # Búsqueda Simple (un solo término o condición)
                return self._busqueda_simple(df_dic, df_desc, term_proc)
        except Exception as e:
            logging.exception("Error durante la orquestación de la búsqueda.")
            messagebox.showerror("Error Inesperado", f"Ocurrió un error al procesar la búsqueda:\n{e}")
            return None # Indicador de error

    def buscar_en_descripciones_directo(self, term_buscar: str) -> pd.DataFrame:
        # (Sin cambios funcionales, solo separador '/')
        logging.info(f"Búsqueda Directa en Descripciones: '{term_buscar}'");
        if self.datos_descripcion is None or self.datos_descripcion.empty:
            logging.warning("Intento de búsqueda directa sin descripciones cargadas o vacías.")
            messagebox.showwarning("Faltan Datos", "No hay datos de descripciones cargados para realizar la búsqueda directa.")
            return pd.DataFrame() # Devuelve DataFrame vacío

        term_limpio = term_buscar.strip().upper();
        if not term_limpio:
             # Si el término está vacío después de limpiar, devolvemos todo
            logging.info("Término de búsqueda directa vacío. Devolviendo todas las descripciones.")
            return self.datos_descripcion.copy()

        df_desc = self.datos_descripcion.copy();
        res = pd.DataFrame(columns=df_desc.columns) # Preparamos un resultado vacío por si acaso
        try:
            # Concatenamos todas las columnas de texto por fila para buscar en todo el texto
            txt_filas = df_desc.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            mask = pd.Series(False, index=df_desc.index) # Máscara inicial en False

            if '+' in term_limpio: # Lógica AND
                palabras = [p.strip() for p in term_limpio.split('+') if p.strip()]
                if not palabras: return res # Si no hay palabras válidas, devuelve vacío
                mask = pd.Series(True, index=df_desc.index) # Empezamos en True para AND
                for p in palabras:
                    # Usamos \b para buscar palabras completas
                    mask &= txt_filas.str.contains(r"\b"+re.escape(p)+r"\b", regex=True, na=False)
            elif '|' in term_limpio or '/' in term_limpio: # Lógica OR
                sep = '|' if '|' in term_limpio else '/'
                palabras = [p.strip() for p in term_limpio.split(sep) if p.strip()]
                if not palabras: return res # Si no hay palabras válidas, devuelve vacío
                # Empezamos en False para OR (ya está inicializada así)
                for p in palabras:
                    mask |= txt_filas.str.contains(r"\b"+re.escape(p)+r"\b", regex=True, na=False)
            else: # Búsqueda simple (un solo término)
                mask = txt_filas.str.contains(r"\b"+re.escape(term_limpio)+r"\b", regex=True, na=False)

            res = df_desc[mask]; # Aplicamos la máscara final
            logging.info(f"Búsqueda directa completada. Resultados: {len(res)}.")
        except Exception as e:
            logging.exception("Error durante la búsqueda directa en descripciones.")
            messagebox.showerror("Error Búsqueda Directa", f"Ocurrió un error:\n{e}")
            return pd.DataFrame(columns=df_desc.columns) # Devuelve vacío en caso de error
        return res

# --- Clase ExtractorMagnitud (Sin cambios) ---
# (Código omitido por brevedad, es idéntico al anterior)
class ExtractorMagnitud:
    MAGNITUDES_PREDEFINIDAS: List[str] = [
        "A","AMP","AMPS","AH","ANTENNA","BASE","BIT","ETH","FE","G","GB",
        "GBE","GE","GIGABIT","GBASE","GBASEWAN","GBIC","GBIT","GBPS","GH",
        "GHZ","HZ","KHZ","KM","KVA","KW","LINEAS","LINES","MHZ","NM","PORT",
        "PORTS","PTOS","PUERTO","PUERTOS","P","V","VA","VAC","VC","VCC",
        "VCD","VDC","W","WATTS","E","FE","GBE","GE","POTS","STM"
    ]
    def __init__(self, magnitudes: Optional[List[str]] = None):
        self.magnitudes = magnitudes if magnitudes is not None else self.MAGNITUDES_PREDEFINIDAS

    @staticmethod
    def _quitar_diacronicos_y_acentos(texto: str) -> str:
        if not isinstance(texto, str) or not texto: return ""
        try:
            # Normaliza a 'NFKD' (descompone caracteres acentuados) y filtra los diacríticos (combinantes)
            forma_normalizada = unicodedata.normalize('NFKD', texto)
            return ''.join(c for c in forma_normalizada if not unicodedata.combining(c))
        except TypeError: return "" # Por si acaso entra algo que no sea string

    def buscar_cantidad_para_magnitud(self, mag: str, descripcion: str) -> Optional[str]:
        if not isinstance(mag, str) or not mag: return None
        if not isinstance(descripcion, str) or not descripcion: return None

        mag_upper = mag.upper() # Convertimos magnitud a mayúsculas
        texto_limpio = self._quitar_diacronicos_y_acentos(descripcion.upper()) # Limpiamos y convertimos descripción
        if not texto_limpio: return None

        mag_escapada = re.escape(mag_upper) # Escapamos la magnitud para usarla en regex

        # Patrón: Busca un número (entero o decimal con . o ,), seguido opcionalmente por espacio o 'X',
        # seguido por la magnitud exacta (\b para palabra completa), y que no le sigan letras/números.
        patron_principal = re.compile(
            r"(\d+([.,]\d+)?)[ X]{0,1}(\b" + mag_escapada + r"\b)(?![a-zA-Z0-9])"
        )

        # Buscamos todas las coincidencias y devolvemos la primera cantidad encontrada
        for match in patron_principal.finditer(texto_limpio):
            return match.group(1).strip() # group(1) captura el número (\d+([.,]\d+)?)

        return None # Si no se encuentra ninguna coincidencia


# --- Clase Interfaz Gráfica ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE = "config_buscador.json"

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado")
        self.geometry("1200x800") # Un poco más ancho para nuevo botón

        self.config = self._cargar_configuracion()
        indices_cfg = self.config.get("indices_columnas_busqueda_dic", [0, 3]) # Default aquí
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg)
        self.extractor_magnitud = ExtractorMagnitud()
        self.resultados_actuales: Optional[pd.DataFrame] = None
        # >>> INICIO: Atributo para guardar último término buscado (Req 1: Trazabilidad) <<<
        self.ultimo_termino_buscado: Optional[str] = None
        # <<< FIN: Atributo para guardar último término buscado >>>
        self.color_fila_par = "white"; self.color_fila_impar = "#f0f0f0"

        self._configurar_estilo_ttk()
        self._crear_widgets()
        self._configurar_grid()
        self._configurar_eventos()
        self._configurar_tags_treeview()
        self._configurar_orden_tabla(self.tabla_resultados) # Configurar orden inicial (aunque vacía)
        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones.")
        logging.info("Interfaz Gráfica inicializada.")

    def _cargar_configuracion(self) -> Dict: # (Sin cambios)
        config = {}
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f: config = json.load(f)
                logging.info(f"Configuración cargada desde: {self.CONFIG_FILE}")
            except Exception as e:
                logging.error(f"Error al cargar archivo de configuración '{self.CONFIG_FILE}': {e}")
                messagebox.showwarning("Error Configuración", f"No se pudo cargar la configuración:\n{e}")
        else:
            logging.info("Archivo de configuración no encontrado. Se creará uno nuevo al cerrar.")
        # Asegurarse de que las claves esperadas existan, con valores por defecto si no
        config.setdefault("last_dic_path", None)
        config.setdefault("last_desc_path", None)
        config.setdefault("indices_columnas_busqueda_dic", [0, 3]) # Default si falta
        return config

    def _guardar_configuracion(self): # (Sin cambios)
        # Actualizar el diccionario de configuración con los valores actuales
        self.config["last_dic_path"] = self.motor.archivo_diccionario_actual
        self.config["last_desc_path"] = self.motor.archivo_descripcion_actual
        # Guardar también los índices usados (aunque se cargan al inicio, por consistencia)
        self.config["indices_columnas_busqueda_dic"] = self.motor.indices_columnas_busqueda_dic
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4) # indent=4 para que sea legible
            logging.info(f"Configuración guardada en: {self.CONFIG_FILE}")
        except Exception as e:
            logging.error(f"Error al guardar la configuración en '{self.CONFIG_FILE}': {e}")
            messagebox.showerror("Error Configuración", f"No se pudo guardar la configuración:\n{e}")

    def _configurar_estilo_ttk(self): # (Sin cambios)
        style = ttk.Style(self); themes = style.theme_names(); os_name = platform.system()
        # Preferencias de temas por sistema operativo
        prefs = {"Windows":["vista","xpnative","clam"],"Darwin":["aqua","clam"],"Linux":["clam","alt","default"]}
        # Buscar el primer tema preferido que esté disponible
        theme_to_use = next((t for t in prefs.get(os_name, ["clam","default"]) if t in themes), None)
        # Si no se encontró ninguno preferido, usar el actual o el default o el primero disponible
        if not theme_to_use:
            theme_to_use = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))

        if theme_to_use:
            logging.info(f"Aplicando tema TTK: {theme_to_use}")
            try:
                style.theme_use(theme_to_use)
            except tk.TclError as e:
                logging.warning(f"No se pudo aplicar el tema '{theme_to_use}': {e}. Usando tema por defecto.")
        else:
            logging.warning("No se encontró ningún tema TTK disponible.")

    def _crear_widgets(self):
        # --- Marco Principal de Controles ---
        self.marco_controles = ttk.LabelFrame(self, text="Controles")

        # --- Sección Carga de Archivos ---
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.lbl_dic_cargado = ttk.Label(self.marco_controles, text="Dic: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1) # Borde para destacar
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion, state="disabled") # Deshabilitado al inicio
        self.lbl_desc_cargado = ttk.Label(self.marco_controles, text="Desc: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1) # Borde para destacar

        # --- Sección Búsqueda ---
        # Marco para los botones de operadores
        self.frame_ops = ttk.Frame(self.marco_controles)

        # Botones de operadores para insertar en la entrada de búsqueda
        self.btn_and = ttk.Button(self.frame_ops, text="+", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, "+"))
        self.btn_or = ttk.Button(self.frame_ops, text="|", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, "|"))
        self.btn_not = ttk.Button(self.frame_ops, text="#", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, "#"))
        self.btn_gt = ttk.Button(self.frame_ops, text=">", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, ">"))
        self.btn_lt = ttk.Button(self.frame_ops, text="<", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, "<"))
        self.btn_ge = ttk.Button(self.frame_ops, text="≥", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, ">=")) # Usar >=
        self.btn_le = ttk.Button(self.frame_ops, text="≤", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, "<=")) # Usar <=
        self.btn_range = ttk.Button(self.frame_ops, text="-", width=3, command=lambda: self.entrada_busqueda.insert(tk.INSERT, "-")) # Para rangos n1-n2

        # Organizar botones de operadores horizontalmente dentro de su frame
        self.btn_and.pack(side=tk.LEFT, padx=2); self.btn_or.pack(side=tk.LEFT, padx=2)
        self.btn_not.pack(side=tk.LEFT, padx=2); self.btn_gt.pack(side=tk.LEFT, padx=2)
        self.btn_lt.pack(side=tk.LEFT, padx=2); self.btn_ge.pack(side=tk.LEFT, padx=2)
        self.btn_le.pack(side=tk.LEFT, padx=2); self.btn_range.pack(side=tk.LEFT, padx=2)

        # Entrada de texto para la búsqueda
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50)
        # >>> INICIO: Botón Buscar en Vista Previa (COMENTADO - ver explicación) <<<
        # Si quieres activarlo, descomenta la línea de abajo y la correspondiente en _configurar_grid
        # self.btn_buscar_en_preview = ttk.Button(self.marco_controles, text="Ir a", command=self._buscar_y_enfocar_en_preview, state="disabled", width=5)
        # <<< FIN: Botón Buscar en Vista Previa >>>

        # Botones principales de acción
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda, state="disabled") # Deshabilitado al inicio
        self.btn_ayuda = ttk.Button(self.marco_controles, text="?", command=self._mostrar_ayuda, width=3) # Botón de ayuda corto
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar", command=self._exportar_resultados, state="disabled") # Deshabilitado al inicio

        # --- Sección Tablas y Barras de Scroll ---
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario:")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados / Descripciones:")

        # Frame y Treeview para la vista previa del diccionario
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings", height=8) # Limitar altura preview
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)

        # Frame y Treeview para los resultados / descripciones
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings")
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)

        # --- Barra de Estado ---
        self.barra_estado = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W, borderwidth=1) # Borde para destacar

        # Actualizar etiquetas de archivos cargados con la info inicial (probablemente ninguno)
        self._actualizar_etiquetas_archivos()

    def _configurar_tags_treeview(self): # (Sin cambios)
        # Configura colores alternos para las filas de ambas tablas
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par)
            tabla.tag_configure('impar', background=self.color_fila_impar)
            # Podrías añadir más tags si quisieras resaltar filas de otra forma

    def _configurar_grid(self):
        # --- Configuración del Grid de la Ventana Principal ---
        self.grid_rowconfigure(2, weight=1) # Fila para preview diccionario (peso 1)
        self.grid_rowconfigure(4, weight=3) # Fila para resultados (peso 3, más espacio)
        self.grid_columnconfigure(0, weight=1) # Columna única se expande horizontalmente

        # --- Posicionamiento del Marco de Controles ---
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5)) # Arriba, pegado al norte y este/oeste

        # --- Configuración del Grid DENTRO del Marco de Controles ---
        # Fila 0: Carga de Archivos
        # Hacemos que las etiquetas de archivo se expandan un poco si hay espacio
        self.marco_controles.grid_columnconfigure(1, weight=1)
        self.marco_controles.grid_columnconfigure(3, weight=1)
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew") # Ocupa 1 columna ahora

        # Fila 1: Botones de Operadores
        self.frame_ops.grid(row=1, column=0, columnspan=4, padx=5, pady=(5,0), sticky="w") # Ocupa el ancho de los 4 primeros elementos de la fila 2

        # Fila 2: Búsqueda y Acciones
        # Hacemos que la entrada de búsqueda se expanda
        self.marco_controles.grid_columnconfigure(1, weight=1) # Columna de la entrada
        # >>> INICIO: ELIMINADA la línea que causaba error <<<
        # Ya no intentamos colocar self.lbl_busqueda porque no existe
        # <<< FIN: ELIMINADA la línea que causaba error >>>
        self.entrada_busqueda.grid(row=2, column=0, columnspan=2, padx=5, pady=(0,5), sticky="ew") # Ocupa 2 columnas ahora
        # >>> INICIO: Posición del botón "Ir a" (COMENTADO) <<<
        # Si descomentas el botón en _crear_widgets, descomenta esta línea también
        # self.btn_buscar_en_preview.grid(row=2, column=2, padx=(5,0), pady=(0,5), sticky="w")
        # <<< FIN: Posición del botón "Ir a" >>>
        # Ajustar las columnas de los botones restantes
        self.btn_buscar.grid(row=2, column=2, padx=(2,0), pady=(0,5), sticky="w") # Ahora columna 2
        self.btn_ayuda.grid(row=2, column=3, padx=(2,5), pady=(0,5), sticky="w")  # Ahora columna 3
        self.btn_exportar.grid(row=2, column=4, padx=(10, 5), pady=(0,5), sticky="e") # Columna 4, pegado a la derecha

        # --- Posicionamiento de Etiquetas de Tablas ---
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0)) # Encima de su tabla
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))  # Encima de su tabla

        # --- Posicionamiento de Frames y Tablas (Grid interno sin cambios) ---
        # Frame Diccionario
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1); self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew"); self.scrolly_diccionario.grid(row=0, column=1, sticky="ns"); self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")
        # Frame Resultados
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1); self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew"); self.scrolly_resultados.grid(row=0, column=1, sticky="ns"); self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # --- Posicionamiento de la Barra de Estado ---
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0)) # Abajo del todo, ocupando el ancho

    def _configurar_eventos(self): # (Sin cambios)
        # Asociar Enter en la entrada de búsqueda con la función de buscar
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())
        # Asociar el cierre de la ventana con la función on_closing
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _actualizar_estado(self, mensaje: str): # (Sin cambios)
        # Actualiza el texto de la barra de estado inferior
        self.barra_estado.config(text=mensaje)
        logging.info(f"Estado UI: {mensaje}") # También lo mandamos al log
        self.update_idletasks() # Forzar actualización de la UI

    def _mostrar_ayuda(self): # (Sin cambios)
        # Muestra una ventana emergente con la ayuda de sintaxis
        ayuda = """Sintaxis de Búsqueda en Diccionario:
-------------------------------------
- Texto simple: Busca la palabra o frase exacta (insensible a mayús/minús).
  Ej: `router cisco`
- `término1 + término2`: Busca filas que contengan AMBOS términos (AND).
  Ej: `tarjeta + 16 puertos`
- `término1 | término2`: Busca filas que contengan AL MENOS UNO de los términos (OR).
  Ej: `modulo | SFP`
- `término1 / término2`: Alternativa para OR.
  Ej: `switch / conmutador`
- Comparaciones numéricas (aplican a columnas configuradas si son numéricas):
  - `>numero`: Mayor que. Ej: `>1000`
  - `<numero`: Menor que. Ej: `<50`
  - `>=numero`: Mayor o igual que. Ej: `>=48`
  - `<=numero`: Menor o igual que. Ej: `<=10.5`
- Rangos numéricos (ambos incluidos):
  - `num1-num2`: Entre num1 y num2. Ej: `10-20` (buscará 10, 11, ..., 20)
- Negación (excluir filas):
  - `#término`: Excluye filas que coincidan con 'término'.
    'término' puede ser texto, comparación o rango.
  Ej: `switch + #gestionable` (busca 'switch' pero no los que contengan 'gestionable')
  Ej: `tarjeta + #>8` (busca 'tarjeta' pero no las que tengan número > 8)

Notas:
- La negación (#) se aplica al final.
- Las comparaciones y rangos intentan convertir el texto de las celdas a número.
- La búsqueda es insensible a mayúsculas/minúsculas y acentos.

Búsqueda Directa (si el término no está en diccionario):
------------------------------------------------------
Se busca directamente en las descripciones:
- `texto`: Busca el texto.
- `t1 + t2`: Busca descripciones con AMBOS términos.
- `t1 | t2` o `t1 / t2`: Busca descripciones con AL MENOS UNO.
"""
        messagebox.showinfo("Ayuda - Sintaxis de Búsqueda", ayuda)

    def _configurar_orden_tabla(self, tabla: ttk.Treeview): # (Sin cambios)
        # Permite ordenar la tabla haciendo clic en las cabeceras de columna
        cols = tabla["columns"]
        if cols: # Solo si la tabla tiene columnas definidas
            for col in cols:
                # Asigna un comando a cada cabecera para llamar a _ordenar_columna
                # Usamos lambda para capturar el nombre de la columna (c=col) en el momento
                tabla.heading(col, text=col, anchor=tk.W,
                              command=lambda c=col: self._ordenar_columna(tabla, c, False))

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool): # (Sin cambios)
        # Lógica para ordenar el DataFrame subyacente y actualizar la tabla
        # Solo ordena la tabla de resultados (self.resultados_actuales)
        if tabla != self.tabla_resultados or self.resultados_actuales is None or self.resultados_actuales.empty:
            logging.debug(f"Intento de ordenar tabla vacía o no aplicable ({'Diccionario' if tabla == self.tabla_diccionario else 'Resultados'}).")
            return # No hacer nada si no hay resultados

        logging.info(f"Ordenando resultados por columna '{col}', descendente={reverse}")
        try:
            # Intentamos ordenar numéricamente si es posible, si no, alfabéticamente
            # `pd.to_numeric` con errors='coerce' convierte no números a NaN, que `sort_values` maneja
            # `na_position='last'` pone los valores no numéricos (o vacíos) al final
            df_ordenado = self.resultados_actuales.sort_values(
                by=col,
                ascending=not reverse, # ascending es lo opuesto a reverse
                na_position='last',
                key=lambda x: pd.to_numeric(x, errors='coerce') # Clave para intentar orden numérico
            )
            self.resultados_actuales = df_ordenado # Guardamos el DataFrame ordenado
            self._actualizar_tabla(tabla, self.resultados_actuales) # Redibujar la tabla con los datos ordenados
            # Actualizar el comando de la cabecera para que la próxima vez ordene al revés
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, not reverse))
            self._actualizar_estado(f"Resultados ordenados por '{col}' ({'Asc' if not reverse else 'Desc'}).")
        except Exception as e:
            logging.exception(f"Error al intentar ordenar por columna '{col}'")
            messagebox.showerror("Error al Ordenar", f"No se pudo ordenar por '{col}':\n{e}")
            # Restaurar el comando original por si acaso
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, False))

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None, columnas_a_mostrar: Optional[List[str]] = None):
        # (Sin cambios funcionales importantes, quizás logging añadido)
        logging.debug(f"Actualizando tabla {'Diccionario' if tabla == self.tabla_diccionario else 'Resultados'}.")
        # --- Limpiar tabla ---
        try:
            # Borrar todas las filas existentes
            for i in tabla.get_children(): tabla.delete(i)
        except tk.TclError as e:
            # A veces puede dar error si la tabla ya está destruida o en mal estado
            logging.warning(f"Error Tcl al limpiar tabla: {e}")
            pass # Intentar continuar
        # Resetear columnas a vacío
        tabla["columns"] = ()

        # --- Comprobar si hay datos ---
        if datos is None or datos.empty:
            logging.debug("No hay datos para mostrar en la tabla.")
            # Podríamos poner un mensaje en la tabla, pero la dejamos vacía
            # tabla.insert("", "end", values=["(Sin datos)"], tags=('impar',))
            return

        # --- Determinar columnas a mostrar ---
        if columnas_a_mostrar:
            # Si se especifican columnas, usar solo las que existan en el DataFrame
            cols_ok = [c for c in columnas_a_mostrar if c in datos.columns]
            if not cols_ok:
                logging.warning(f"Columnas especificadas {columnas_a_mostrar} no encontradas. Mostrando todas.")
                cols_ok = list(datos.columns) # Fallback a todas
        else:
            # Si no se especifican, mostrar todas
            cols_ok = list(datos.columns)

        df_mostrar = datos[cols_ok] # DataFrame con solo las columnas seleccionadas

        # --- Configurar columnas en Treeview ---
        tabla["columns"] = cols_ok
        for col in cols_ok:
            # Configurar cabecera (texto y alineación)
            tabla.heading(col, text=str(col), anchor=tk.W)
            # Configurar columna (alineación y ancho inicial)
            tabla.column(col, anchor=tk.W)
            # Intentar calcular un ancho razonable basado en el contenido y el nombre de la cabecera
            try:
                # Tomar una muestra (ej. 100 filas) para calcular el ancho
                muestra = min(len(df_mostrar), 100)
                sub_df = df_mostrar.iloc[:muestra][col].dropna().astype(str)
                # Ancho del contenido más largo en la muestra
                ancho_contenido = sub_df.str.len().max() if not sub_df.empty else 0
                # Ancho del nombre de la cabecera
                ancho_cabecera = len(str(col))
                # Calcular ancho final (ajustar multiplicadores si es necesario)
                ancho = max(70, min(int(max(ancho_cabecera * 9, ancho_contenido * 7) + 20), 400)) # Mín 70, Máx 400
                tabla.column(col, width=ancho, minwidth=70) # Aplicar ancho calculado
            except Exception as e:
                # Si falla el cálculo, poner un ancho por defecto
                logging.warning(f"Error calculando ancho para columna '{col}': {e}. Usando ancho por defecto.")
                tabla.column(col, width=100, minwidth=50)

        # --- Insertar filas ---
        # Limitar el número de filas si se especificó un límite
        if limite_filas is not None and len(df_mostrar) > limite_filas:
            df_final = df_mostrar.head(limite_filas)
            logging.debug(f"Mostrando las primeras {limite_filas} de {len(df_mostrar)} filas.")
        else:
            df_final = df_mostrar
            logging.debug(f"Mostrando {len(df_final)} filas.")

        # Iterar e insertar cada fila
        for i, (_, row) in enumerate(df_final.iterrows()):
            # Convertir todos los valores a string (manejar NaNs como "")
            vals = [str(v) if pd.notna(v) else "" for v in row.values]
            # Asignar tag 'par' o 'impar' para colores alternos
            tag = 'par' if i % 2 == 0 else 'impar'
            try:
                # Insertar la fila con sus valores y tag
                tabla.insert("", "end", values=vals, tags=(tag,))
            except tk.TclError as e:
                # Fallback por si hay caracteres problemáticos incluso después de str()
                logging.warning(f"Error Tcl insertando fila {i}: {e}. Intentando con ASCII.")
                try:
                    # Intentar convertir a ASCII ignorando errores
                    vals_ascii = [v.encode('ascii', 'ignore').decode('ascii') for v in vals]
                    tabla.insert("", "end", values=vals_ascii, tags=(tag,))
                except Exception as e_inner:
                    logging.error(f"Fallo el fallback ASCII para fila {i}: {e_inner}")
                    # Omitir esta fila si falla de nuevo

        # --- Reconfigurar ordenación (solo si es la tabla de resultados) ---
        if tabla == self.tabla_resultados:
            self._configurar_orden_tabla(tabla) # Para que las nuevas columnas sean ordenables

    def _actualizar_etiquetas_archivos(self):
        """Actualiza las etiquetas que muestran los nombres de los archivos cargados."""
        dic_name = os.path.basename(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else "Ninguno"
        desc_name = os.path.basename(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else "Ninguno"

        # Truncar nombres largos para que quepan mejor
        max_len_label = 25
        dic_display = f"Dic: {dic_name}" if len(dic_name) <= max_len_label else f"Dic: ...{dic_name[-(max_len_label-4):]}"
        desc_display = f"Desc: {desc_name}" if len(desc_name) <= max_len_label else f"Desc: ...{desc_name[-(max_len_label-4):]}"

        self.lbl_dic_cargado.config(text=dic_display)
        self.lbl_desc_cargado.config(text=desc_display)

        # Podríamos añadir tooltips (mensajes emergentes) con la ruta completa si quisiéramos
        # Ejemplo (necesitaría una clase Tooltip o librería externa):
        # Tooltip(self.lbl_dic_cargado, text=self.motor.archivo_diccionario_actual or "Ningún diccionario cargado")
        # Tooltip(self.lbl_desc_cargado, text=self.motor.archivo_descripcion_actual or "Ninguna descripción cargada")

    def _cargar_diccionario(self):
        # (Usa config, guarda config, actualiza UI)
        # Obtener directorio del último archivo cargado desde la config
        last_dir = os.path.dirname(self.config.get("last_dic_path", "") or "") or None
        ruta = filedialog.askopenfilename(
            title="Seleccionar Archivo Diccionario",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
            initialdir=last_dir # Empezar en el último directorio usado
        )
        if not ruta: logging.info("Carga de diccionario cancelada por el usuario."); return # Si el usuario cancela

        nombre_archivo = os.path.basename(ruta)
        self._actualizar_estado(f"Cargando diccionario: {nombre_archivo}...")
        # Limpiar vista previa y deshabilitar botones dependientes
        self._actualizar_tabla(self.tabla_diccionario, None)
        self.btn_cargar_descripciones["state"] = "disabled"
        self.btn_buscar["state"] = "disabled"
        self.btn_exportar["state"] = "disabled"
        # Deshabilitar botón preview si existe
        # if hasattr(self, 'btn_buscar_en_preview'): self.btn_buscar_en_preview["state"] = "disabled"

        # Intentar cargar usando el motor
        if self.motor.cargar_excel_diccionario(ruta):
            # Si la carga fue exitosa (y pasó la validación de columnas)
            self._guardar_configuracion() # Guardar la nueva ruta en config
            df_dic = self.motor.datos_diccionario
            if df_dic is not None: # Doble chequeo por si acaso
                num_filas = len(df_dic)
                self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado. Actualizando vista previa...")

                # Actualizar etiqueta de la tabla de diccionario con info de columnas/índices
                cols_busqueda_nombres = self.motor._obtener_nombres_columnas_busqueda(df_dic)
                indices_str = ', '.join(map(str, self.motor.indices_columnas_busqueda_dic))
                lbl_text = f"Vista Previa Diccionario (Índices usados: {indices_str})"
                if cols_busqueda_nombres:
                     lbl_text = f"Vista Previa Dic ({', '.join(cols_busqueda_nombres)} - Índices: {indices_str})"
                self.lbl_tabla_diccionario.config(text=lbl_text)

                # Actualizar la tabla de vista previa (con límite de filas)
                self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols_busqueda_nombres)

                # Actualizar título de la ventana
                self.title(f"Buscador - Dic: {nombre_archivo}")
                # Habilitar botón para cargar descripciones
                self.btn_cargar_descripciones["state"] = "normal"
                # Habilitar botón de preview si existe
                # if hasattr(self, 'btn_buscar_en_preview'): self.btn_buscar_en_preview["state"] = "normal"
                # Habilitar botón de búsqueda si las descripciones ya estaban cargadas
                if self.motor.datos_descripcion is not None:
                    self.btn_buscar["state"] = "normal"
                self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado y listo.")
            # else: # El motor ya debería haber mostrado un error si df_dic es None aquí
        else:
            # La carga falló (el motor ya mostró mensaje de error)
            self._actualizar_estado("Error al cargar el diccionario. Verifique el archivo y los logs.")
            self.title("Buscador Avanzado") # Resetear título

        # Actualizar las etiquetas de nombre de archivo en la UI
        self._actualizar_etiquetas_archivos()

    def _cargar_excel_descripcion(self):
        # (Usa config, guarda config, actualiza UI)
        last_dir = os.path.dirname(self.config.get("last_desc_path", "") or "") or None
        ruta = filedialog.askopenfilename(
            title="Seleccionar Archivo de Descripciones",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
            initialdir=last_dir
        )
        if not ruta: logging.info("Carga de descripciones cancelada."); return

        nombre_archivo = os.path.basename(ruta)
        self._actualizar_estado(f"Cargando descripciones: {nombre_archivo}...")
        # Limpiar tabla de resultados, resetear resultados guardados y último término
        self._actualizar_tabla(self.tabla_resultados, None)
        self.resultados_actuales = None
        self.ultimo_termino_buscado = None
        # Deshabilitar botones dependientes
        self.btn_buscar["state"] = "disabled"
        self.btn_exportar["state"] = "disabled"

        # Intentar cargar usando el motor
        if self.motor.cargar_excel_descripcion(ruta):
            self._guardar_configuracion() # Guardar nueva ruta
            df_desc = self.motor.datos_descripcion
            if df_desc is not None:
                num_filas = len(df_desc)
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando datos...")
                # Mostrar TODAS las descripciones cargadas en la tabla de resultados
                self._actualizar_tabla(self.tabla_resultados, df_desc)
                # Guardar una copia para futuras búsquedas y exportación
                self.resultados_actuales = df_desc.copy()
                # Habilitar exportación (se pueden exportar todas las descripciones)
                self.btn_exportar["state"] = "normal"
                # Habilitar búsqueda si el diccionario ya está cargado
                if self.motor.datos_diccionario is not None:
                    self.btn_buscar["state"] = "normal"

                # Actualizar título de la ventana con ambos nombres
                if self.motor.archivo_diccionario_actual:
                    dic_n = os.path.basename(self.motor.archivo_diccionario_actual)
                    desc_n = nombre_archivo
                    self.title(f"Buscador - Dic: {dic_n} | Desc: {desc_n}")
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas y listas.")
            # else: error ya manejado por el motor
        else:
            # Carga falló
            self._actualizar_estado("Error al cargar las descripciones. Verifique el archivo y los logs.")
            # No cambiamos el título si falla

        # Actualizar etiquetas de nombre de archivo
        self._actualizar_etiquetas_archivos()

    # >>> INICIO: Método para Buscar y Enfocar en Preview (Req 2 - Recordatorio: Botón está COMENTADO) <<<
    def _buscar_y_enfocar_en_preview(self):
        """Busca el término de entrada en la tabla de vista previa del diccionario y hace scroll/selecciona la primera coincidencia."""
        termino_buscar = self.entrada_busqueda.get().strip()
        if not termino_buscar:
            messagebox.showinfo("Término Vacío", "Ingrese un término en la barra de búsqueda para buscar en la vista previa.")
            return

        # Verificar que la tabla de diccionario tenga items (filas)
        items_preview = self.tabla_diccionario.get_children('') # Obtiene los IDs de los items (filas)
        if not items_preview: # Chequeo más robusto que mirar self.motor.datos_diccionario
            messagebox.showwarning("Vista Previa Vacía", "La vista previa del diccionario está vacía. Cargue un diccionario válido.")
            return

        # Búsqueda insensible a mayúsculas/minúsculas
        termino_upper = termino_buscar.upper()
        logging.info(f"Buscando '{termino_buscar}' en la vista previa del diccionario...")
        self._actualizar_estado(f"Buscando '{termino_buscar}' en vista previa...")

        # Obtener los índices de las columnas *visibles* en la preview
        # Necesitamos esto porque `item['values']` devuelve una tupla basada en el orden de `tabla["columns"]`
        try:
             # Corrección: Asegurarse que `self.tabla_diccionario["columns"]` no esté vacío antes de iterar
            column_ids = self.tabla_diccionario["columns"]
            if not column_ids:
                 logging.error("La tabla de diccionario no tiene columnas definidas.")
                 messagebox.showerror("Error Interno", "La tabla de vista previa no tiene columnas.")
                 return

            # Creamos un diccionario {nombre_columna: indice_en_tupla_values}
            # Usamos tabla.heading(col_id, 'text') por si el nombre de columna tiene espacios o caracteres raros
            col_indices_preview = {self.tabla_diccionario.heading(col_id, 'text'): i for i, col_id in enumerate(column_ids)}

        except Exception as e:
             logging.error(f"Error obteniendo índices de columnas de la preview: {e}")
             messagebox.showerror("Error Interno", "No se pudieron obtener las columnas de la vista previa.")
             return


        found_item_id = None
        # Iterar sobre los IDs de las filas en la tabla de preview
        for item_id in items_preview:
            try:
                # Obtener los valores de la fila actual como una tupla de strings
                valores_fila = self.tabla_diccionario.item(item_id, 'values')
                # Buscar el término (en mayúsculas) dentro de CUALQUIERA de los valores de la fila
                # Convertimos cada valor a string por si acaso (aunque Treeview suele devolver strings)
                if any(termino_upper in str(val).upper() for val in valores_fila):
                    found_item_id = item_id
                    break # Detenerse en la primera coincidencia encontrada
            except Exception as e:
                # Error procesando una fila específica, registrar y continuar
                logging.warning(f"Error al procesar item {item_id} en la vista previa: {e}")
                continue

        # --- Acciones después de buscar ---
        if found_item_id:
            # ¡Encontrado!
            logging.info(f"Término '{termino_buscar}' encontrado en preview (item ID: {found_item_id}). Enfocando...")
            try:
                # Deseleccionar cualquier fila previamente seleccionada para claridad
                current_selection = self.tabla_diccionario.selection()
                if current_selection:
                    self.tabla_diccionario.selection_remove(current_selection)

                # Seleccionar la fila encontrada
                self.tabla_diccionario.selection_set(found_item_id)
                # Hacer scroll para que la fila sea visible
                self.tabla_diccionario.see(found_item_id)
                # Poner el foco en la fila (opcional, a veces puede ser molesto)
                # self.tabla_diccionario.focus(found_item_id)
                self._actualizar_estado(f"'{termino_buscar}' encontrado y enfocado en la vista previa.")
                # Devolver el foco a la entrada de búsqueda para que el usuario pueda seguir escribiendo
                self.entrada_busqueda.focus_set()
                self.entrada_busqueda.selection_range(0, tk.END) # Seleccionar texto actual para fácil reemplazo

            except tk.TclError as e:
                # Error específico de Tkinter al intentar interactuar con el item
                logging.error(f"Error Tcl al intentar enfocar item {found_item_id}: {e}")
                self._actualizar_estado(f"Error al intentar enfocar '{termino_buscar}'.")
            except Exception as e:
                 # Otro error inesperado
                 logging.error(f"Error inesperado al enfocar item {found_item_id}: {e}")
                 self._actualizar_estado(f"Error inesperado al enfocar '{termino_buscar}'.")
        else:
            # No encontrado
            logging.info(f"Término '{termino_buscar}' no encontrado en las filas visibles de la vista previa.")
            messagebox.showinfo("No Encontrado", f"El término '{termino_buscar}' no se encontró en las filas actuales de la vista previa.")
            self._actualizar_estado(f"'{termino_buscar}' no encontrado en la vista previa.")
            # Devolver foco igualmente
            self.entrada_busqueda.focus_set()
            self.entrada_busqueda.selection_range(0, tk.END)
    # <<< FIN: Método para Buscar y Enfocar en Preview >>>

    def _ejecutar_busqueda(self):
        # (Guarda el término buscado para usarlo en exportación)
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Faltan Archivos", "Asegúrese de haber cargado tanto el archivo Diccionario como el de Descripciones antes de buscar.")
            return

        termino = self.entrada_busqueda.get() # Obtener texto del campo de búsqueda

        # >>> INICIO: Guardar término para exportación (Req 1 - Sin cambios) <<<
        # Guardamos el término (si no está vacío) para usarlo luego en el nombre del archivo exportado
        self.ultimo_termino_buscado = termino if termino.strip() else None
        # <<< FIN: Guardar término para exportación >>>

        # --- Manejo de Búsqueda Vacía ---
        if not termino.strip(): # Si el término está vacío o solo espacios
            logging.info("Búsqueda vacía detectada. Mostrando todas las descripciones.")
            df_desc = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc)
            self.resultados_actuales = df_desc.copy() if df_desc is not None else None
            num_filas = len(df_desc) if df_desc is not None else 0
            # Habilitar exportación si hay filas
            self.btn_exportar["state"] = "normal" if num_filas > 0 else "disabled"
            self._actualizar_estado(f"Mostrando todas las {num_filas} descripciones.")
            return # Termina aquí si la búsqueda era vacía

        # --- Ejecutar Búsqueda Normal ---
        self._actualizar_estado(f"Buscando '{termino}'...")
        # Limpiar tabla de resultados y deshabilitar exportación mientras busca
        self._actualizar_tabla(self.tabla_resultados, None)
        self.resultados_actuales = None
        self.btn_exportar["state"] = "disabled"

        # ¡Llamada principal al Motor de Búsqueda!
        resultados = self.motor.buscar(termino)

        # --- Procesar Resultados de la Búsqueda ---
        if resultados is None:
            # Error durante la búsqueda (el motor ya debería haber mostrado mensaje)
            self._actualizar_estado(f"Error durante la búsqueda de '{termino}'. Revise los logs.")
            # No hacemos nada más, la tabla ya está vacía
        elif isinstance(resultados, tuple):
            # El motor devuelve una tupla si el término no se encontró en el diccionario
            # (o fue negado completamente). La tupla contiene los DataFrames originales.
            logging.info(f"Término '{termino}' no encontrado o negado en el diccionario.")
            self._actualizar_estado(f"'{termino}' no encontrado/negado en el diccionario.")
            # Preguntar al usuario si quiere buscar directamente en descripciones
            resp = messagebox.askyesno(
                "Término no en Diccionario",
                f"El término o condición '{termino}' no se encontró (o fue negado) en el archivo Diccionario.\n\n¿Desea buscar '{termino}' directamente en el archivo de Descripciones?"
            )
            if resp: # Si el usuario dice sí
                self._actualizar_estado(f"Buscando '{termino}' directamente en descripciones...")
                # Marcamos el término guardado para indicar que fue búsqueda directa
                self.ultimo_termino_buscado = f"{termino} (Directo)"
                # Llamar al método de búsqueda directa del motor
                res_directos = self.motor.buscar_en_descripciones_directo(termino)
                # Actualizar tabla y resultados
                self._actualizar_tabla(self.tabla_resultados, res_directos)
                self.resultados_actuales = res_directos
                num_res = len(res_directos) if res_directos is not None else 0
                # Habilitar exportación si hay resultados
                self.btn_exportar["state"] = "normal" if num_res > 0 else "disabled"
                msg_final = f"Búsqueda directa de '{termino}': {num_res} resultados encontrados."
                self._actualizar_estado(msg_final)
                if num_res == 0:
                    messagebox.showinfo("Sin Coincidencias", msg_final)
                # Ejecutar demo del extractor si hay resultados y columnas
                if num_res > 0 and res_directos is not None and not res_directos.empty and len(res_directos.columns) > 0:
                    self._demo_extractor(res_directos, "búsqueda directa")

            else: # Si el usuario dice no
                self._actualizar_estado(f"Búsqueda de '{termino}' cancelada.")
                # La tabla de resultados permanece vacía
        elif isinstance(resultados, pd.DataFrame):
            # ¡Búsqueda normal exitosa! (Puede tener 0 o más filas)
            self.resultados_actuales = resultados # Guardar resultados
            num_res = len(resultados)
            self._actualizar_tabla(self.tabla_resultados, resultados) # Mostrar en tabla
            # Habilitar exportación si hay resultados
            self.btn_exportar["state"] = "normal" if num_res > 0 else "disabled"
            msg_final = f"Búsqueda '{termino}': {num_res} resultados encontrados."
            self._actualizar_estado(msg_final)
            # Informar si se encontraron términos pero 0 coincidencias finales
            if num_res == 0:
                messagebox.showinfo("Sin Coincidencias Finales", f"Se encontraron términos/condiciones en el diccionario para '{termino}', pero no produjeron coincidencias finales en las descripciones.")
            # Ejecutar demo del extractor si hay resultados y columnas
            if num_res > 0 and not resultados.empty and len(resultados.columns) > 0:
                self._demo_extractor(resultados, "búsqueda normal")
        else:
            # Caso inesperado, el motor devolvió algo raro
            logging.error(f"Tipo de resultado inesperado de motor.buscar: {type(resultados)}")
            self._actualizar_estado(f"Error: Tipo de resultado inesperado tras buscar '{termino}'.")
            messagebox.showerror("Error Interno", f"Se recibió un tipo de resultado inesperado ({type(resultados)}).")

        # >>> INICIO: Llamada a enfocar preview al final de la búsqueda (COMENTARIO) <<<
        # Ibar, esta es la línea que te comenté que me parece extraña aquí.
        # Llama a la función que busca en la *vista previa del diccionario* DESPUÉS
        # de haber mostrado los resultados de las *descripciones*.
        # Si no la necesitas, puedes comentarla o borrarla. Si la necesitas, ¡perfecto!
        # self._buscar_y_enfocar_en_preview()
        # <<< FIN: Llamada a enfocar preview al final de la búsqueda >>>


    def _demo_extractor(self, df_res: pd.DataFrame, tipo_busqueda: str): # (Sin cambios)
        # Intenta extraer magnitudes de la primera celda del primer resultado
        # Es solo una demostración, no afecta la funcionalidad principal
        if df_res is None or df_res.empty or len(df_res.columns) == 0:
            return # No hacer nada si no hay datos

        try:
            # Tomar el texto de la primera celda (fila 0, columna 0)
            texto_primera_celda = str(df_res.iloc[0, 0])
            logging.info(f"--- DEMO Extractor Magnitud (desde {tipo_busqueda}) ---")
            logging.info(f"Texto analizado: '{texto_primera_celda[:100]}...'") # Loguear inicio del texto
            encontrado = False
            # Iterar sobre las magnitudes predefinidas en el extractor
            for mag in self.extractor_magnitud.magnitudes:
                cantidad = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag, texto_primera_celda)
                # Si encuentra una cantidad para esa magnitud, la loguea
                if cantidad is not None:
                    logging.info(f"  -> Magnitud '{mag}': encontrada cantidad '{cantidad}'")
                    encontrado = True
            if not encontrado:
                logging.info("  (No se encontraron magnitudes predefinidas en este texto)")
            logging.info("--- FIN DEMO Extractor Magnitud ---")
        except IndexError:
             logging.warning(f"Error en demo extractor ({tipo_busqueda}): No se pudo acceder a la celda [0, 0].")
        except Exception as e:
            # Captura cualquier otro error durante la demo
            logging.warning(f"Error inesperado durante la demo del extractor ({tipo_busqueda}): {e}")

    # >>> INICIO: Método para sanitizar nombre archivo (Req 1 - Sin cambios) <<<
    def _sanitizar_nombre_archivo(self, texto: str, max_len: int = 50) -> str:
        """Limpia un texto para usarlo como parte segura de un nombre de archivo."""
        if not texto:
            return "resultados" # Nombre por defecto si el texto está vacío

        # 1. Reemplazar caracteres inválidos en nombres de archivo de Windows/Unix
        #    Los caracteres son: < > : " / \ | ? * y también # que usamos nosotros
        texto_limpio = re.sub(r'[<>:"/\\|?*#]', '_', texto)

        # 2. Reemplazar caracteres de control (aunque es raro tenerlos aquí)
        texto_limpio = "".join(c for c in texto_limpio if c not in string.control)

        # 3. Normalizar espacios: reemplazar múltiples espacios/saltos con uno solo, quitar espacios al inicio/final
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip()

        # 4. Limitar la longitud
        texto_cortado = texto_limpio[:max_len]

        # 5. Asegurarse de que no termine con punto o espacio (problemático en Windows)
        texto_final = texto_cortado.rstrip('._- ')

        # 6. Si después de todo queda vacío, devolver el default
        if not texto_final:
            return "resultados"

        return texto_final
    # <<< FIN: Método para sanitizar nombre archivo >>>

    def _exportar_resultados(self):
        # (Usa el último término buscado para el nombre de archivo sugerido)
        if self.resultados_actuales is None or self.resultados_actuales.empty:
            messagebox.showwarning("Sin Resultados", "No hay resultados para exportar. Realice una búsqueda primero.")
            return

        # >>> INICIO: Crear nombre archivo sugerido (Req 1 - Sin cambios) <<<
        base_nombre = "resultados" # Default
        # Si hubo una búsqueda previa, usar ese término (sanitizado)
        if self.ultimo_termino_buscado:
            term_sanitizado = self._sanitizar_nombre_archivo(self.ultimo_termino_buscado)
            base_nombre = f"resultados_{term_sanitizado}"
        # <<< FIN: Crear nombre archivo sugerido >>>

        # Definir tipos de archivo para el diálogo "Guardar como"
        tipos_archivo = [
            ("Archivo Excel (.xlsx)", "*.xlsx"),
            ("Archivo CSV (UTF-8)", "*.csv"),
            ("Archivo Excel 97-2003 (.xls)", "*.xls") # Formato antiguo
        ]

        # >>> INICIO: Usar initialfile (Req 1 - Sin cambios) <<<
        ruta_guardar = filedialog.asksaveasfilename(
            title="Guardar Resultados Como...",
            initialfile=f"{base_nombre}.xlsx", # Sugerir nombre con extensión .xlsx
            defaultextension=".xlsx", # Extensión por defecto si no se elige tipo
            filetypes=tipos_archivo
        )
        # <<< FIN: Usar initialfile >>>

        if not ruta_guardar:
            logging.info("Exportación cancelada por el usuario.")
            self._actualizar_estado("Exportación cancelada.")
            return # Si el usuario cancela

        self._actualizar_estado("Exportando resultados...")
        logging.info(f"Intentando exportar {len(self.resultados_actuales)} filas a: {ruta_guardar}")

        # (Lógica de exportación a diferentes formatos sin cambios)
        try:
            # Extraer extensión para decidir el formato
            extension = ruta_guardar.split('.')[-1].lower()
            df_a_exportar = self.resultados_actuales # Usamos los resultados actuales

            if extension == 'csv':
                # Guardar como CSV con codificación UTF-8 con BOM (para mejor compatibilidad Excel)
                df_a_exportar.to_csv(ruta_guardar, index=False, encoding='utf-8-sig')
            elif extension == 'xlsx':
                # Guardar como XLSX usando openpyxl (recomendado)
                df_a_exportar.to_excel(ruta_guardar, index=False, engine='openpyxl')
            elif extension == 'xls':
                # Guardar como XLS (formato antiguo, requiere xlwt)
                try:
                    import xlwt # type: ignore # Intentar importar xlwt
                    # Advertir y truncar si hay demasiadas filas para XLS
                    max_filas_xls = 65535
                    if len(df_a_exportar) > max_filas_xls:
                        messagebox.showwarning("Límite de Filas XLS", f"El formato .xls solo soporta {max_filas_xls} filas.\nSe exportarán solo las primeras {max_filas_xls} filas.")
                        df_a_exportar = df_a_exportar.head(max_filas_xls)
                    df_a_exportar.to_excel(ruta_guardar, index=False, engine='xlwt')
                except ImportError:
                    # Si xlwt no está instalado
                    logging.error("La librería 'xlwt' es necesaria para exportar a .xls.")
                    messagebox.showerror("Librería Faltante", "Para guardar en formato .xls antiguo, necesita instalar la librería 'xlwt'.\nPuede hacerlo con: pip install xlwt")
                    self._actualizar_estado("Error: Falta librería 'xlwt' para exportar a .xls.")
                    return # Detener exportación
            else:
                # Extensión no reconocida
                logging.error(f"Extensión de archivo no soportada para exportación: {extension}")
                messagebox.showerror("Formato No Soportado", f"La extensión de archivo '.{extension}' no está soportada para la exportación.")
                self._actualizar_estado(f"Error: Extensión '.{extension}' no soportada.")
                return # Detener exportación

            # Si todo fue bien
            logging.info(f"Exportación completada exitosamente a {ruta_guardar}")
            messagebox.showinfo("Exportación Exitosa", f"{len(df_a_exportar)} filas exportadas correctamente a:\n{ruta_guardar}")
            self._actualizar_estado(f"Resultados exportados a {os.path.basename(ruta_guardar)}.")

        except ImportError as ie:
             # Error si falta openpyxl (aunque se chequea al inicio, por si acaso)
            logging.exception("Falta librería necesaria para exportar (probablemente openpyxl).")
            messagebox.showerror("Librería Faltante", f"Falta una librería necesaria para exportar:\n{ie}\nAsegúrese de tener 'openpyxl' instalado.")
            self._actualizar_estado("Error: Falta librería para exportar.")
        except Exception as e:
            # Cualquier otro error durante la exportación
            logging.exception("Error inesperado durante la exportación de resultados.")
            messagebox.showerror("Error al Exportar", f"Ocurrió un error al intentar guardar el archivo:\n{e}")
            self._actualizar_estado("Error durante la exportación.")


    def on_closing(self): # (Sin cambios)
        """Acciones a realizar al cerrar la ventana principal."""
        logging.info("Cerrando la aplicación...")
        # Guardar la configuración actual antes de salir
        self._guardar_configuracion()
        # Destruir la ventana de Tkinter para cerrar la aplicación limpiamente
        self.destroy()

# --- Bloque Principal (`if __name__ == "__main__":`) ---
if __name__ == "__main__":
    # --- Configuración del Logging ---
    log_file = 'buscador_app.log' # Nombre del archivo de log
    logging.basicConfig(
        level=logging.DEBUG, # Nivel mínimo de mensajes a registrar (DEBUG es el más bajo)
        # Formato del mensaje de log: timestamp - nombre archivo:línea - nivel - mensaje
        format='%(asctime)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s',
        handlers=[
            # Handler para escribir logs a un archivo
            logging.FileHandler(log_file, encoding='utf-8', mode='w'), # 'w' sobrescribe el log cada vez
            # Handler para mostrar logs en la consola también
            logging.StreamHandler()
        ]
    )

    # --- Mensajes de Inicio en el Log ---
    logging.info("=============================================")
    logging.info("=== Iniciando Aplicación Buscador Avanzado ===")
    # Podrías añadir una versión aquí si la tienes: logging.info("=== Versión: 1.0 ===")
    logging.info(f"Plataforma: {platform.system()} {platform.release()}")
    logging.info(f"Versión de Python: {platform.python_version()}")
    logging.info("=============================================")

    # --- Comprobación de Dependencias Críticas ---
    missing_deps = []
    try:
        import pandas as pd
        logging.info(f"Pandas versión: {pd.__version__}")
    except ImportError:
        missing_deps.append("pandas")
        logging.critical("Dependencia crítica faltante: pandas")

    try:
        import openpyxl
        logging.info(f"openpyxl versión: {openpyxl.__version__}")
    except ImportError:
        missing_deps.append("openpyxl")
        logging.critical("Dependencia crítica faltante: openpyxl (necesaria para .xlsx)")

    # xlwt es opcional (solo para .xls), se comprueba al exportar si es necesario.

    # Si falta alguna dependencia crítica, mostrar error y salir
    if missing_deps:
        error_msg = f"Faltan librerías esenciales para ejecutar la aplicación: {', '.join(missing_deps)}.\n\nPor favor, instálelas usando pip:\npip install {' '.join(missing_deps)}"
        logging.critical(error_msg)
        # Crear una ventana mínima solo para mostrar el error si Tkinter está disponible
        try:
            root = tk.Tk()
            root.withdraw() # Ocultar la ventana principal vacía
            messagebox.showerror("Dependencias Faltantes", error_msg)
            root.destroy()
        except tk.TclError:
             # Si Tkinter mismo falla, al menos el log lo registra
             print(f"ERROR CRÍTICO: {error_msg}")
        exit(1) # Salir del programa

    # --- Iniciar la Interfaz Gráfica ---
    try:
        app = InterfazGrafica() # Crear la instancia de la aplicación
        app.mainloop() # Iniciar el bucle principal de Tkinter
    except Exception as main_error:
        # Capturar cualquier error fatal no manejado durante la ejecución de la app
        logging.critical("¡Error fatal no capturado en la aplicación!", exc_info=True) # exc_info=True añade el traceback al log
        # Intentar mostrar un último mensaje de error al usuario
        try:
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror("Error Fatal", f"Ha ocurrido un error crítico inesperado:\n{main_error}\n\nLa aplicación se cerrará. Consulte el archivo '{log_file}' para más detalles.")
            root_err.destroy()
        except Exception as fallback_error:
            # Si hasta mostrar el mensaje de error falla...
            logging.error(f"No se pudo mostrar el mensaje de error fatal al usuario: {fallback_error}")
            print(f"ERROR FATAL: {main_error}. Consulte {log_file}.") # Mensaje por consola como último recurso
    finally:
        # Este bloque se ejecuta siempre, tanto si hubo error como si se cerró normalmente
        logging.info("=== Finalizando Aplicación Buscador ===")
