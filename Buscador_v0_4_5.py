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
        # (Sin cambios funcionales desde la última versión con todos los operadores y negación)
        if df is None or df.empty or not cols_nombres or not terms_analizados: return pd.Series(False, index=df.index if df is not None else None)
        cols_ok = [c for c in cols_nombres if c in df.columns];
        if not cols_ok: logging.error(f"Ninguna col {cols_nombres} válida."); return pd.Series(False, index=df.index)
        terms_pos = [i for i in terms_analizados if not i['negate']]; terms_neg = [i for i in terms_analizados if i['negate']]
        # Máscara positiva
        if terms_pos:
            mask_pos = pd.Series(op_principal.upper() == 'AND', index=df.index) # True para AND, False para OR
            for item in terms_pos:
                mask_item = pd.Series(False, index=df.index); tipo, valor = item['tipo'], item['valor']
                for col_n in cols_ok:
                    col = df[col_n]; mask_col = pd.Series(False, index=df.index)
                    try: # Lógica de comparación/búsqueda
                        if tipo == 'str': mask_col = col.astype(str).str.contains(re.escape(str(valor)), case=False, na=False, regex=True)
                        elif tipo in ['gt','lt','ge','le']:
                            col_num = pd.to_numeric(col, errors='coerce')
                            if tipo == 'gt': mask_col = col_num > valor
                            elif tipo == 'lt': mask_col = col_num < valor
                            elif tipo == 'ge': mask_col = col_num >= valor
                            else: mask_col = col_num <= valor #le
                            mask_col = mask_col.fillna(False)
                        elif tipo == 'range':
                            min_v, max_v = valor; col_num = pd.to_numeric(col, errors='coerce')
                            mask_col = (col_num >= min_v) & (col_num <= max_v); mask_col = mask_col.fillna(False)
                        mask_item |= mask_col
                    except Exception as e: logging.warning(f"Error {item['original']} en col '{col_n}': {e}")
                if op_principal.upper() == 'AND': mask_pos &= mask_item
                else: mask_pos |= mask_item
        else: mask_pos = pd.Series(op_principal.upper() == 'AND', index=df.index) # Si no hay positivos
        # Máscara negativa (True si CUALQUIER negación coincide)
        mask_neg_comb = pd.Series(False, index=df.index)
        if terms_neg:
            for item in terms_neg:
                mask_item_neg = pd.Series(False, index=df.index); tipo, valor = item['tipo'], item['valor']
                for col_n in cols_ok:
                    col = df[col_n]; mask_col = pd.Series(False, index=df.index)
                    try: # Misma lógica de comparación
                        if tipo == 'str': mask_col = col.astype(str).str.contains(re.escape(str(valor)), case=False, na=False, regex=True)
                        elif tipo in ['gt','lt','ge','le']:
                            col_num = pd.to_numeric(col, errors='coerce')
                            if tipo == 'gt': mask_col = col_num > valor
                            elif tipo == 'lt': mask_col = col_num < valor
                            elif tipo == 'ge': mask_col = col_num >= valor
                            else: mask_col = col_num <= valor
                            mask_col = mask_col.fillna(False)
                        elif tipo == 'range':
                            min_v, max_v = valor; col_num = pd.to_numeric(col, errors='coerce')
                            mask_col = (col_num >= min_v) & (col_num <= max_v); mask_col = mask_col.fillna(False)
                        mask_item_neg |= mask_col
                    except Exception as e: logging.warning(f"Error neg {item['original']} en col '{col_n}': {e}")
                mask_neg_comb |= mask_item_neg # OR para combinar negaciones
        # Combinar final: Positivo Y NO Negativo
        mask_final = mask_pos & (~mask_neg_comb)
        logging.debug(f"Mask Pos: {mask_pos.sum()}, Mask Neg: {mask_neg_comb.sum()}, Mask Final: {mask_final.sum()}")
        return mask_final

    def _busqueda_simple(self, df_dic: pd.DataFrame, df_desc: pd.DataFrame, term: str) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        # (Sin cambios funcionales, usa lógica unificada)
        logging.info(f"Búsqueda Simple: '{term}'"); cols_n = self._obtener_nombres_columnas_busqueda(df_dic)
        if cols_n is None: return (df_dic, df_desc)
        terms_b = [term.strip()]; terms_a = self._analizar_terminos(terms_b)
        if not terms_a: logging.warning(f"Término simple inválido: '{term}'"); messagebox.showwarning("Inválido", f"'{term}'"); return (df_dic, df_desc)
        mask = self._aplicar_mascara_diccionario(df_dic, cols_n, terms_a, 'OR')
        if not mask.any(): logging.info(f"'{term}' no encontrado/negado."); return (df_dic, df_desc)
        logging.info(f"'{term}' encontrado. Extrayendo..."); coinc = df_dic[mask]; terms_d = self._extraer_terminos_diccionario(coinc, cols_n)
        if not terms_d: logging.warning(f"Sin términos válidos extraídos."); messagebox.showinfo("Aviso", f"{len(coinc)} fila(s) pero sin términos válidos."); return pd.DataFrame(columns=df_desc.columns)
        logging.info(f"Buscando {len(terms_d)} términos en descrips..."); return self._buscar_terminos_en_descripciones(df_desc, terms_d)

    def _busqueda_compuesta(self, df_dic: pd.DataFrame, df_desc: pd.DataFrame, term: str, sep: str, op: str, req_all: bool) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        # (Sin cambios funcionales, usa lógica unificada)
        logging.info(f"Búsqueda Compuesta ({op}, sep='{sep}'): '{term}'"); cols_n = self._obtener_nombres_columnas_busqueda(df_dic)
        if cols_n is None: return (df_dic, df_desc)
        terms_b = [p.strip() for p in term.split(sep) if p.strip()]
        if not terms_b: logging.warning(f"Sin términos válidos con sep '{sep}'."); messagebox.showwarning("Inválido", f"'{term}'"); return (df_dic, df_desc)
        terms_a = self._analizar_terminos(terms_b)
        if not terms_a: logging.warning(f"Términos inválidos tras análisis: '{term}'."); messagebox.showwarning("Inválido", f"'{term}'"); return (df_dic, df_desc)
        mask = self._aplicar_mascara_diccionario(df_dic, cols_n, terms_a, op)
        if not mask.any(): logging.info(f"Combinación '{term}' ({op}) no encontrada/negada."); return (df_dic, df_desc)
        logging.info(f"Combinación '{term}' ({op}) encontrada. Extrayendo..."); coinc = df_dic[mask]; terms_d = self._extraer_terminos_diccionario(coinc, cols_n)
        if not terms_d: logging.warning(f"Sin términos válidos extraídos."); messagebox.showinfo("Aviso", f"{len(coinc)} fila(s) pero sin términos válidos."); return pd.DataFrame(columns=df_desc.columns)
        logging.info(f"Buscando {len(terms_d)} términos en descrips..."); return self._buscar_terminos_en_descripciones(df_desc, terms_d, require_all=req_all)

    def buscar(self, term_buscar: str) -> Union[None, pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        # (Sin cambios funcionales, usa lógica unificada)
        logging.info(f"--- Nueva Búsqueda --- Término: '{term_buscar}'")
        if self.datos_diccionario is None: logging.error("Diccionario no cargado."); messagebox.showwarning("Falta", "Cargue Diccionario."); return None
        if self.datos_descripcion is None: logging.error("Descripciones no cargadas."); messagebox.showwarning("Falta", "Cargue Descripciones."); return None
        term_proc = term_buscar.strip();
        if not term_proc: logging.info("Vacía."); return self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame()
        df_dic, df_desc = self.datos_diccionario.copy(), self.datos_descripcion.copy()
        if df_dic.empty: logging.error("DF Diccionario vacío."); messagebox.showerror("Error", "Diccionario vacío."); return None
        if df_desc.empty: logging.error("DF Descripciones vacío."); messagebox.showerror("Error", "Descripciones vacío."); return pd.DataFrame(columns=df_desc.columns)
        try:
            if '+' in term_proc: return self._busqueda_compuesta(df_dic, df_desc, term_proc, '+', 'AND', False)
            elif '|' in term_proc: return self._busqueda_compuesta(df_dic, df_desc, term_proc, '|', 'OR', False)
            elif '/' in term_proc: return self._busqueda_compuesta(df_dic, df_desc, term_proc, '/', 'OR', False)
            else: return self._busqueda_simple(df_dic, df_desc, term_proc)
        except Exception as e: logging.exception("Error orquestación búsqueda."); messagebox.showerror("Error", f"{e}"); return None

    def buscar_en_descripciones_directo(self, term_buscar: str) -> pd.DataFrame:
         # (Sin cambios funcionales, solo separador '/')
        logging.info(f"Búsqueda Directa: '{term_buscar}'");
        if self.datos_descripcion is None or self.datos_descripcion.empty: logging.warning("Descrips no cargadas."); return pd.DataFrame()
        term_limpio = term_buscar.strip().upper();
        if not term_limpio: return self.datos_descripcion.copy()
        df_desc = self.datos_descripcion.copy(); res = pd.DataFrame(columns=df_desc.columns)
        try:
            txt_filas = df_desc.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            mask = pd.Series(False, index=df_desc.index)
            if '+' in term_limpio:
                ps = [p.strip() for p in term_limpio.split('+') if p.strip()]; mask = pd.Series(True, index=df_desc.index) # AND
                if not ps: return res
                for p in ps: mask &= txt_filas.str.contains(r"\b"+re.escape(p)+r"\b", regex=True, na=False)
            elif '|' in term_limpio or '/' in term_limpio: # OR
                sep = '|' if '|' in term_limpio else '/'; ps = [p.strip() for p in term_limpio.split(sep) if p.strip()]
                if not ps: return res
                for p in ps: mask |= txt_filas.str.contains(r"\b"+re.escape(p)+r"\b", regex=True, na=False)
            else: mask = txt_filas.str.contains(r"\b"+re.escape(term_limpio)+r"\b", regex=True, na=False) # Simple
            res = df_desc[mask]; logging.info(f"Búsqueda directa OK. Resultados: {len(res)}.")
        except Exception as e: logging.exception("Error búsqueda directa."); messagebox.showerror("Error", f"{e}"); return pd.DataFrame(columns=df_desc.columns)
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
            forma_normalizada = unicodedata.normalize('NFKD', texto)
            return ''.join(c for c in forma_normalizada if not unicodedata.combining(c))
        except TypeError: return ""

    def buscar_cantidad_para_magnitud(self, mag: str, descripcion: str) -> Optional[str]:
        if not isinstance(mag, str) or not mag: return None
        if not isinstance(descripcion, str) or not descripcion: return None
        mag_upper = mag.upper()
        texto_limpio = self._quitar_diacronicos_y_acentos(descripcion.upper())
        if not texto_limpio: return None
        mag_escapada = re.escape(mag_upper)
        patron_principal = re.compile(
            r"(\d+([.,]\d+)?)[ X]{0,1}(\b" + mag_escapada + r"\b)(?![a-zA-Z0-9])"
        )
        for match in patron_principal.finditer(texto_limpio):
            return match.group(1).strip()
        return None

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
                logging.info(f"Config cargada: {self.CONFIG_FILE}")
            except Exception as e: logging.error(f"Error cargar config: {e}"); messagebox.showwarning("Config", f"{e}")
        else: logging.info("Config no encontrada. Se creará.")
        config.setdefault("last_dic_path", None); config.setdefault("last_desc_path", None)
        config.setdefault("indices_columnas_busqueda_dic", [0, 3]) # Default si falta
        return config

    def _guardar_configuracion(self): # (Sin cambios)
        self.config["last_dic_path"] = self.motor.archivo_diccionario_actual
        self.config["last_desc_path"] = self.motor.archivo_descripcion_actual
        self.config["indices_columnas_busqueda_dic"] = self.motor.indices_columnas_busqueda_dic
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(self.config, f, indent=4)
            logging.info(f"Config guardada: {self.CONFIG_FILE}")
        except Exception as e: logging.error(f"Error guardar config: {e}"); messagebox.showerror("Config", f"{e}")

    def _configurar_estilo_ttk(self): # (Sin cambios)
        style = ttk.Style(self); themes = style.theme_names(); os_name = platform.system()
        prefs = {"Windows":["vista","xpnative","clam"],"Darwin":["aqua","clam"],"Linux":["clam","alt","default"]}
        theme = next((t for t in prefs.get(os_name, ["clam","default"]) if t in themes), None)
        if not theme: theme = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))
        if theme:
            logging.info(f"Usando tema: {theme}")
            try:
                style.theme_use(theme)
            except tk.TclError:
                pass
        else: logging.warning("No se encontró tema TTK.")

    def _crear_widgets(self):
        self.marco_controles = ttk.LabelFrame(self, text="Controles")
        # Carga
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.lbl_dic_cargado = ttk.Label(self.marco_controles, text="Dic: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN)
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion, state="disabled")
        self.lbl_desc_cargado = ttk.Label(self.marco_controles, text="Desc: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN)
        # Búsqueda
        self.lbl_busqueda = ttk.Label(self.marco_controles, text="Buscar (+ AND, |/ OR, # NOT, >, <, >=, <=, num-num):")
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50)
        # >>> INICIO: Botón Buscar en Vista Previa (Req 2) <<<
        self.btn_buscar_en_preview = ttk.Button(self.marco_controles, text="Ir a", command=self._buscar_y_enfocar_en_preview, state="disabled", width=5)
        # <<< FIN: Botón Buscar en Vista Previa >>>
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda, state="disabled")
        self.btn_ayuda = ttk.Button(self.marco_controles, text="?", command=self._mostrar_ayuda, width=3)
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar", command=self._exportar_resultados, state="disabled") # Texto más corto

        # Tablas y Barras
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario:")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados / Descripciones:")
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings", height=8) # Limitar altura preview
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings")
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)
        self.barra_estado = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W)
        self._actualizar_etiquetas_archivos() # Actualizar etiquetas con config inicial

    def _configurar_tags_treeview(self): # (Sin cambios)
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par)
            tabla.tag_configure('impar', background=self.color_fila_impar)

    def _configurar_grid(self):
        # Grid ventana principal
        self.grid_rowconfigure(2, weight=1) # Preview Diccionario
        self.grid_rowconfigure(4, weight=3) # Resultados
        self.grid_columnconfigure(0, weight=1)
        # Marco controles
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        # Fila 0 del marco (Carga) - Ajustar para nuevas etiquetas
        self.marco_controles.grid_columnconfigure(4, weight=1) # Hacer que descripciones expanda un poco
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, columnspan=2, padx=(2,5), pady=5, sticky="ew") # Ocupa más espacio
        # Fila 1 del marco (Búsqueda) - Reorganizar botones
        self.marco_controles.grid_columnconfigure(2, weight=1) # Columna de entrada expande
        self.lbl_busqueda.grid(row=1, column=0, columnspan=2, padx=5, pady=(5,0), sticky="w")
        self.entrada_busqueda.grid(row=1, column=2, padx=5, pady=(5,5), sticky="ew")
        # >>> INICIO: Añadir btn_buscar_en_preview al grid <<<
        self.btn_buscar_en_preview.grid(row=1, column=3, padx=(5,0), pady=(5,5), sticky="w")
        # <<< FIN: Añadir btn_buscar_en_preview al grid >>>
        self.btn_buscar.grid(row=1, column=4, padx=(2,0), pady=(5,5), sticky="w") # Mover Buscar
        self.btn_ayuda.grid(row=1, column=5, padx=(2,5), pady=(5,5), sticky="w")  # Mover Ayuda
        self.btn_exportar.grid(row=1, column=6, padx=(10, 5), pady=5, sticky="e")# Mover Exportar

        # Etiquetas Tablas
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0))
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))
        # Frames y Tablas (Grid interno sin cambios)
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1); self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew"); self.scrolly_diccionario.grid(row=0, column=1, sticky="ns"); self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1); self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew"); self.scrolly_resultados.grid(row=0, column=1, sticky="ns"); self.scrollx_resultados.grid(row=1, column=0, sticky="ew")
        # Barra de estado
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0))

    def _configurar_eventos(self): # (Sin cambios)
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())
        self.protocol("WM_DELETE_WINDOW", self.on_closing) # Guardar config al cerrar

    def _actualizar_estado(self, mensaje: str): # (Sin cambios)
        self.barra_estado.config(text=mensaje)
        self.update_idletasks()

    def _mostrar_ayuda(self): # (Sin cambios)
        ayuda = """Sintaxis de Búsqueda:
- Texto: Busca la palabra/frase.
- `t1 + t2`: AND (ambos deben coincidir).
- `t1 | t2`: OR (al menos uno coincide).
- `t1 / t2`: OR (alternativo).
- `>n`, `<n`, `>=n`, `<=n`: Comparación numérica.
- `n1-n2`: Rango numérico (ambos incluidos).
- `#t`: NOT (excluye filas que coincidan con t).
  (t puede ser texto, comparación o rango).

Notas: Texto insensible a mayús/minús. Comparaciones/rangos aplican a columnas numéricas. # se aplica al final."""
        messagebox.showinfo("Ayuda - Sintaxis", ayuda)

    def _configurar_orden_tabla(self, tabla: ttk.Treeview): # (Sin cambios)
        cols = tabla["columns"]
        if cols:
            for col in cols:
                tabla.heading(col, text=col, anchor=tk.W, command=lambda c=col: self._ordenar_columna(tabla, c, False))

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool): # (Sin cambios)
        if self.resultados_actuales is None or self.resultados_actuales.empty: return
        logging.info(f"Ordenando por '{col}', reverse={reverse}")
        try:
            df_ordenado = self.resultados_actuales.sort_values(
                by=col, ascending=not reverse, na_position='last',
                key=lambda x: pd.to_numeric(x, errors='coerce')
            )
            self.resultados_actuales = df_ordenado
            self._actualizar_tabla(tabla, self.resultados_actuales) # Redibujar
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, not reverse))
        except Exception as e: logging.exception(f"Error ordenando por '{col}'"); messagebox.showerror("Error", f"{e}")

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = 1000, columnas_a_mostrar: Optional[List[str]] = None):
        # (Sin cambios funcionales)
        try:
            for i in tabla.get_children(): tabla.delete(i)
        except tk.TclError: pass
        if datos is None or datos.empty: tabla["columns"] = (); return
        cols_disp = columnas_a_mostrar if columnas_a_mostrar else list(datos.columns)
        cols_ok = [c for c in cols_disp if c in datos.columns];
        if not cols_ok: cols_ok = list(datos.columns) # Fallback a todas si las especificadas no existen
        df_mostrar = datos[cols_ok]
        tabla["columns"] = cols_ok
        for col in cols_ok: # Configurar cabeceras/anchos
            tabla.heading(col, text=str(col), anchor=tk.W); tabla.column(col, anchor=tk.W) # Reset comando aquí
            try: # Calcular ancho
                s = min(len(df_mostrar),100); sub = df_mostrar.iloc[:s][col].dropna().astype(str)
                cw=sub.str.len().max() if not sub.empty else 0; hw=len(str(col))
                w=max(70,min(int(max(hw*9,cw*7)+20),400)); tabla.column(col,width=w,minwidth=70)
            except: tabla.column(col,width=100,minwidth=50)
        # Insertar filas
        df_final = df_mostrar.head(limite_filas) if limite_filas is not None and len(df_mostrar) > limite_filas else df_mostrar
        for i, (_, row) in enumerate(df_final.iterrows()):
            vals = [str(v) if pd.notna(v) else "" for v in row.values]
            tag = 'par' if i%2==0 else 'impar'
            try: tabla.insert("", "end", values=vals, tags=(tag,))
            except tk.TclError: # Fallback simple
                 try: tabla.insert("", "end", values=[str(v).encode('ascii', 'ignore').decode('ascii') for v in row.values], tags=(tag,))
                 except: pass
        # Reconfigurar ordenación si es la tabla de resultados
        if tabla == self.tabla_resultados: self._configurar_orden_tabla(tabla)


    def _actualizar_etiquetas_archivos(self):
        """Actualiza las etiquetas que muestran los nombres de los archivos cargados."""
        dic_name = os.path.basename(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else "Ninguno"
        desc_name = os.path.basename(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else "Ninguno"
        self.lbl_dic_cargado.config(text=f"Dic: {dic_name}")
        self.lbl_desc_cargado.config(text=f"Desc: {desc_name}")
        # Ajustar tooltips si se quiere
        # self.lbl_dic_cargado.tooltip = self.motor.archivo_diccionario_actual
        # self.lbl_desc_cargado.tooltip = self.motor.archivo_descripcion_actual


    def _cargar_diccionario(self):
        # (Usa config, guarda config, actualiza UI)
        last_dir = os.path.dirname(self.config.get("last_dic_path", "") or "") or None
        ruta = filedialog.askopenfilename(title="Seleccionar Diccionario", filetypes=[("Excel", "*.xlsx *.xls")], initialdir=last_dir)
        if not ruta: return
        self._actualizar_estado(f"Cargando dic: {os.path.basename(ruta)}...")
        self._actualizar_tabla(self.tabla_diccionario, None)
        self.btn_cargar_descripciones["state"] = "disabled"; self.btn_buscar["state"] = "disabled"; self.btn_exportar["state"] = "disabled"; self.btn_buscar_en_preview["state"] = "disabled"
        if self.motor.cargar_excel_diccionario(ruta):
            self._guardar_configuracion(); df_dic = self.motor.datos_diccionario
            if df_dic is not None: # Carga y validación OK
                n = len(df_dic); self._actualizar_estado(f"Diccionario ({n} filas). Vista previa...")
                cols = self.motor._obtener_nombres_columnas_busqueda(df_dic)
                indices_str = ', '.join(map(str, self.motor.indices_columnas_busqueda_dic))
                lbl_text = f"Vista Previa Dic (Índices: {indices_str}):"
                if cols: lbl_text = f"Vista Previa Dic ({', '.join(cols)} - Índices: {indices_str}):"
                self.lbl_tabla_diccionario.config(text=lbl_text)
                self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols) # Límite preview
                self.title(f"Buscador - Dic: {os.path.basename(ruta)}")
                self.btn_cargar_descripciones["state"] = "normal"
                self.btn_buscar_en_preview["state"] = "normal" # Habilitar botón preview
                if self.motor.datos_descripcion is not None: self.btn_buscar["state"] = "normal"
                self._actualizar_estado(f"Diccionario '{os.path.basename(ruta)}' ({n}) cargado.")
            # else: # Error ya manejado por el motor
        else: self._actualizar_estado("Error al cargar diccionario."); self.title("Buscador")
        self._actualizar_etiquetas_archivos()

    def _cargar_excel_descripcion(self):
        # (Usa config, guarda config, actualiza UI)
        last_dir = os.path.dirname(self.config.get("last_desc_path", "") or "") or None
        ruta = filedialog.askopenfilename(title="Seleccionar Descripciones", filetypes=[("Excel", "*.xlsx *.xls")], initialdir=last_dir)
        if not ruta: return
        self._actualizar_estado(f"Cargando desc: {os.path.basename(ruta)}...")
        self._actualizar_tabla(self.tabla_resultados, None); self.resultados_actuales = None; self.ultimo_termino_buscado = None # Resetear término
        self.btn_buscar["state"] = "disabled"; self.btn_exportar["state"] = "disabled"
        if self.motor.cargar_excel_descripcion(ruta):
            self._guardar_configuracion(); df_desc = self.motor.datos_descripcion
            if df_desc is not None:
                 n = len(df_desc); self._actualizar_estado(f"Descripciones ({n} filas). Mostrando...");
                 self._actualizar_tabla(self.tabla_resultados, df_desc) # Mostrar todas
                 self.resultados_actuales = df_desc.copy()
                 self.btn_exportar["state"] = "normal" # Puede exportar todo
                 if self.motor.datos_diccionario is not None: self.btn_buscar["state"] = "normal"; self.btn_buscar_en_preview["state"] = "normal"
                 if self.motor.archivo_diccionario_actual:
                     dic_n=os.path.basename(self.motor.archivo_diccionario_actual); desc_n=os.path.basename(ruta)
                     self.title(f"Buscador - Dic: {dic_n} | Desc: {desc_n}")
                 self._actualizar_estado(f"Descripciones '{os.path.basename(ruta)}' ({n}) cargadas.")
            # else: error ya manejado
        else: self._actualizar_estado("Error al cargar descripciones.")
        self._actualizar_etiquetas_archivos()

    # >>> INICIO: Método para Buscar y Enfocar en Preview (Req 2) <<<
    def _buscar_y_enfocar_en_preview(self):
        """Busca el término de entrada en la tabla de vista previa y hace scroll/selecciona."""
        termino_buscar = self.entrada_busqueda.get().strip()
        if not termino_buscar:
            messagebox.showinfo("Vacío", "Ingrese un término para buscar en la vista previa.")
            return
        if self.motor.datos_diccionario is None or self.tabla_diccionario.get_children() == ():
             messagebox.showwarning("Sin Datos", "Cargue el diccionario y asegúrese de que la vista previa tenga datos.")
             return

        termino_upper = termino_buscar.upper() # Búsqueda insensible al caso
        logging.info(f"Buscando '{termino_buscar}' en la vista previa del diccionario...")
        self._actualizar_estado(f"Buscando '{termino_buscar}' en vista previa...")

        items_preview = self.tabla_diccionario.get_children('')
        col_indices_preview = {self.tabla_diccionario.heading(col)["text"]: i for i, col in enumerate(self.tabla_diccionario["columns"])}

        found_item = None
        for item_id in items_preview:
            try:
                valores = self.tabla_diccionario.item(item_id, 'values')
                # Convertir valores a string y buscar substring (case-insensitive)
                # Solo buscamos en las columnas que se están mostrando
                if any(termino_upper in str(valores[idx]).upper() for col_name, idx in col_indices_preview.items()):
                    found_item = item_id
                    break # Detenerse en la primera coincidencia
            except Exception as e:
                 logging.warning(f"Error procesando item {item_id} en preview: {e}")
                 continue # Saltar al siguiente item

        if found_item:
            logging.info(f"Término '{termino_buscar}' encontrado en preview (item: {found_item}). Enfocando...")
            try:
                # Deseleccionar cualquier cosa previa
                current_selection = self.tabla_diccionario.selection()
                if current_selection:
                    self.tabla_diccionario.selection_remove(current_selection)

                self.tabla_diccionario.selection_set(found_item) # Seleccionar
                self.tabla_diccionario.see(found_item)          # Hacer scroll
                self.tabla_diccionario.focus(found_item)        # Dar foco (opcional)
                self._actualizar_estado(f"'{termino_buscar}' encontrado y enfocado en vista previa.")
                # Devolver foco a la entrada para seguir escribiendo
                self.entrada_busqueda.focus_set()
            except tk.TclError as e:
                logging.error(f"Error Tcl al enfocar item {found_item}: {e}")
                self._actualizar_estado(f"Error al enfocar '{termino_buscar}'.")
        else:
            logging.info(f"Término '{termino_buscar}' no encontrado en las filas visibles de la vista previa.")
            messagebox.showinfo("No Encontrado", f"'{termino_buscar}' no encontrado en las filas actuales de la vista previa.")
            self._actualizar_estado(f"'{termino_buscar}' no encontrado en vista previa.")
            self.entrada_busqueda.focus_set() # Devolver foco igualmente
    # <<< FIN: Método para Buscar y Enfocar en Preview >>>


    def _ejecutar_busqueda(self):
        # (Guarda el término buscado para usarlo en exportación)
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Faltan Archivos", "Cargue Diccionario y Descripciones."); return
        termino = self.entrada_busqueda.get()
        # >>> INICIO: Guardar término para exportación (Req 1) <<<
        self.ultimo_termino_buscado = termino if termino.strip() else None # Guardar si no está vacío
        # <<< FIN: Guardar término para exportación >>>

        if not termino.strip(): # Búsqueda vacía
            df_desc = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc); self.resultados_actuales = df_desc.copy() if df_desc is not None else None
            n = len(df_desc) if df_desc is not None else 0; self.btn_exportar["state"] = "normal" if n > 0 else "disabled"
            self._actualizar_estado(f"Mostrando {n} descripciones."); return

        self._actualizar_estado(f"Buscando '{termino}'..."); self._actualizar_tabla(self.tabla_resultados, None)
        self.resultados_actuales = None; self.btn_exportar["state"] = "disabled"
        resultados = self.motor.buscar(termino) # LLAMADA AL MOTOR

        # Procesar resultados (sin cambios funcionales)
        if resultados is None: self._actualizar_estado(f"Error buscando '{termino}'.")
        elif isinstance(resultados, tuple): # No en diccionario
            self._actualizar_estado(f"'{termino}' no encontrado/negado en diccionario.");
            resp = messagebox.askyesno("No en Diccionario", f"'{termino}' no encontrado/negado.\n¿Buscar directo?")
            if resp: # Búsqueda directa
                self._actualizar_estado(f"Buscando '{termino}' directo..."); self.ultimo_termino_buscado = f"{termino} (Directo)" # Marcar como directo
                res_directos = self.motor.buscar_en_descripciones_directo(termino)
                self._actualizar_tabla(self.tabla_resultados, res_directos); self.resultados_actuales = res_directos
                n = len(res_directos) if res_directos is not None else 0; self.btn_exportar["state"] = "normal" if n > 0 else "disabled"
                msg = f"Búsqueda directa '{termino}': {n} resultados."; self._actualizar_estado(msg);
                if n == 0: messagebox.showinfo("Sin Coincidencias", msg)
                if n > 0 and len(res_directos.columns)>0: self._demo_extractor(res_directos, "directo")
            else: self._actualizar_estado(f"Búsqueda '{termino}' cancelada.")
        elif isinstance(resultados, pd.DataFrame): # Búsqueda normal OK
            self.resultados_actuales = resultados; n = len(resultados)
            self._actualizar_tabla(self.tabla_resultados, resultados)
            self.btn_exportar["state"] = "normal" if n > 0 else "disabled"
            msg = f"Búsqueda '{termino}': {n} resultados."; self._actualizar_estado(msg)
            if n == 0: messagebox.showinfo("Sin Coincidencias", f"Términos/condiciones encontradas para '{termino}', pero 0 coincidencias finales.")
            if n > 0 and len(resultados.columns)>0: self._demo_extractor(resultados, "normal")
        else: self._actualizar_estado(f"Error: Tipo resultado inesperado ({type(resultados)}).")

    def _demo_extractor(self, df_res: pd.DataFrame, tipo_busqueda: str): # (Sin cambios)
        try:
            txt = str(df_res.iloc[0, 0]); logging.info(f"--- DEMO Extractor ({tipo_busqueda}) ---")
            found = False
            for mag in self.extractor_magnitud.magnitudes:
                cant = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag, txt)
                if cant is not None: logging.info(f"  -> '{mag}': {cant}"); found = True
            if not found: logging.info("  (No se encontraron mags predefinidas)")
            logging.info("--- FIN DEMO ---")
        except Exception as e: logging.warning(f"Error demo extractor ({tipo_busqueda}): {e}")

    # >>> INICIO: Método para sanitizar nombre archivo (Req 1) <<<
    def _sanitizar_nombre_archivo(self, texto: str, max_len: int = 50) -> str:
        """Limpia un texto para usarlo como parte de un nombre de archivo."""
        if not texto: return "resultados"
        # Quitar operadores problemáticos y espacios múltiples
        texto_limpio = re.sub(r'[<>:"/\\|?*#]', '_', texto) # Reemplazar caracteres inválidos
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip() # Normalizar espacios
        # Limitar longitud
        return texto_limpio[:max_len].rstrip('._- ') # Quitar basura al final tras cortar
    # <<< FIN: Método para sanitizar nombre archivo >>>


    def _exportar_resultados(self):
        # (Usa el último término buscado para el nombre de archivo sugerido)
        if self.resultados_actuales is None or self.resultados_actuales.empty:
             messagebox.showwarning("Sin Resultados", "No hay resultados."); return

        # >>> INICIO: Crear nombre archivo sugerido (Req 1) <<<
        base_nombre = "resultados"
        if self.ultimo_termino_buscado:
             term_sanitizado = self._sanitizar_nombre_archivo(self.ultimo_termino_buscado)
             base_nombre = f"resultados_{term_sanitizado}"
        # <<< FIN: Crear nombre archivo sugerido >>>

        tipos = [("Excel", "*.xlsx"),("CSV UTF-8", "*.csv"),("Excel 97-2003", "*.xls")]
        # >>> INICIO: Usar initialfile (Req 1) <<<
        ruta = filedialog.asksaveasfilename(title="Guardar Resultados",
                                            initialfile=f"{base_nombre}.xlsx", # Nombre sugerido
                                            defaultextension=".xlsx",
                                            filetypes=tipos)
        # <<< FIN: Usar initialfile >>>
        if not ruta: logging.info("Exportación cancelada."); return

        self._actualizar_estado("Exportando..."); logging.info(f"Exportando {len(self.resultados_actuales)} filas a {ruta}")
        # (Lógica de exportación a diferentes formatos sin cambios)
        try:
            ext=ruta.split('.')[-1].lower(); df_exp=self.resultados_actuales
            if ext=='csv': df_exp.to_csv(ruta, index=False, encoding='utf-8-sig')
            elif ext=='xlsx': df_exp.to_excel(ruta, index=False, engine='openpyxl')
            elif ext=='xls':
                try:
                    import xlwt # type: ignore
                    if len(df_exp)>65535: messagebox.showwarning("Límite XLS","Truncado a 65535"); df_exp=df_exp.head(65535)
                    df_exp.to_excel(ruta, index=False, engine='xlwt')
                except ImportError: logging.error("Falta xlwt"); messagebox.showerror("Falta xlwt", "`pip install xlwt`"); self._actualizar_estado("Error: Falta xlwt."); return
            else: messagebox.showerror("Extensión Inválida", f"{ext}"); self._actualizar_estado("Error: Extensión inválida."); return
            logging.info("Exportación OK."); messagebox.showinfo("Éxito", f"{len(df_exp)} filas exportadas:\n{ruta}")
            self._actualizar_estado(f"Resultados exportados.")
        except ImportError as ie: logging.exception("Falta lib exportar"); messagebox.showerror("Falta Librería", f"{ie}"); self._actualizar_estado("Error: Librería faltante.")
        except Exception as e: logging.exception("Error exportación"); messagebox.showerror("Error", f"{e}"); self._actualizar_estado("Error al exportar.")


    def on_closing(self): # (Sin cambios)
        """Acciones al cerrar: guardar config y destruir ventana."""
        logging.info("Cerrando aplicación y guardando configuración.")
        self._guardar_configuracion()
        self.destroy()

# --- Bloque Principal ---
if __name__ == "__main__":
    log_file = 'buscador_app.log'
    logging.basicConfig(
        level=logging.DEBUG, # DEBUG para ver todo durante desarrollo
        format='%(asctime)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s', # Formato más detallado
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8', mode='w'), # 'w' para sobrescribir log cada vez
            logging.StreamHandler()
        ]
    )
    logging.info("=============================================")
    logging.info("=== Iniciando Aplicación Buscador vX.Y.Z ===") # Considerar añadir versión
    logging.info(f"Plataforma: {platform.system()} {platform.release()}")
    logging.info(f"Python: {platform.python_version()}")
    logging.info("=============================================")

    # Comprobación dependencias
    missing = []
    try: import pandas as pd; logging.info(f"Pandas version: {pd.__version__}")
    except ImportError: missing.append("pandas")
    try: import openpyxl; logging.info(f"openpyxl version: {openpyxl.__version__}")
    except ImportError: missing.append("openpyxl")
    # xlwt es opcional, se comprueba en exportación

    if missing:
        logging.critical(f"Dependencias faltantes: {missing}")
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Dependencias Faltantes", f"Instale: {', '.join(missing)}")
        root.destroy(); exit(1)

    # Iniciar UI
    try:
        app = InterfazGrafica()
        app.mainloop()
    except Exception as main_e:
        logging.critical("Error fatal en la aplicación:", exc_info=True)
        try: # Último intento de mostrar mensaje
            root_err = tk.Tk(); root_err.withdraw()
            messagebox.showerror("Error Fatal", f"Error crítico:\n{main_e}\nConsulte '{log_file}'.")
            root_err.destroy()
        except: pass
    finally:
         logging.info("=== Finalizando Aplicación Buscador ===")