# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Tuple, Union, Set, Callable, Dict, Any
from enum import Enum, auto # Añadido para OrigenResultados
import traceback
import platform
import unicodedata
import logging
import json
import os
import string 

# --- Configuración del Logging ---
# (Se configura en el bloque __main__)

# --- Enumeraciones ---
class OrigenResultados(Enum):
    NINGUNO = 0
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto()
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto()
    DIRECTO_DESCRIPCION = auto()
    DIRECTO_DESCRIPCION_VACIA = auto() # Incluye búsqueda vacía ""

    @property
    def es_via_diccionario(self) -> bool:
        """Indica si el origen involucró una coincidencia inicial en el diccionario."""
        return self in {OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
                         OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
                         OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC}

    @property
    def es_directo_descripcion(self) -> bool:
        """Indica si el origen fue una búsqueda directa en descripciones (incluyendo vacía)."""
        return self in {OrigenResultados.DIRECTO_DESCRIPCION,
                         OrigenResultados.DIRECTO_DESCRIPCION_VACIA}


# --- Clases de Lógica ---

class ManejadorExcel:
    # (Sin cambios)
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
            messagebox.showerror("Error de Archivo", f"No se encontró el archivo:\n{ruta}\n\nVerifique que la ruta sea correcta y que el archivo exista.")
            return None
        except Exception as e:
            logging.exception(f"Error inesperado al cargar archivo: {ruta}")
            messagebox.showerror("Error al Cargar", 
                f"No se pudo cargar el archivo:\n{ruta}\n\nError: {e}\n\n"
                "Posibles causas:\n"
                "- El archivo está siendo usado por otro programa\n"
                "- No tiene instalado 'openpyxl' para archivos .xlsx\n"
                "- El archivo está corrupto o en formato no soportado")
            return None

class MotorBusqueda:
    # (Sin cambios internos)
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[str] = None
        self.archivo_descripcion_actual: Optional[str] = None
        self.indices_columnas_busqueda_dic: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else [0, 3]
        logging.info(f"MotorBusqueda inicializado. Índices búsqueda diccionario: {self.indices_columnas_busqueda_dic}")
        self.patron_comparacion_compilado = re.compile(r"^([<>]=?)(\d+([.,]\d+)?).*$")
        self.patron_rango_compilado = re.compile(r"^(\d+([.,]\d+)?)-(\d+([.,]\d+)?)$")
        self.patron_negacion_compilado = re.compile(r"^#(.+)$")
        self.extractor_magnitud = ExtractorMagnitud()  # Añadimos una instancia de ExtractorMagnitud

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
        if not self.indices_columnas_busqueda_dic:
            logging.error("La lista de índices de columnas de búsqueda está vacía en la configuración.")
            messagebox.showerror("Error de Configuración", 
                "No hay índices de columna definidos para la búsqueda en el diccionario.\n\n"
                "Por favor, configure los índices de las columnas que desea utilizar para la búsqueda.")
            return False
        max_indice_requerido = max(self.indices_columnas_busqueda_dic) if self.indices_columnas_busqueda_dic else -1

        if num_cols == 0:
            logging.error("Diccionario sin columnas.")
            messagebox.showerror("Error de Diccionario", 
                "El archivo del diccionario está vacío o no contiene columnas.\n\n"
                "Verifique que el archivo:\n"
                "- No esté vacío\n"
                "- Tenga al menos una columna de datos\n"
                "- Esté en formato Excel válido")
            return False
        elif max_indice_requerido >= num_cols:
            logging.error(f"Diccionario tiene {num_cols} cols, necesita índice {max_indice_requerido} (es decir, al menos {max_indice_requerido+1} columnas).")
            messagebox.showerror("Error de Diccionario", 
                f"El diccionario necesita al menos {max_indice_requerido + 1} columnas para los índices configurados ({self.indices_columnas_busqueda_dic}), "
                f"pero solo tiene {num_cols}.\n\n"
                "Por favor, verifique que:\n"
                "- El archivo tiene suficientes columnas\n"
                "- Los índices configurados son correctos")
            return False
        return True

    def _obtener_nombres_columnas_busqueda(self, df: pd.DataFrame) -> Optional[List[str]]:
        if df is None: 
            logging.error("Intento obtener cols de DataFrame nulo.")
            return None
        columnas_disponibles = df.columns
        cols_encontradas_nombres = []
        num_cols_df = len(columnas_disponibles)
        indices_validos = []
        for indice in self.indices_columnas_busqueda_dic:
            if isinstance(indice, int) and 0 <= indice < num_cols_df:
                cols_encontradas_nombres.append(columnas_disponibles[indice])
                indices_validos.append(indice)
            else:
                logging.warning(f"Índice {indice} inválido o fuera de rango (0-{num_cols_df-1}). Se omitirá.")
        if not cols_encontradas_nombres:
            logging.error(f"No se encontraron columnas válidas para índices: {self.indices_columnas_busqueda_dic}")
            messagebox.showerror("Error de Configuración", 
                f"No hay columnas válidas para los índices configurados: {self.indices_columnas_busqueda_dic}\n\n"
                "Por favor, verifique que:\n"
                "- Los índices configurados son correctos\n"
                "- Las columnas existen en el archivo\n"
                "- Los índices están dentro del rango válido (0-{num_cols_df-1})")
            return None
        logging.debug(f"Columnas búsqueda diccionario: {cols_encontradas_nombres} (Índices: {indices_validos})")
        return cols_encontradas_nombres

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, cols_nombres: List[str]) -> Set[str]:
        # (Sin cambios)
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
        # (Sin cambios)
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

    def _parse_numero(self, num_str: str) -> Union[int, float, None]:
        # (Sin cambios)
        if not isinstance(num_str, str): return None
        try: return float(num_str.replace(',', '.'))
        except ValueError: return None

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        # (Sin cambios)
        palabras_analizadas = []; patron_comp = self.patron_comparacion_compilado
        patron_rango = self.patron_rango_compilado; patron_neg = self.patron_negacion_compilado
        for term_orig in terminos_brutos:
            term = term_orig.strip(); negate = False; item = {'original': term_orig}
            if not term: continue
            match_neg = patron_neg.match(term)
            if match_neg: negate = True; term = match_neg.group(1).strip()
            if not term: continue
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

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols_validas: List[str], termino_analizado: Dict[str, Any]) -> pd.Series:
        mask_item_total = pd.Series(False, index=df.index)
        tipo, valor = termino_analizado['tipo'], termino_analizado['valor']

        for col_n in cols_validas:
            col = df[col_n]
            mask_col_item = pd.Series(False, index=df.index)
            try:
                if tipo == 'str':
                    mask_col_item = col.astype(str).str.contains(r"\b"+re.escape(str(valor))+r"\b", case=False, na=False, regex=True)
                elif tipo in ['gt', 'lt', 'ge', 'le']:
                    # Convertir la columna a string para buscar patrones
                    col_str = col.astype(str)
                    
                    # Extraer el valor numérico y la unidad/palabra asociada
                    valor_str = str(valor)
                    match = re.match(r'(\d+(?:[.,]\d+)?)([a-zA-ZáéíóúÁÉÍÓÚñÑ\s]+)?', valor_str)
                    
                    if match:
                        num_valor = float(match.group(1).replace(',', '.'))
                        unidad = match.group(2).strip() if match.group(2) else None
                        
                        # Normalizar la unidad si existe
                        if unidad:
                            unidad = self.extractor_magnitud._quitar_diacronicos_y_acentos(unidad.upper())
                        
                        # Buscar patrones numéricos en la columna
                        patrones_numeros = re.finditer(r'(\d+(?:[.,]\d+)?)([a-zA-ZáéíóúÁÉÍÓÚñÑ\s]+)?', col_str)
                        
                        for match_num in patrones_numeros:
                            num_col = float(match_num.group(1).replace(',', '.'))
                            unidad_col = match_num.group(2).strip() if match_num.group(2) else None
                            
                            # Normalizar la unidad de la columna si existe
                            if unidad_col:
                                unidad_col = self.extractor_magnitud._quitar_diacronicos_y_acentos(unidad_col.upper())
                            
                            # Verificar la condición numérica
                            cumple_condicion = False
                            if tipo == 'gt': cumple_condicion = num_col > num_valor
                            elif tipo == 'lt': cumple_condicion = num_col < num_valor
                            elif tipo == 'ge': cumple_condicion = num_col >= num_valor
                            else: cumple_condicion = num_col <= num_valor  # le
                            
                            if cumple_condicion:
                                # Si hay unidad especificada, verificar que coincida
                                if unidad:
                                    if unidad_col:
                                        # Verificar si la unidad está en la lista de magnitudes predefinidas
                                        es_magnitud_predefinida = unidad in self.extractor_magnitud.magnitudes
                                        es_magnitud_col_predefinida = unidad_col in self.extractor_magnitud.magnitudes
                                        
                                        # Si la unidad es una sola letra o está en la lista predefinida
                                        if len(unidad) == 1 or es_magnitud_predefinida:
                                            if unidad.lower() == unidad_col.lower():
                                                mask_col_item |= col_str.str.contains(r"\b"+re.escape(match_num.group(0))+r"\b", case=False, na=False, regex=True)
                                        else:
                                            # Si es una palabra, buscar que contenga la palabra
                                            if unidad.lower() in unidad_col.lower():
                                                mask_col_item |= col_str.str.contains(r"\b"+re.escape(match_num.group(0))+r"\b", case=False, na=False, regex=True)
                                else:
                                    # Si no hay unidad especificada, solo verificar el número
                                    mask_col_item |= col_str.str.contains(r"\b"+re.escape(match_num.group(0))+r"\b", case=False, na=False, regex=True)
                    else:
                        # Si no hay unidad, comportamiento original
                        col_num = pd.to_numeric(col, errors='coerce')
                        if tipo == 'gt': mask_col_item = col_num > num_valor
                        elif tipo == 'lt': mask_col_item = col_num < num_valor
                        elif tipo == 'ge': mask_col_item = col_num >= num_valor
                        else: mask_col_item = col_num <= num_valor  # le
                        mask_col_item = mask_col_item.fillna(False)
                
                elif tipo == 'range':
                    min_v, max_v = valor
                    col_num = pd.to_numeric(col, errors='coerce')
                    mask_col_item = (col_num >= min_v) & (col_num <= max_v)
                    mask_col_item = mask_col_item.fillna(False)
                
                mask_item_total |= mask_col_item
            except Exception as e:
                logging.warning(f"Error procesando término '{termino_analizado.get('original', 'N/A')}' en columna '{col_n}': {e}")
        return mask_item_total

    def _aplicar_mascara_diccionario(self, df: pd.DataFrame, cols_nombres: List[str], terms_analizados: List[Dict[str, Any]], op_principal: str) -> pd.Series:
        # (Sin cambios)
        if df is None or df.empty or not cols_nombres or not terms_analizados:
            return pd.Series(False, index=df.index if df is not None else None)

        cols_ok = [c for c in cols_nombres if c in df.columns]
        if not cols_ok:
            logging.error(f"Ninguna columna de {cols_nombres} es válida en el DataFrame.")
            return pd.Series(False, index=df.index)

        terms_pos = [item for item in terms_analizados if not item.get('negate', False)]
        terms_neg = [item for item in terms_analizados if item.get('negate', False)]

        op_es_and = op_principal.upper() == 'AND'
        if not terms_pos:
            mask_pos_final = pd.Series(op_es_and, index=df.index)
        else:
            mask_pos_final = pd.Series(True, index=df.index) if op_es_and else pd.Series(False, index=df.index)
            for item in terms_pos:
                mask_item = self._generar_mascara_para_un_termino(df, cols_ok, item)
                if op_es_and: mask_pos_final &= mask_item
                else: mask_pos_final |= mask_item

        mask_neg_combinada = pd.Series(False, index=df.index)
        if terms_neg:
            for item in terms_neg:
                mask_item_neg = self._generar_mascara_para_un_termino(df, cols_ok, item)
                mask_neg_combinada |= mask_item_neg

        mask_final = mask_pos_final & (~mask_neg_combinada)

        logging.debug(f"Mask Pos Final: {mask_pos_final.sum()}, Mask Neg Combinada: {mask_neg_combinada.sum()}, Mask Final: {mask_final.sum()}")
        return mask_final

    def buscar_en_descripciones_directo(self, term_buscar: str) -> pd.DataFrame:
        logging.info(f"Búsqueda Directa en Descripciones: '{term_buscar}'")
        if self.datos_descripcion is None or self.datos_descripcion.empty:
            logging.warning("Intento de búsqueda directa sin descripciones cargadas o vacías.")
            messagebox.showwarning("Datos Faltantes", 
                "No hay datos de descripciones cargados para realizar la búsqueda directa.\n\n"
                "Por favor, cargue primero el archivo de descripciones.")
            return pd.DataFrame()

        term_limpio = term_buscar.strip().upper()
        if not term_limpio:
            logging.info("Término de búsqueda directa vacío. Devolviendo todas las descripciones.")
            return self.datos_descripcion.copy()

        df_desc = self.datos_descripcion.copy()
        res = pd.DataFrame(columns=df_desc.columns)
        try:
            txt_filas = df_desc.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            mask = pd.Series(False, index=df_desc.index)

            if '+' in term_limpio:
                palabras = [p.strip() for p in term_limpio.split('+') if p.strip()]
                if not palabras: return res
                mask = pd.Series(True, index=df_desc.index)
                for p in palabras: mask &= txt_filas.str.contains(r"\b"+re.escape(p)+r"\b", regex=True, na=False)
            elif '|' in term_limpio or '/' in term_limpio:
                sep = '|' if '|' in term_limpio else '/'
                palabras = [p.strip() for p in term_limpio.split(sep) if p.strip()]
                if not palabras: return res
                for p in palabras: mask |= txt_filas.str.contains(r"\b"+re.escape(p)+r"\b", regex=True, na=False)
            else: mask = txt_filas.str.contains(r"\b"+re.escape(term_limpio)+r"\b", regex=True, na=False)

            res = df_desc[mask]
            logging.info(f"Búsqueda directa completada. Resultados: {len(res)}.")
        except Exception as e:
            logging.exception("Error durante la búsqueda directa en descripciones.")
            messagebox.showerror("Error de Búsqueda", 
                f"Ocurrió un error durante la búsqueda directa:\n{e}\n\n"
                "Por favor, intente nuevamente o contacte al soporte técnico si el problema persiste.")
        return res

# --- Clase ExtractorMagnitud ---
# (Se mantiene igual, omitida por brevedad)
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
        self.title("Buscador Avanzado (con Salvar Regla)")
        self.geometry("1250x800")

        self.config = self._cargar_configuracion()
        indices_cfg = self.config.get("indices_columnas_busqueda_dic", [0, 3])
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg)
        self.extractor_magnitud = ExtractorMagnitud()

        self.resultados_actuales: Optional[pd.DataFrame] = None
        
        # >>> INICIO: Usar StringVar para entrada_busqueda y vincular trace <<<
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        # <<< FIN: Usar StringVar para entrada_busqueda y vincular trace <<<
        
        self.ultimo_termino_buscado: Optional[str] = None

        self.reglas_guardadas: List[Dict[str, Any]] = []
        self.df_candidato_diccionario: Optional[pd.DataFrame] = None
        self.df_candidato_descripcion: Optional[pd.DataFrame] = None
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO

        self.color_fila_par = "white"; self.color_fila_impar = "#f0f0f0"

        self._configurar_estilo_ttk()
        self._crear_widgets()
        self._configurar_grid()
        self._configurar_eventos()
        self._configurar_tags_treeview()
        self._configurar_orden_tabla(self.tabla_resultados)
        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones.")
        
        # Deshabilitar todos los botones operacionales al inicio
        self._deshabilitar_botones_operadores()
        
        # Actualizar el estado general de los botones
        self._actualizar_botones_estado_general()
        
        logging.info("Interfaz Gráfica inicializada.")

    # >>> INICIO: Definición del callback _on_texto_busqueda_change <<<
    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str):
        """Callback que se ejecuta cuando el contenido de texto_busqueda_var cambia."""
        # Actualizar estado de botones operadores
        self._actualizar_estado_botones_operadores()
        
        # Obtener el texto actual y verificar si está vacío
        texto_actual = self.texto_busqueda_var.get().strip()
        
        # Si el texto está vacío y tenemos los archivos cargados, mostrar todas las descripciones
        if not texto_actual and self.motor.datos_diccionario is not None and self.motor.datos_descripcion is not None:
            # Resetear resultados previos
            self.resultados_actuales = None
            self.df_candidato_diccionario = None
            self.df_candidato_descripcion = None
            self.origen_principal_resultados = OrigenResultados.NINGUNO
            
            # Mostrar todas las descripciones
            df_desc_all = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc_all)
            self.resultados_actuales = df_desc_all.copy() if df_desc_all is not None else None
            self.df_candidato_descripcion = self.resultados_actuales
            self.origen_principal_resultados = OrigenResultados.DIRECTO_DESCRIPCION_VACIA
            
            # Mostrar vista previa del diccionario
            df_dic_preview = self.motor.datos_diccionario
            if df_dic_preview is not None:
                cols_preview = self.motor._obtener_nombres_columnas_busqueda(df_dic_preview)
                self._actualizar_tabla(self.tabla_diccionario, df_dic_preview, limite_filas=100, columnas_a_mostrar=cols_preview)
            
            num_filas = len(df_desc_all) if df_desc_all is not None else 0
            self._actualizar_estado(f"Mostrando todas las {num_filas} descripciones.")
            self._actualizar_botones_estado_general()
        # Si hay texto y los archivos están cargados, ejecutar la búsqueda automáticamente
        elif texto_actual and self.motor.datos_diccionario is not None and self.motor.datos_descripcion is not None:
            # Usar after para evitar múltiples búsquedas mientras el usuario escribe
            if hasattr(self, '_busqueda_pendiente'):
                self.after_cancel(self._busqueda_pendiente)
            self._busqueda_pendiente = self.after(500, self._ejecutar_busqueda)  # Esperar 500ms antes de buscar

    def _cargar_configuracion(self) -> Dict:
        # (Sin cambios)
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
        config.setdefault("last_dic_path", None)
        config.setdefault("last_desc_path", None)
        config.setdefault("indices_columnas_busqueda_dic", [0, 3])
        return config

    def _guardar_configuracion(self):
        # (Sin cambios)
        self.config["last_dic_path"] = self.motor.archivo_diccionario_actual
        self.config["last_desc_path"] = self.motor.archivo_descripcion_actual
        self.config["indices_columnas_busqueda_dic"] = self.motor.indices_columnas_busqueda_dic
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            logging.info(f"Configuración guardada en: {self.CONFIG_FILE}")
        except Exception as e:
            logging.error(f"Error al guardar la configuración en '{self.CONFIG_FILE}': {e}")
            messagebox.showerror("Error Configuración", f"No se pudo guardar la configuración:\n{e}")

    def _configurar_estilo_ttk(self):
        # (Sin cambios)
        style = ttk.Style(self); themes = style.theme_names(); os_name = platform.system()
        prefs = {"Windows":["vista","xpnative","clam"],"Darwin":["aqua","clam"],"Linux":["clam","alt","default"]}
        theme_to_use = next((t for t in prefs.get(os_name, ["clam","default"]) if t in themes), None)
        if not theme_to_use:
            theme_to_use = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))
        if theme_to_use:
            logging.info(f"Aplicando tema TTK: {theme_to_use}")
            try: style.theme_use(theme_to_use)
            except tk.TclError as e: logging.warning(f"No se pudo aplicar el tema '{theme_to_use}': {e}. Usando tema por defecto.")
        else: logging.warning("No se encontró ningún tema TTK disponible.")

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
        self.marco_controles.grid_columnconfigure(3, weight=1)
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew")

        self.frame_ops.grid(row=1, column=0, columnspan=6, padx=5, pady=(5,0), sticky="w")

        self.entrada_busqueda.grid(row=2, column=0, columnspan=3, padx=5, pady=(0,5), sticky="ew")
        self.btn_salvar_regla.grid(row=2, column=3, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_ayuda.grid(row=2, column=4, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_exportar.grid(row=2, column=5, padx=(10, 5), pady=(0,5), sticky="e")

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
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _actualizar_estado(self, mensaje: str):
        # (Sin cambios)
        self.barra_estado.config(text=mensaje)
        logging.info(f"Estado UI: {mensaje}")
        self.update_idletasks()

    def _mostrar_ayuda(self):
        ayuda = """Sintaxis de Búsqueda en Diccionario:
-------------------------------------
- Texto simple: Busca la palabra o frase exacta (insensible a mayús/minús).
  Ej: `router cisco`

- Operadores Lógicos:
  * `término1 + término2`: Busca filas que contengan AMBOS términos (AND).
    Ej: `tarjeta + 16 puertos`
  * `término1 | término2`: Busca filas que contengan AL MENOS UNO de los términos (OR).
    Ej: `modulo | SFP`
  * `término1 / término2`: Alternativa para OR.
    Ej: `switch / conmutador`

- Comparaciones numéricas (aplican a columnas configuradas si son numéricas):
  * `>numero`: Mayor que. Ej: `>1000` o `>1000w`
  * `<numero`: Menor que. Ej: `<50` o `<50v`
  * `>=numero` o `≥numero`: Mayor o igual que. Ej: `>=48` o `>=48a`
  * `<=numero` o `≤numero`: Menor o igual que. Ej: `<=10.5` o `<=10.5w`
  * Los operadores de comparación pueden usarse al inicio de la búsqueda
  * Pueden ir seguidos de una letra o unidad (ej: w, v, a)
  * Si la unidad es una letra, debe estar pegada al número (ej: >1000w)
  * Si la unidad es una palabra, puede estar separada por espacio (ej: >1000 vatios)
  * La búsqueda es insensible a mayúsculas/minúsculas para las unidades

- Rangos numéricos (ambos incluidos):
  * `num1-num2`: Entre num1 y num2. Ej: `10-20` (buscará 10, 11, ..., 20)
  * El número antes del guión debe ser un dígito
  * Requiere al menos una palabra antes de usar el operador de rango

- Negación (excluir filas):
  * `#término`: Excluye filas que coincidan con 'término'.
    'término' puede ser texto, comparación o rango.
    Ej: `switch + #gestionable` (busca 'switch' pero no los que contengan 'gestionable')
    Ej: `tarjeta + #>8` (busca 'tarjeta' pero no las que tengan número > 8)

Restricciones y Reglas:
---------------------
1. No se permiten operadores duplicados (ej: `++`, `||`)
2. La negación (#) solo se puede usar al inicio de un término
3. Los operadores de comparación (> < >= <=) no se pueden combinar en un mismo término
4. El operador de rango (-) requiere un número antes del guión
5. Los operadores lógicos (+ | /) requieren un término antes y después
6. No se permiten espacios entre operadores y números (ej: `> 10` es inválido)
7. El operador de rango (-) requiere al menos una palabra antes de usarlo
8. Para unidades de una letra, debe estar pegada al número (ej: >1000w)
9. Para palabras como unidades, puede haber espacio (ej: >1000 vatios)

Notas sobre la Negación:
----------------------
- La negación (#) debe colocarse al inicio del término que se desea excluir
- Se aplica después de evaluar las condiciones positivas
- Si un término negado coincide en cualquier columna configurada, se excluye la fila
- Se puede combinar con otros operadores: `término1 + #término2 | #término3`

Búsqueda Directa (si el término no está en diccionario):
------------------------------------------------------
Se busca directamente en las descripciones:
- `texto`: Busca el texto.
- `t1 + t2`: Busca descripciones con AMBOS términos.
- `t1 | t2` o `t1 / t2`: Busca descripciones con AL MENOS UNO.

Notas:
- La búsqueda es insensible a mayúsculas/minúsculas y acentos
- Los términos se buscan como palabras completas (no subcadenas)
- Se ignoran espacios extra entre términos
"""
        messagebox.showinfo("Ayuda - Sintaxis de Búsqueda", ayuda)

    def _configurar_tags_treeview(self):
        # (Sin cambios)
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par)
            tabla.tag_configure('impar', background=self.color_fila_impar)

    def _configurar_orden_tabla(self, tabla: ttk.Treeview):
        # (Sin cambios)
        cols = tabla["columns"]
        if cols:
            for col in cols:
                tabla.heading(col, text=col, anchor=tk.W,
                              command=lambda c=col: self._ordenar_columna(tabla, c, False))

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool):
        # (Sin cambios)
        df_para_ordenar = self.resultados_actuales
        if tabla != self.tabla_resultados or df_para_ordenar is None or df_para_ordenar.empty:
            logging.debug("Intento de ordenar tabla de resultados vacía o no aplicable.")
            return

        logging.info(f"Ordenando resultados por columna '{col}', descendente={reverse}")
        try:
            # Convertir a numérico para ordenar, pero mantener tipos originales para mostrar si es posible
            # Si la columna es puramente numérica (o puede serlo), se ordena numéricamente.
            # Si no, se ordena lexicográficamente.
            # Pandas intentará la conversión numérica con 'coerce', los no numéricos serán NaN.
            df_ordenado = df_para_ordenar.sort_values(
                by=col, ascending=not reverse, na_position='last',
                key=lambda x: pd.to_numeric(x, errors='coerce') if x.dtype == object or pd.api.types.is_string_dtype(x) else x
            )
            self.resultados_actuales = df_ordenado
            self._actualizar_tabla(tabla, self.resultados_actuales)
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, not reverse))
            self._actualizar_estado(f"Resultados ordenados por '{col}' ({'Asc' if not reverse else 'Desc'}).")
        except Exception as e:
            logging.exception(f"Error al intentar ordenar por columna '{col}'")
            messagebox.showerror("Error al Ordenar", f"No se pudo ordenar por '{col}':\n{e}")
            # Resetear el comando de ordenación para evitar bucles de error
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, False))


    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None, columnas_a_mostrar: Optional[List[str]] = None):
        # (Sin cambios respecto a la versión anterior del script)
        logging.debug(f"Actualizando tabla {'Diccionario' if tabla == self.tabla_diccionario else 'Resultados'}.")
        try:
            for i in tabla.get_children(): tabla.delete(i)
        except tk.TclError as e: logging.warning(f"Error Tcl al limpiar tabla: {e}"); pass
        tabla["columns"] = ()

        if datos is None or datos.empty:
            logging.debug("No hay datos para mostrar en la tabla.")
            return

        cols_originales = list(datos.columns)
        if columnas_a_mostrar:
            cols_ok = [c for c in columnas_a_mostrar if c in cols_originales]
            if not cols_ok:
                logging.warning(f"Columnas especificadas {columnas_a_mostrar} no encontradas en datos. Mostrando todas.")
                cols_ok = cols_originales
        else:
            cols_ok = cols_originales
        
        if not cols_ok:
            logging.warning("DataFrame no tiene columnas para mostrar.")
            return

        df_mostrar = datos[cols_ok]
        tabla["columns"] = cols_ok
        for col in cols_ok:
            tabla.heading(col, text=str(col), anchor=tk.W)
            try:
                muestra = min(len(df_mostrar), 100)
                if col in df_mostrar.columns:
                    sub_df = df_mostrar.iloc[:muestra][col].dropna().astype(str)
                    ancho_contenido = sub_df.str.len().max() if not sub_df.empty else 0
                else:
                    ancho_contenido = 0
                    logging.warning(f"Columna '{col}' inesperadamente ausente al calcular ancho.")
                ancho_cabecera = len(str(col))
                ancho = max(70, min(int(max(ancho_cabecera * 9, ancho_contenido * 7) + 20), 400))
                tabla.column(col, anchor=tk.W, width=ancho, minwidth=70)
            except Exception as e:
                logging.warning(f"Error calculando ancho para columna '{col}': {e}. Usando ancho por defecto.")
                tabla.column(col, anchor=tk.W, width=100, minwidth=50)

        df_final = df_mostrar.head(limite_filas) if limite_filas is not None and len(df_mostrar) > limite_filas else df_mostrar
        logging.debug(f"Mostrando {len(df_final)} filas.")
        for i, (_, row) in enumerate(df_final.iterrows()):
            vals = [str(v) if pd.notna(v) else "" for v in row.values]
            tag = 'par' if i % 2 == 0 else 'impar'
            try: tabla.insert("", "end", values=vals, tags=(tag,))
            except tk.TclError as e:
                logging.warning(f"Error Tcl insertando fila {i}: {e}. Intentando con ASCII.")
                try:
                    vals_ascii = [v.encode('ascii', 'ignore').decode('ascii') for v in vals]
                    tabla.insert("", "end", values=vals_ascii, tags=(tag,))
                except Exception as e_inner: logging.error(f"Fallo el fallback ASCII para fila {i}: {e_inner}")
        if tabla == self.tabla_resultados: # Solo configurar orden para tabla de resultados
            self._configurar_orden_tabla(tabla)


    def _actualizar_etiquetas_archivos(self):
        # (Sin cambios)
        dic_name = os.path.basename(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else "Ninguno"
        desc_name = os.path.basename(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else "Ninguno"
        max_len_label = 25 
        dic_display = f"Dic: {dic_name}" if len(dic_name) <= max_len_label else f"Dic: ...{dic_name[-(max_len_label-4):]}"
        desc_display = f"Desc: {desc_name}" if len(desc_name) <= max_len_label else f"Desc: ...{desc_name[-(max_len_label-4):]}"
        self.lbl_dic_cargado.config(text=dic_display)
        self.lbl_desc_cargado.config(text=desc_display)

    def _actualizar_botones_estado_general(self):
        """Actualiza el estado de los botones basado en el estado general de la aplicación."""
        dic_cargado = self.motor.datos_diccionario is not None
        desc_cargado = self.motor.datos_descripcion is not None

        # Primero actualizar el estado base de los botones de operadores según el diccionario
        estado_base_operadores = 'normal' if dic_cargado else 'disabled'
        
        # Establecer el estado base de los botones de operadores
        self.btn_and['state'] = estado_base_operadores
        self.btn_or['state'] = estado_base_operadores
        self.btn_not['state'] = estado_base_operadores
        self.btn_gt['state'] = estado_base_operadores
        self.btn_lt['state'] = estado_base_operadores
        self.btn_ge['state'] = estado_base_operadores
        self.btn_le['state'] = estado_base_operadores
        self.btn_range['state'] = estado_base_operadores

        # Solo aplicar las validaciones específicas de los operadores si el diccionario está cargado
        if dic_cargado:
            self._actualizar_estado_botones_operadores()
        else:
            # Si no hay diccionario cargado, asegurarse de que todos los botones operacionales estén deshabilitados
            self._deshabilitar_botones_operadores()

        # Estado de otros botones
        self.btn_cargar_descripciones['state'] = 'normal' if dic_cargado else 'disabled'
        
        # Lógica botón Salvar Regla
        puede_salvar_fcd = self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty
        puede_salvar_rfd_o_rdd = self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty
        estado_salvar = 'disabled'

        if self.origen_principal_resultados != OrigenResultados.NINGUNO:
            if self.origen_principal_resultados.es_directo_descripcion: # Incluye búsqueda vacía
                if puede_salvar_rfd_o_rdd: estado_salvar = 'normal'
            elif self.origen_principal_resultados.es_via_diccionario:
                # Se puede salvar si hay FCD, o si hay RFD (y el origen es VIA_DICCIONARIO_CON_RESULTADOS_DESC)
                if puede_salvar_fcd or \
                   (self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC and puede_salvar_rfd_o_rdd):
                    estado_salvar = 'normal'
        self.btn_salvar_regla['state'] = estado_salvar
        
        self.btn_exportar['state'] = 'normal' if self.reglas_guardadas else 'disabled'

    def _cargar_diccionario(self):
        """Carga el archivo de diccionario y actualiza la interfaz."""
        
        try:
            # Siempre mantener el botón de cargar diccionario habilitado
            self.btn_cargar_diccionario['state'] = 'normal'
            
            last_dir = os.path.dirname(self.config.get("last_dic_path", "") or "") or None
            ruta = filedialog.askopenfilename(title="Seleccionar Archivo Diccionario", filetypes=[("Archivos Excel", "*.xlsx *.xls")], initialdir=last_dir)
            if not ruta: logging.info("Carga de diccionario cancelada."); return

            nombre_archivo = os.path.basename(ruta)
            self._actualizar_estado(f"Cargando diccionario: {nombre_archivo}...")
            self._actualizar_tabla(self.tabla_diccionario, None) 
            self.resultados_actuales = None
            self.df_candidato_diccionario = None
            self.df_candidato_descripcion = None
            self.origen_principal_resultados = OrigenResultados.NINGUNO
            self._actualizar_tabla(self.tabla_resultados, None) 

            if self.motor.cargar_excel_diccionario(ruta):
                self._guardar_configuracion() 
                df_dic = self.motor.datos_diccionario
                if df_dic is not None:
                    num_filas = len(df_dic)
                    cols_busqueda_nombres = self.motor._obtener_nombres_columnas_busqueda(df_dic)
                    indices_str = ', '.join(map(str, self.motor.indices_columnas_busqueda_dic))
                    lbl_text = f"Vista Previa Diccionario (Índices: {indices_str})"
                    if cols_busqueda_nombres: lbl_text = f"Vista Previa Dic ({', '.join(cols_busqueda_nombres)} - Índices: {indices_str})"
                    self.lbl_tabla_diccionario.config(text=lbl_text)
                    self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols_busqueda_nombres)
                    self.title(f"Buscador - Dic: {nombre_archivo}")
                    self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado.")
            else:
                self._actualizar_estado("Error al cargar el diccionario.")
                self.title("Buscador Avanzado (con Salvar Regla)")

            self._actualizar_etiquetas_archivos()
            self._actualizar_botones_estado_general()
        except Exception as e:
            logging.error(f"Error al cargar el diccionario: {e}")
            messagebox.showerror("Error al Cargar Diccionario", f"No se pudo cargar el archivo del diccionario:\n{e}")

    def _cargar_excel_descripcion(self):
        last_dir = os.path.dirname(self.config.get("last_desc_path", "") or "") or None
        ruta = filedialog.askopenfilename(title="Seleccionar Archivo de Descripciones", filetypes=[("Archivos Excel", "*.xlsx *.xls")], initialdir=last_dir)
        if not ruta: logging.info("Carga de descripciones cancelada."); return

        nombre_archivo = os.path.basename(ruta)
        self._actualizar_estado(f"Cargando descripciones: {nombre_archivo}...")
        # Resetear resultados previos al cargar nuevas descripciones
        self.resultados_actuales = None
        self.df_candidato_diccionario = None
        self.df_candidato_descripcion = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO
        # No es necesario limpiar la tabla aquí, _actualizar_tabla lo hará.

        if self.motor.cargar_excel_descripcion(ruta):
            self._guardar_configuracion()
            df_desc = self.motor.datos_descripcion
            if df_desc is not None:
                num_filas = len(df_desc)
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando datos...")
                self._actualizar_tabla(self.tabla_resultados, df_desc) # Mostrar todas las descripciones

                # Actualizar título de la ventana
                if self.motor.archivo_diccionario_actual:
                    dic_n = os.path.basename(self.motor.archivo_diccionario_actual)
                    self.title(f"Buscador - Dic: {dic_n} | Desc: {nombre_archivo}")
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas.")
        else:
            self._actualizar_estado("Error al cargar las descripciones.")
        
        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _buscar_y_enfocar_en_preview(self):
        # (Sin cambios)
        termino_buscar = self.texto_busqueda_var.get().strip() # Usar StringVar
        if not termino_buscar: return
        items_preview = self.tabla_diccionario.get_children('')
        if not items_preview: return

        termino_upper = termino_buscar.upper()
        logging.info(f"Buscando '{termino_buscar}' en la vista previa del diccionario...")

        found_item_id = None
        for item_id in items_preview:
            try:
                valores_fila = self.tabla_diccionario.item(item_id, 'values')
                if any(termino_upper in str(val).upper() for val in valores_fila):
                    found_item_id = item_id; break
            except Exception as e: logging.warning(f"Error procesando item {item_id} en preview: {e}"); continue
        
        if found_item_id:
            logging.info(f"Término '{termino_buscar}' encontrado en preview (item ID: {found_item_id}).")
            try:
                current_selection = self.tabla_diccionario.selection()
                if current_selection: self.tabla_diccionario.selection_remove(current_selection)
                self.tabla_diccionario.selection_set(found_item_id)
                self.tabla_diccionario.see(found_item_id)
            except Exception as e: logging.error(f"Error al enfocar item {found_item_id} en preview: {e}")
        else:
            logging.info(f"Término '{termino_buscar}' no encontrado en vista previa del diccionario.")


    def _parsear_termino_busqueda_inicial(self, termino_raw: str) -> Tuple[str, List[Dict[str, Any]]]:
        """
        Parsea el término de búsqueda inicial y retorna el término original y los términos analizados.
        
        Args:
            termino_raw (str): El término de búsqueda en bruto.
            
        Returns:
            Tuple[str, List[Dict[str, Any]]]: Una tupla con:
                - El término original
                - Lista de términos analizados con sus operadores y valores
                
        Raises:
            ValueError: Si el término no puede ser parseado correctamente.
        """
        if not isinstance(termino_raw, str):
            raise ValueError("El término de búsqueda debe ser una cadena de texto")
            
        termino_limpio = termino_raw.strip()
        if not termino_limpio:
            return termino_raw, []
            
        try:
            # Determinar el operador principal y separar los términos
            op_principal = 'OR'
            terminos_brutos = []
            
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
                terminos_brutos = [termino_limpio]
                
            if not terminos_brutos:
                logging.warning(f"Término '{termino_raw}' vacío tras parseo.")
                return termino_raw, []
                
            # Analizar cada término
            terminos_analizados = []
            for term in terminos_brutos:
                term_analizado = {'operador': op_principal, 'valor': term}
                terminos_analizados.append(term_analizado)
                
            return termino_raw, terminos_analizados
            
        except Exception as e:
            logging.error(f"Error parseando término '{termino_raw}': {str(e)}")
            raise ValueError(f"No se pudo parsear el término de búsqueda: {str(e)}")

    def _procesar_busqueda_via_diccionario(self, termino_original: str, terminos_analizados: List[Dict[str, Any]], op_principal: str,
                                         terminos_brutos: List[str]) -> Tuple[OrigenResultados, pd.DataFrame]:
        """
        Procesa la búsqueda utilizando el diccionario como fuente de datos.
        
        Args:
            termino_original (str): El término de búsqueda original.
            terminos_analizados (List[Dict[str, Any]]): Lista de términos analizados con sus operadores.
            op_principal (str): Operador principal de la búsqueda ('AND' u 'OR').
            terminos_brutos (List[str]): Lista de términos en bruto.
            
        Returns:
            Tuple[OrigenResultados, pd.DataFrame]: Una tupla con:
                - El origen de los resultados
                - DataFrame con los resultados de la búsqueda
                
        Raises:
            ValueError: Si hay un error en el procesamiento de la búsqueda.
        """
        try:
            if not terminos_analizados:
                logging.warning("No hay términos para procesar en la búsqueda")
                return OrigenResultados.DICCIONARIO, pd.DataFrame()
                
            # Extraer términos del diccionario
            terminos_diccionario = self._extraer_terminos_diccionario(terminos_brutos)
            if not terminos_diccionario:
                logging.warning("No se encontraron términos en el diccionario")
                return OrigenResultados.DICCIONARIO, pd.DataFrame()
                
            # Buscar términos en las descripciones
            resultados = self._buscar_terminos_en_descripciones(terminos_diccionario, op_principal)
            if resultados.empty:
                logging.info("No se encontraron resultados en las descripciones")
                return OrigenResultados.DICCIONARIO, pd.DataFrame()
                
            # Ordenar resultados
            resultados_ordenados = self._ordenar_resultados(resultados)
            
            return OrigenResultados.DICCIONARIO, resultados_ordenados
            
        except Exception as e:
            logging.error(f"Error procesando búsqueda vía diccionario: {str(e)}")
            raise ValueError(f"Error en el procesamiento de la búsqueda: {str(e)}")

    def _procesar_busqueda_directa_descripcion(self, termino_original: str, df_desc_original: pd.DataFrame) -> Tuple[OrigenResultados, pd.DataFrame]:
        """
        Procesa la búsqueda directamente en las descripciones.
        
        Args:
            termino_original (str): El término de búsqueda original.
            df_desc_original (pd.DataFrame): DataFrame con las descripciones.
            
        Returns:
            Tuple[OrigenResultados, pd.DataFrame]: Una tupla con:
                - El origen de los resultados
                - DataFrame con los resultados de la búsqueda
                
        Raises:
            ValueError: Si hay un error en el procesamiento de la búsqueda.
        """
        try:
            if df_desc_original.empty:
                logging.warning("No hay descripciones disponibles para buscar")
                return OrigenResultados.DESCRIPCION, pd.DataFrame()
                
            # Aplicar búsqueda directa
            mascara = self.motor._aplicar_mascara_descripcion(df_desc_original, termino_original)
            if not mascara.any():
                logging.info(f"No se encontraron resultados para '{termino_original}' en las descripciones")
                return OrigenResultados.DESCRIPCION, pd.DataFrame()
                
            # Filtrar resultados
            resultados = df_desc_original[mascara].copy()
            
            # Ordenar resultados
            resultados_ordenados = self._ordenar_resultados(resultados)
            
            return OrigenResultados.DESCRIPCION, resultados_ordenados
            
        except Exception as e:
            logging.error(f"Error procesando búsqueda directa en descripciones: {str(e)}")
            raise ValueError(f"Error en el procesamiento de la búsqueda: {str(e)}")

    def _ejecutar_busqueda(self):
        """Ejecuta la búsqueda actual y actualiza la interfaz."""
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", 
                "Por favor, cargue el Diccionario y las Descripciones antes de realizar una búsqueda.\n\n"
                "El Diccionario contiene los términos de referencia y las Descripciones son los datos donde se buscará.")
            return

        texto_busqueda = self.entrada_busqueda.get().strip()
        if not texto_busqueda:
            return

        self._actualizar_estado("Ejecutando búsqueda...")
        self._deshabilitar_botones_operadores()
        
        try:
            termino_original, terminos_analizados = self._parsear_termino_busqueda_inicial(texto_busqueda)
            
            if not terminos_analizados:
                self._actualizar_estado("No se encontraron términos válidos para buscar")
                return
                
            op_principal = terminos_analizados[0].get('operador', 'AND')
            
            if self.motor.datos_diccionario is not None and not self.motor.datos_diccionario.empty:
                origen, resultados = self._procesar_busqueda_via_diccionario(termino_original, terminos_analizados, op_principal, terminos_analizados)
            else:
                origen, resultados = self._procesar_busqueda_directa_descripcion(termino_original, self.motor.datos_descripcion)
                
            self.origen_principal_resultados = origen
            self.resultados_actuales = resultados
            
            if not resultados.empty:
                self._actualizar_tabla(self.tabla_resultados, resultados)
                if logging.getLogger().getEffectiveLevel() == logging.DEBUG:
                    self._demo_extractor(resultados, origen.name)

            if self.origen_principal_resultados.es_via_diccionario and \
               self.motor.datos_diccionario is not None and not self.motor.datos_diccionario.empty:
                self._buscar_y_enfocar_en_preview()
                
        except Exception as e:
            self._actualizar_estado(f"Error en la búsqueda: {str(e)}")
            logging.error(f"Error en la búsqueda: {str(e)}", exc_info=True)
        finally:
            self._actualizar_estado_botones_operadores()
            self._actualizar_botones_estado_general()

    def _demo_extractor(self, df_res: pd.DataFrame, tipo_busqueda: str):
        # (Sin cambios)
        if df_res is None or df_res.empty or len(df_res.columns) == 0: return
        try:
            texto_primera_celda = str(df_res.iloc[0, 0]) 
            logging.info(f"--- DEMO Extractor Magnitud (desde {tipo_busqueda}) ---")
            logging.info(f"Texto analizado: '{texto_primera_celda[:100]}...'")
            encontrado = False
            for mag in self.extractor_magnitud.magnitudes:
                cantidad = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag, texto_primera_celda)
                if cantidad is not None:
                    logging.info(f"  -> Magnitud '{mag}': encontrada cantidad '{cantidad}'")
                    encontrado = True
            if not encontrado: logging.info("  (No se encontraron magnitudes predefinidas en este texto)")
            logging.info("--- FIN DEMO Extractor Magnitud ---")
        except IndexError: logging.warning(f"Error en demo extractor ({tipo_busqueda}): No se pudo acceder a la celda [0, 0].")
        except Exception as e: logging.warning(f"Error inesperado durante la demo del extractor ({tipo_busqueda}): {e}")

    def _sanitizar_nombre_archivo(self, texto: str, max_len: int = 50) -> str:
        # (Sin cambios)
        if not texto: return "resultados"
        texto_limpio = re.sub(r'[<>:"/\\|?*#]', '_', texto) # Quitar caracteres inválidos
        texto_limpio = "".join(c for c in texto_limpio if c not in string.control) # Quitar caracteres de control
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip() # Normalizar espacios
        texto_cortado = texto_limpio[:max_len] 
        texto_final = texto_cortado.rstrip('._- ') # Quitar caracteres de terminación comunes si están al final
        if not texto_final: return "resultados" # Si todo se eliminó
        return texto_final

    def _mostrar_dialogo_seleccion_salvado_via_diccionario(self) -> Dict[str, bool]:
        # (Sin cambios)
        decision = {'confirmed': False, 'save_fcd': False, 'save_rfd': False}
        choice_window = tk.Toplevel(self)
        choice_window.title("Elegir Datos a Salvar")
        choice_window.geometry("450x200") 
        choice_window.resizable(False, False)
        choice_window.transient(self) 
        choice_window.grab_set() 
        tk.Label(choice_window, text=f"Para '{self.ultimo_termino_buscado}', elija qué salvar:").pack(pady=10, padx=10)
        var_salvar_fcd = tk.BooleanVar(value=False)
        var_salvar_rfd = tk.BooleanVar(value=False)
        frame_checkboxes = ttk.Frame(choice_window)
        frame_checkboxes.pack(fill=tk.X, padx=20)
        puede_salvar_fcd = self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty
        puede_salvar_rfd = self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty and \
                           self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
        chk_fcd_widget = ttk.Checkbutton(frame_checkboxes, text=f"Coincidencias del Diccionario ({len(self.df_candidato_diccionario or [])} filas)", variable=var_salvar_fcd, state="normal" if puede_salvar_fcd else "disabled")
        chk_fcd_widget.pack(anchor=tk.W, pady=2)
        chk_rfd_widget = ttk.Checkbutton(frame_checkboxes, text=f"Resultados en Descripciones (vía Dic, {len(self.df_candidato_descripcion or [])} filas)", variable=var_salvar_rfd, state="normal" if puede_salvar_rfd else "disabled")
        chk_rfd_widget.pack(anchor=tk.W, pady=2)
        # Preseleccionar y deshabilitar si solo hay una opción viable
        if puede_salvar_fcd and not puede_salvar_rfd: var_salvar_fcd.set(True); chk_fcd_widget.configure(state="disabled")
        elif not puede_salvar_fcd and puede_salvar_rfd: var_salvar_rfd.set(True); chk_rfd_widget.configure(state="disabled")
        
        def on_confirm_choice():
            if not var_salvar_fcd.get() and not var_salvar_rfd.get():
                messagebox.showwarning("Ninguna Selección", "Por favor, seleccione al menos una opción para salvar.", parent=choice_window)
                return
            decision.update({'confirmed': True, 'save_fcd': var_salvar_fcd.get(), 'save_rfd': var_salvar_rfd.get()})
            choice_window.destroy()
        ttk.Button(choice_window, text="Confirmar y Salvar Selección", command=on_confirm_choice).pack(pady=15)
        self.wait_window(choice_window) # Esperar a que se cierre el diálogo
        return decision

    def _salvar_regla_actual(self):
        # (Sin cambios)
        origen_nombre = self.origen_principal_resultados.name if self.origen_principal_resultados != OrigenResultados.NINGUNO else "NINGUNO"
        logging.info(f"Intentando salvar regla. Origen principal: {origen_nombre}, Último término: '{self.ultimo_termino_buscado}'")
        if not self.ultimo_termino_buscado: # Debería estar seteado por _ejecutar_busqueda
            messagebox.showerror("Error", "No hay un término de búsqueda para asociar a la regla.")
            return
        
        puede_salvar_fcd = self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty
        puede_salvar_rfd_o_rdd = self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty

        if not puede_salvar_fcd and not puede_salvar_rfd_o_rdd:
            messagebox.showwarning("Sin Datos Salvables", "No hay datos de la búsqueda actual que se puedan salvar.")
            return

        salvo_algo_en_esta_llamada = False
        if self.origen_principal_resultados.es_via_diccionario:
            decision = self._mostrar_dialogo_seleccion_salvado_via_diccionario()
            if decision['confirmed']:
                timestamp_actual = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S") # Usar el mismo timestamp
                if decision['save_fcd'] and puede_salvar_fcd:
                    self.reglas_guardadas.append({
                        'termino_busqueda': self.ultimo_termino_buscado,
                        'tipo_fuente': "DICCIONARIO_COINCIDENCIAS", # Nombre específico para FCD
                        'datos_relevantes': self.df_candidato_diccionario.copy(),
                        'timestamp': timestamp_actual 
                    })
                    salvo_algo_en_esta_llamada = True
                    logging.info(f"Regla salvada (FCD): '{self.ultimo_termino_buscado}'")
                
                # Solo salvar RFD si viene de VIA_DICCIONARIO_CON_RESULTADOS_DESC
                if decision['save_rfd'] and (self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC and puede_salvar_rfd_o_rdd):
                    self.reglas_guardadas.append({
                        'termino_busqueda': self.ultimo_termino_buscado,
                        'tipo_fuente': OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC.name,
                        'datos_relevantes': self.df_candidato_descripcion.copy(),
                        'timestamp': timestamp_actual 
                    })
                    salvo_algo_en_esta_llamada = True
                    logging.info(f"Regla salvada (RFD): '{self.ultimo_termino_buscado}'")

        elif self.origen_principal_resultados.es_directo_descripcion: # Incluye DIRECTO_DESCRIPCION y DIRECTO_DESCRIPCION_VACIA
            if puede_salvar_rfd_o_rdd:
                self.reglas_guardadas.append({
                    'termino_busqueda': self.ultimo_termino_buscado,
                    'tipo_fuente': self.origen_principal_resultados.name,
                    'datos_relevantes': self.df_candidato_descripcion.copy(),
                    'timestamp': pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                salvo_algo_en_esta_llamada = True
                logging.info(f"Regla salvada (Directa/Vacía): '{self.ultimo_termino_buscado}'")
            else: messagebox.showwarning("Sin Datos", "No hay resultados de la búsqueda directa/vacía para salvar.")
        else: # Origen NINGUNO u otro no manejado
            if self.origen_principal_resultados != OrigenResultados.NINGUNO: # Solo error si no es NINGUNO
                 messagebox.showerror("Error", f"No se puede determinar qué salvar para el origen: {self.origen_principal_resultados.name}.")

        if salvo_algo_en_esta_llamada:
            num_total_reglas = len(self.reglas_guardadas)
            self._actualizar_estado(f"Regla(s) nueva(s) guardada(s). Total: {num_total_reglas}.")
        else: # Si no se confirmó o no había nada que salvar bajo las condiciones
             self._actualizar_estado("Ninguna regla fue salvada en esta operación.")
        
        self.btn_salvar_regla["state"] = "disabled" # Deshabilitar después de intento, _actualizar_botones_estado_general lo re-evaluará
        self._actualizar_botones_estado_general()


    def _exportar_resultados(self):
        # (Sin cambios respecto a la versión anterior del script)
        if not self.reglas_guardadas:
            messagebox.showwarning("Sin Reglas", "No hay reglas guardadas para exportar. Use 'Salvar Regla' primero.")
            return
        
        timestamp_export = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        nombre_sugerido_base = f"exportacion_reglas_{timestamp_export}"
        tipos_archivo = [("Archivo Excel (*.xlsx)", "*.xlsx")]

        ruta_guardar = filedialog.asksaveasfilename(
            title="Exportar Reglas Guardadas Como...",
            initialfile=f"{nombre_sugerido_base}.xlsx",
            defaultextension=".xlsx", filetypes=tipos_archivo
        )
        if not ruta_guardar:
            logging.info("Exportación de reglas cancelada."); self._actualizar_estado("Exportación de reglas cancelada."); return

        self._actualizar_estado("Exportando reglas guardadas..."); num_reglas = len(self.reglas_guardadas)
        logging.info(f"Intentando exportar {num_reglas} regla(s) a: {ruta_guardar}")
        
        try:
            extension = ruta_guardar.split('.')[-1].lower()
            if extension == 'xlsx':
                with pd.ExcelWriter(ruta_guardar, engine='openpyxl') as writer:
                    datos_indice = []
                    for i, regla in enumerate(self.reglas_guardadas):
                        df_datos = regla.get('datos_relevantes')
                        num_filas_datos = len(df_datos) if df_datos is not None else 0
                        datos_indice.append({
                            "ID_Regla_Hoja": f"R{i+1}",
                            "Termino_Busqueda": regla.get('termino_busqueda', 'N/A'),
                            "Fuente_Datos": regla.get('tipo_fuente', 'N/A'),
                            "Filas_Resultado": num_filas_datos,
                            "Timestamp_Guardado": regla.get('timestamp', 'N/A')
                        })
                    df_indice = pd.DataFrame(datos_indice)
                    if not df_indice.empty:
                        df_indice.to_excel(writer, sheet_name="Indice_Reglas", index=False)

                    for i, regla in enumerate(self.reglas_guardadas):
                        df_regla_datos = regla.get('datos_relevantes')
                        if df_regla_datos is not None and isinstance(df_regla_datos, pd.DataFrame) and not df_regla_datos.empty:
                            term_sanitizado_para_hoja = self._sanitizar_nombre_archivo(regla.get('termino_busqueda','S_T'), max_len=15)
                            fuente_abbr = regla.get('tipo_fuente','FNT')[:3] 
                            nombre_hoja_base = f"R{i+1}_{fuente_abbr}_{term_sanitizado_para_hoja}"
                            nombre_hoja = nombre_hoja_base[:31] # Límite Excel para nombres de hoja

                            original_nombre_hoja = nombre_hoja; count = 1
                            # Usar writer.book.sheetnames para verificar nombres de hoja existentes
                            # Esta es la forma correcta con openpyxl a través de ExcelWriter
                            while nombre_hoja in writer.book.sheetnames:
                                nombre_hoja = f"{original_nombre_hoja[:28]}_{count}"; count +=1 # Acortar para sufijo
                                if count > 50 : nombre_hoja = f"ErrorNombre{i}"; break # Evitar bucle infinito
                            try:
                                df_regla_datos.to_excel(writer, sheet_name=nombre_hoja, index=False)
                            except Exception as e_sheet: logging.error(f"Error escribiendo hoja '{nombre_hoja}': {e_sheet}")
                        else: logging.warning(f"Regla R{i+1} ('{regla.get('termino_busqueda')}') sin datos para hoja.")
            else: # Si no es xlsx
                messagebox.showerror("Formato No Soportado", "Solo se soporta exportación de reglas a Excel (.xlsx)."); return

            logging.info(f"Exportación de {num_reglas} regla(s) completada a {ruta_guardar}")
            messagebox.showinfo("Exportación Exitosa", f"{num_reglas} regla(s) exportadas a:\n{ruta_guardar}")
            self._actualizar_estado(f"Reglas exportadas a {os.path.basename(ruta_guardar)}.")

            if messagebox.askyesno("Limpiar Reglas", "Exportación exitosa.\n¿Limpiar reglas guardadas internamente?"):
                self.reglas_guardadas.clear()
                self._actualizar_estado("Reglas guardadas limpiadas.")
                logging.info("Reglas guardadas limpiadas por usuario.")
            self._actualizar_botones_estado_general()

        except Exception as e:
            logging.exception("Error inesperado exportando reglas."); messagebox.showerror("Error Exportar", f"{e}")
            self._actualizar_estado("Error exportando reglas.")

    # >>> INICIO: Nuevos métodos para validar y actualizar botones de operadores <<<
    def _actualizar_estado_botones_operadores(self):
        """Actualiza el estado de los botones de operadores basado en el contenido del último término lógico."""
        texto_completo = self.texto_busqueda_var.get()
        cursor_pos = self.entrada_busqueda.index(tk.INSERT)

        # Obtener el segmento lógico actual basado en la posición del cursor
        # Un segmento es la parte del texto entre operadores lógicos (+, |) o inicio/fin de cadena.
        inicio_segmento = 0
        temp_pos = texto_completo.rfind('+', 0, cursor_pos)
        if temp_pos != -1: inicio_segmento = temp_pos + 1
        temp_pos = texto_completo.rfind('|', 0, cursor_pos)
        if temp_pos != -1: inicio_segmento = max(inicio_segmento, temp_pos + 1)
        
        # Considerar espacios después del delimitador lógico
        segmento_con_espacio_inicial = texto_completo[inicio_segmento:cursor_pos]
        segmento_actual = segmento_con_espacio_inicial.lstrip() # Quitar espacios al inicio del segmento para análisis

        # Determinar tipos de operadores presentes en el segmento actual
        tiene_comparacion_o_rango_visible = any(op in segmento_actual for op in ['>', '<', '>=', '<=', '-'])
        # Una comprobación más robusta sería analizar el término, pero para UI puede ser suficiente
        # analizar visualmente o el término parseado.
        # Ejemplo de análisis más robusto:
        terminos_analizados_segmento = self.motor._analizar_terminos([segmento_actual.strip()])
        es_termino_comparativo_o_rango = False
        if terminos_analizados_segmento:
            tipo_termino = terminos_analizados_segmento[0].get('tipo')
            if tipo_termino in ['gt', 'lt', 'ge', 'le', 'range']:
                es_termino_comparativo_o_rango = True
        
        tiene_negacion_segmento = segmento_actual.startswith('#')

        # Actualizar estados
        # Botón NOT (#): deshabilitado si el segmento ya tiene negación
        self.btn_not['state'] = 'disabled' if tiene_negacion_segmento else 'normal'
        
        # Estado para operadores de comparación (>, <, >=, <=)
        estado_comparacion = 'disabled' if es_termino_comparativo_o_rango or tiene_comparacion_o_rango_visible else 'normal'
        
        # Estado para operador de rango (-)
        estado_rango = 'disabled'
        if not texto_completo.strip():  # Si no hay texto, deshabilitar rango
            estado_rango = 'disabled'
        elif es_termino_comparativo_o_rango or tiene_comparacion_o_rango_visible:  # Si ya hay comparación/rango
            estado_rango = 'disabled'
        else:
            estado_rango = 'normal'
        
        # Aplicar estados a los botones
        for btn in [self.btn_gt, self.btn_lt, self.btn_ge, self.btn_le]:
            btn.config(state=estado_comparacion)
        
        # Aplicar estado específico al botón de rango
        self.btn_range.config(state=estado_rango)
        
        # Operadores lógicos (+, |): deshabilitados si el texto está vacío o el último carácter no es un espacio
        # y es un operador lógico.
        texto_limpio_final = texto_completo.rstrip() # Ignorar espacios al final del texto completo
        estado_logico = 'normal'
        if not texto_completo.strip(): # Si la entrada está vacía (o solo espacios)
            estado_logico = 'disabled'
        elif texto_limpio_final and texto_limpio_final[-1] in ['+', '|']: # Si termina en + o |
            estado_logico = 'disabled'
        
        self.btn_and.config(state=estado_logico)
        self.btn_or.config(state=estado_logico)
    # <<< FIN: Nuevos métodos para validar y actualizar botones de operadores <<<

    def on_closing(self):
        # (Sin cambios)
        logging.info("Cerrando la aplicación...")
        self._guardar_configuracion()
        self.destroy()

    def _insertar_operador_validado(self, operador: str):
        """Inserta un operador en la posición actual del cursor si es válido hacerlo."""
        # Si no hay diccionario cargado, no permitir la inserción
        if self.motor.datos_diccionario is None:
            return

        texto_actual = self.texto_busqueda_var.get()
        cursor_pos = self.entrada_busqueda.index(tk.INSERT)
        
        # Obtener el segmento actual
        inicio_segmento = 0
        temp_pos = texto_actual.rfind('+', 0, cursor_pos)
        if temp_pos != -1: inicio_segmento = temp_pos + 1
        temp_pos = texto_actual.rfind('|', 0, cursor_pos)
        if temp_pos != -1: inicio_segmento = max(inicio_segmento, temp_pos + 1)
        
        segmento_actual = texto_actual[inicio_segmento:cursor_pos].strip()
        
        # Validar si se puede insertar el operador
        puede_insertar = True
        
        if operador in ['>', '<', '>=', '<=']:
            # No permitir si ya hay un operador de comparación en el segmento
            if any(op in segmento_actual for op in ['>', '<', '>=', '<=', '-']):
                puede_insertar = False
        elif operador == '-':
            # No permitir si no hay texto antes o ya hay un operador de comparación/rango
            if not segmento_actual or any(op in segmento_actual for op in ['>', '<', '>=', '<=', '-']):
                puede_insertar = False
        elif operador in ['+', '|']:
            # No permitir si el texto está vacío o termina en operador lógico
            if not texto_actual.strip() or texto_actual.rstrip()[-1] in ['+', '|']:
                puede_insertar = False
        elif operador == '#':
            # No permitir si ya hay negación en el segmento
            if segmento_actual.startswith('#'):
                puede_insertar = False
        
        if puede_insertar:
            # Insertar el operador
            nuevo_texto = texto_actual[:cursor_pos] + operador + texto_actual[cursor_pos:]
            self.texto_busqueda_var.set(nuevo_texto)
            # Mover el cursor después del operador insertado
            self.entrada_busqueda.icursor(cursor_pos + len(operador))
            # Actualizar estado de botones
            self._actualizar_estado_botones_operadores()

    def _deshabilitar_botones_operadores(self):
        """Deshabilita todos los botones operacionales de la interfaz."""
        # Deshabilitar botones de operadores lógicos
        self.btn_and['state'] = 'disabled'
        self.btn_or['state'] = 'disabled'
        self.btn_not['state'] = 'disabled'
        
        # Deshabilitar botones de comparación
        self.btn_gt['state'] = 'disabled'
        self.btn_lt['state'] = 'disabled'
        self.btn_ge['state'] = 'disabled'
        self.btn_le['state'] = 'disabled'
        
        # Deshabilitar botón de rango
        self.btn_range['state'] = 'disabled'

    def _ordenar_resultados(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Ordena los resultados de la búsqueda.
        
        Args:
            df (pd.DataFrame): DataFrame con los resultados a ordenar.
            
        Returns:
            pd.DataFrame: DataFrame ordenado.
        """
        if df is None or df.empty:
            return df
            
        try:
            # Intentar ordenar por la primera columna numérica si existe
            cols_numericas = df.select_dtypes(include=['int64', 'float64']).columns
            if len(cols_numericas) > 0:
                return df.sort_values(by=cols_numericas[0], ascending=False)
            
            # Si no hay columnas numéricas, ordenar por la primera columna
            return df.sort_values(by=df.columns[0])
            
        except Exception as e:
            logging.warning(f"Error al ordenar resultados: {e}")
            return df

# --- Bloque Principal (`if __name__ == "__main__":`) ---
if __name__ == "__main__":
    # (Sin cambios)
    log_file = 'buscador_app.log'
    logging.basicConfig(
        level=logging.DEBUG, 
        format='%(asctime)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8', mode='w'), 
            logging.StreamHandler() 
        ]
    )
    logging.info("=============================================")
    logging.info("=== Iniciando Aplicación Buscador Avanzado ===")
    logging.info(f"Plataforma: {platform.system()} {platform.release()}")
    logging.info(f"Versión de Python: {platform.python_version()}")
    logging.info("=============================================")

    missing_deps = []
    try: import pandas as pd; logging.info(f"Pandas versión: {pd.__version__}")
    except ImportError: missing_deps.append("pandas"); logging.critical("Dependencia faltante: pandas")
    try: import openpyxl; logging.info(f"openpyxl versión: {openpyxl.__version__}")
    except ImportError: missing_deps.append("openpyxl"); logging.critical("Dependencia faltante: openpyxl")

    if missing_deps:
        error_msg = f"Faltan librerías: {', '.join(missing_deps)}.\nInstale con: pip install {' '.join(missing_deps)}"
        logging.critical(error_msg)
        try: # Intentar mostrar mensaje en GUI si Tkinter está disponible
            root = tk.Tk(); root.withdraw(); messagebox.showerror("Dependencias Faltantes", error_msg); root.destroy()
        except tk.TclError: # Fallback a consola si Tkinter no funciona
            print(f"ERROR CRÍTICO: {error_msg}")
        exit(1)
    
    try:
        app = InterfazGrafica()
        app.mainloop()
    except Exception as main_error:
        logging.critical("¡Error fatal no capturado en la aplicación!", exc_info=True)
        try:
            root_err = tk.Tk(); root_err.withdraw() # Ventana invisible para messagebox
            messagebox.showerror("Error Fatal", f"Error crítico:\n{main_error}\nConsulte '{log_file}'.")
            root_err.destroy()
        except Exception as fallback_error: # Si ni siquiera Tkinter funciona para el error
            logging.error(f"No se pudo mostrar el mensaje de error fatal: {fallback_error}")
            print(f"ERROR FATAL: {main_error}. Consulte {log_file}.")
    finally:
        logging.info("=== Finalizando Aplicación Buscador ===")
