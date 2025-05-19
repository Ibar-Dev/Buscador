# -*- coding: utf-8 -*-
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Tuple, Union, Set, Callable, Dict, Any
from enum import Enum, auto # Para saber de dónde vienen los resultados, ¡mola!
import traceback
import platform
import unicodedata
import logging
import json
import os
import string # Para limpiar nombres de archivo y que no explote nada

# --- Configuración del Logging ---
# Esto se monta en el bloque principal, al final del todo.

# --- Enumeraciones ---
class OrigenResultados(Enum):
    """
    Una ayudita para saber cómo hemos llegado a los resultados que mostramos.
    Así podemos cambiar la lógica o los mensajes según si vino del diccionario,
    directo de descripciones, o si no se encontró nada.
    """
    NINGUNO = 0 # Pues eso, no hay resultados o no se ha buscado
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto() # Encontrado en dicc, y luego en descripciones! Éxito!
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto() # Encontrado en dicc, pero las palabras clave eran raras (vacías, etc.)
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto() # Encontrado en dicc, pero luego NADA en descripciones
    DIRECTO_DESCRIPCION = auto() # Búsqueda a piñón en las descripciones
    DIRECTO_DESCRIPCION_VACIA = auto() # Si el usuario no escribe nada y le da a buscar, mostramos todo

    @property
    def es_via_diccionario(self) -> bool:
        """Chequeo rápido: ¿Pasamos por el diccionario para llegar aquí?"""
        return self in {OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
                         OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
                         OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC}

    @property
    def es_directo_descripcion(self) -> bool:
        """Chequeo rápido: ¿Fuimos directos a buscar en las descripciones?"""
        return self in {OrigenResultados.DIRECTO_DESCRIPCION,
                         OrigenResultados.DIRECTO_DESCRIPCION_VACIA}


# --- Clases de Lógica ---

class ManejadorExcel:
    # Esta clase es la curranta de los Excel, los carga y maneja errores comunes.
    @staticmethod
    def cargar_excel(ruta: str) -> Optional[pd.DataFrame]:
        # A ver si podemos cargar este Excel...
        logging.info(f"Intentando cargar archivo Excel: {ruta}")
        try:
            # Si es xlsx, le decimos que use openpyxl, si no, que Pandas se apañe
            engine = 'openpyxl' if ruta.endswith('.xlsx') else None
            df = pd.read_excel(ruta, engine=engine)
            logging.info(f"Archivo '{os.path.basename(ruta)}' cargado ({len(df)} filas). ¡Bien!")
            return df
        except FileNotFoundError:
            logging.error(f"¡Ups! Archivo no encontrado: {ruta}")
            messagebox.showerror("Error de Archivo", f"No se encontró el archivo:\n{ruta}\n\nVerifique que la ruta sea correcta y que el archivo exista.")
            return None
        except Exception as e:
            # Error genérico, ¡a la caza del bug!
            logging.exception(f"Error inesperado al cargar archivo: {ruta}")
            messagebox.showerror("Error al Cargar", 
                                 f"No se pudo cargar el archivo:\n{ruta}\n\nError: {e}\n\n"
                                 "Posibles causas:\n"
                                 "- El archivo está siendo usado por otro programa\n"
                                 "- No tiene instalado 'openpyxl' para archivos .xlsx\n"
                                 "- El archivo está corrupto o en formato no soportado")
            return None

class MotorBusqueda:
    # El cerebro de la bestia: aquí se cuece toda la lógica de búsqueda.
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None # El Excel del diccionario
        self.datos_descripcion: Optional[pd.DataFrame] = None # El Excel de las descripciones
        self.archivo_diccionario_actual: Optional[str] = None # Para recordar qué archivo cargamos
        self.archivo_descripcion_actual: Optional[str] = None # Ídem para descripciones

        # Por defecto, buscamos en las columnas 0 y 3 del diccionario, pero se puede configurar
        self.indices_columnas_busqueda_dic: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else [0, 3]
        logging.info(f"MotorBusqueda inicializado. Índices búsqueda diccionario: {self.indices_columnas_busqueda_dic}")

        # Pre-compilamos los regex para que vaya más rápido, ¡son nuestros amigos!
        self.patron_comparacion_compilado = re.compile(r"^([<>]=?)(\d+([.,]\d+)?).*$") # Para ">100", "<=50", etc.
        self.patron_rango_compilado = re.compile(r"^(\d+([.,]\d+)?)-(\d+([.,]\d+)?)$") # Para "10-20"
        self.patron_negacion_compilado = re.compile(r"^#(.+)$") # Para "#no_quiero_esto"
        self.extractor_magnitud = ExtractorMagnitud() # Un ayudante para sacar cosas como "100W" de "texto 100W texto"

    def cargar_excel_diccionario(self, ruta: str) -> bool:
        # Cargamos el diccionario y, si todo va bien, validamos las columnas.
        self.datos_diccionario = ManejadorExcel.cargar_excel(ruta)
        if self.datos_diccionario is None: self.archivo_diccionario_actual = None; return False # Si falla la carga, abortamos
        self.archivo_diccionario_actual = ruta
        if not self._validar_columnas_diccionario():
            # Si las columnas no cuadran con lo esperado, mejor no seguir.
            logging.warning("Validación columnas diccionario fallida. Invalidando carga.")
            self.datos_diccionario = None; self.archivo_diccionario_actual = None; return False
        return True

    def cargar_excel_descripcion(self, ruta: str) -> bool:
        # Cargamos el excel de descripciones, más sencillito.
        self.datos_descripcion = ManejadorExcel.cargar_excel(ruta)
        if self.datos_descripcion is None: self.archivo_descripcion_actual = None; return False
        self.archivo_descripcion_actual = ruta
        return True

    def _validar_columnas_diccionario(self) -> bool:
        # Comprobamos que el diccionario tenga sentido y las columnas que necesitamos.
        # ¡Importante para que no pete luego al buscar!
        if self.datos_diccionario is None: return False # Si no hay datos, no hay nada que validar
        num_cols = len(self.datos_diccionario.columns)
        if not self.indices_columnas_busqueda_dic:
            logging.error("La lista de índices de columnas de búsqueda está vacía en la configuración.")
            messagebox.showerror("Error de Configuración", 
                                 "No hay índices de columna definidos para la búsqueda en el diccionario.\n\n"
                                 "Por favor, configure los índices de las columnas que desea utilizar para la búsqueda.")
            return False
        
        # El índice más alto que necesitamos no puede ser mayor que el número de columnas que hay.
        max_indice_requerido = max(self.indices_columnas_busqueda_dic) if self.indices_columnas_busqueda_dic else -1

        if num_cols == 0:
            logging.error("Diccionario sin columnas. ¡Esto no puede ser!")
            messagebox.showerror("Error de Diccionario", 
                                 "El archivo del diccionario está vacío o no contiene columnas.\n\n"
                                 "Verifique que el archivo:\n"
                                 "- No esté vacío\n"
                                 "- Tenga al menos una columna de datos\n"
                                 "- Esté en formato Excel válido")
            return False
        elif max_indice_requerido >= num_cols:
            logging.error(f"El diccionario tiene {num_cols} columnas, pero se necesita hasta el índice {max_indice_requerido}.")
            messagebox.showerror("Error de Diccionario", 
                                 f"El diccionario necesita al menos {max_indice_requerido + 1} columnas para los índices configurados ({self.indices_columnas_busqueda_dic}), "
                                 f"pero solo tiene {num_cols}.\n\n"
                                 "Por favor, verifique que:\n"
                                 "- El archivo tiene suficientes columnas\n"
                                 "- Los índices configurados son correctos")
            return False
        return True # ¡Todo en orden!

    def _obtener_nombres_columnas_busqueda(self, df: pd.DataFrame) -> Optional[List[str]]:
        # Traduce los índices de las columnas (ej. 0, 3) a sus nombres reales (ej. "Nombre", "Descripción").
        # Útil para que el usuario vea nombres en lugar de números.
        if df is None: 
            logging.error("Intentando obtener nombres de columnas de un DataFrame que no existe. Mal asunto.")
            return None
        columnas_disponibles = df.columns
        cols_encontradas_nombres = []
        num_cols_df = len(columnas_disponibles)
        indices_validos = []
        for indice in self.indices_columnas_busqueda_dic:
            if isinstance(indice, int) and 0 <= indice < num_cols_df: # El índice tiene que ser un número y estar dentro del rango
                cols_encontradas_nombres.append(columnas_disponibles[indice])
                indices_validos.append(indice)
            else:
                logging.warning(f"Índice {indice} es inválido o se sale de las columnas que tenemos (0 a {num_cols_df-1}). Lo ignoramos.")
        
        if not cols_encontradas_nombres:
            logging.error(f"No hemos podido encontrar ninguna columna válida con los índices: {self.indices_columnas_busqueda_dic}")
            messagebox.showerror("Error de Configuración", 
                                 f"No hay columnas válidas para los índices configurados: {self.indices_columnas_busqueda_dic}\n\n"
                                 "Por favor, verifique que:\n"
                                 "- Los índices configurados son correctos\n"
                                 "- Las columnas existen en el archivo\n"
                                 "- Los índices están dentro del rango válido (0-{num_cols_df-1})")
            return None
        logging.debug(f"Vamos a buscar en estas columnas del diccionario: {cols_encontradas_nombres} (correspondientes a los índices: {indices_validos})")
        return cols_encontradas_nombres

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, cols_nombres: List[str]) -> Set[str]:
        # De las filas que coinciden en el diccionario, sacamos todas las palabras clave únicas.
        # Estas palabras luego se buscarán en el archivo de descripciones.
        terms: Set[str] = set() # Usamos un set para que no haya duplicados, ¡eficiencia!
        if df_coincidencias is None or df_coincidencias.empty or not cols_nombres: return terms # Si no hay datos, no hay términos
        
        cols_validas = [c for c in cols_nombres if c in df_coincidencias.columns] # Nos aseguramos de que las columnas existan
        if not cols_validas: logging.warning(f"Ninguna de las columnas {cols_nombres} están en el DataFrame de coincidencias."); return terms
        
        for col in cols_validas:
            try: 
                # Convertimos a string, pasamos a mayúsculas y sacamos los únicos. NaNs se quitan.
                terms.update(df_coincidencias[col].dropna().astype(str).str.upper().unique())
            except Exception as e: 
                logging.warning(f"Algo ha fallado extrayendo términos de la columna '{col}': {e}")
        
        # Quitamos términos vacíos o que solo sean espacios en blanco.
        terms = {t for t in terms if t and not t.isspace()}
        logging.debug(f"Hemos extraído {len(terms)} términos únicos del diccionario.")
        return terms

    def _buscar_terminos_en_descripciones(self, df_desc: pd.DataFrame, terms: Set[str], require_all: bool = False) -> pd.DataFrame:
        # Ahora, con los términos del diccionario, buscamos en el Excel de descripciones.
        cols_orig = list(df_desc.columns) if df_desc is not None else [] # Guardamos las columnas originales por si algo falla
        if df_desc is None or df_desc.empty or not terms: return pd.DataFrame(columns=cols_orig) # Si no hay dónde buscar o qué buscar...

        logging.info(f"Buscando {len(terms)} términos en {len(df_desc)} descripciones (require_all={require_all}).")
        try:
            # Juntamos todas las columnas de cada fila en un solo string para buscar ahí. Todo a mayúsculas.
            txt_filas = df_desc.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            terms_ok = {t for t in terms if t} # Solo términos que no estén vacíos
            if not terms_ok: logging.warning("No hay términos válidos para buscar en descripciones."); return pd.DataFrame(columns=cols_orig)
            
            # Preparamos los términos para regex, buscando palabras completas (\b).
            terms_esc = [r"\b" + re.escape(t) + r"\b" for t in terms_ok]
            
            if require_all: # Si hay que encontrar TODOS los términos (modo AND)
                mask = txt_filas.apply(lambda txt: all(re.search(p, txt, re.I) for p in terms_esc))
            else: # Si vale con encontrar CUALQUIERA de los términos (modo OR)
                mask = txt_filas.str.contains('|'.join(terms_esc), regex=True, na=False, case=False)
            
            res = df_desc[mask]
            logging.info(f"Búsqueda en descripciones ¡lista! Resultados: {len(res)}.")
            return res
        except Exception as e:
            logging.exception("¡Cataplasma! Error buscando en descripciones.")
            messagebox.showerror("Error en Búsqueda Interna", f"Hubo un problema al buscar en las descripciones:\n{e}")
            return pd.DataFrame(columns=cols_orig) # Devolvemos un DataFrame vacío si falla

    def _parse_numero(self, num_str: str) -> Union[int, float, None]:
        # Pequeño ayudante para convertir un texto a número, aceptando comas como decimales.
        if not isinstance(num_str, str): return None # Si no es texto, no es un número que podamos parsear así
        try:
            return float(num_str.replace(',', '.')) # Cambiamos comas por puntos y a flotante
        except ValueError:
            return None # Si no se puede convertir, pues no era un número válido

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        # Esta es la función que "entiende" lo que el usuario escribe en la búsqueda.
        # Separa si es negación (#), comparación (>, <, >=, <=), rango (10-20) o texto normal.
        palabras_analizadas = []
        patron_comp = self.patron_comparacion_compilado # Usamos los regex pre-compilados
        patron_rango = self.patron_rango_compilado
        patron_neg = self.patron_negacion_compilado

        for term_orig in terminos_brutos:
            term = term_orig.strip() # Quitamos espacios al principio y al final
            negate = False
            item = {'original': term_orig} # Guardamos el término original, por si acaso

            if not term: continue # Si está vacío, al siguiente

            match_neg = patron_neg.match(term) # ¿Es una negación?
            if match_neg:
                negate = True
                term = match_neg.group(1).strip() # Nos quedamos con lo que va después del '#'
            
            if not term: continue # Si después del '#' no hay nada, al siguiente

            item['negate'] = negate # Guardamos si es negado o no

            match_comp = patron_comp.match(term) # ¿Es una comparación numérica?
            match_range = patron_rango.match(term) # ¿O un rango?

            if match_comp:
                op, v_str = match_comp.group(1), match_comp.group(2) # Sacamos el operador ('>') y el número ('100')
                v_num = self._parse_numero(v_str)
                if v_num is not None:
                    op_map = {'>':'gt', '<':'lt', '>=':'ge', '<=':'le'} # Mapeamos a claves más cortas
                    item.update({'tipo': op_map[op], 'valor': v_num})
                else:
                    # Si el número no es válido, lo tratamos como texto normal
                    logging.warning(f"Número inválido '{v_str}' en el término '{term_orig}'. Se tratará como texto.")
                    item.update({'tipo': 'str', 'valor': term}) # El 'term' aquí es el que no tiene el operador
            elif match_range:
                v1_str, v2_str = match_range.group(1), match_range.group(3) # Sacamos los dos números del rango
                v1, v2 = self._parse_numero(v1_str), self._parse_numero(v2_str)
                if v1 is not None and v2 is not None:
                    item.update({'tipo': 'range', 'valor': sorted([v1, v2])}) # Guardamos el rango ordenado
                else:
                    logging.warning(f"Rango inválido en '{term_orig}'. Se tratará como texto.")
                    item.update({'tipo': 'str', 'valor': term})
            else:
                # Si no es nada de lo anterior, es un término de texto simple
                item.update({'tipo': 'str', 'valor': term})
            
            palabras_analizadas.append(item)
        logging.debug(f"Términos analizados y listos para la acción: {palabras_analizadas}")
        return palabras_analizadas

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols_validas: List[str], termino_analizado: Dict[str, Any]) -> pd.Series:
        # ¡Aquí está el meollo del match making! Esta función crea una "máscara" (una serie de Trues/Falses)
        # para un solo término analizado (ej. {'tipo':'gt', 'valor':100, 'negate':False, 'original':'>100W'})
        # indicando qué filas del DataFrame cumplen con ese término.
        
        mask_item_total = pd.Series(False, index=df.index) # Empezamos pensando que ninguna fila cumple
        tipo = termino_analizado['tipo']
        valor_busqueda_original_parseado = termino_analizado['valor'] # Este es el número/texto/rango ya procesado por _analizar_terminos

        # Este regex nos ayudará a encontrar "número [espacio opcional] unidad" en las celdas del Excel.
        # Ej: "100W", "50 V", "2.5A"
        patron_num_unidad_df = re.compile(r"(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑ]+)?")

        for col_n in cols_validas: # Miramos en cada una de las columnas que nos han dicho
            col_series = df[col_n].astype(str) # Convertimos la columna a texto para poder usar regex sin problemas
            mask_col_item = pd.Series(False, index=df.index) # Máscara para esta columna en particular

            if tipo in ['gt', 'lt', 'ge', 'le']: # Si buscamos algo como ">100", "<=50A", etc.
                # 'valor_busqueda_original_parseado' ya tiene el número (ej. 100.0).
                # PERO, si el usuario escribió algo como ">100W", _analizar_terminos solo guardó el 100.
                # Así que tenemos que volver al término original para pillar la unidad de búsqueda.
                
                termino_busqueda_completo_original = termino_analizado.get('original', '') # Ej: ">1000W", "#>50V", "ventilador"
                
                # Vamos a intentar sacar el número y la unidad DEL TÉRMINO DE BÚSQUEDA.
                # Primero, quitamos operadores como '>' o '#' si los hay.
                temp_termino_busqueda = termino_busqueda_completo_original
                if temp_termino_busqueda.startswith('#'):
                    temp_termino_busqueda = temp_termino_busqueda[1:]
                if temp_termino_busqueda.startswith(('>', '<', '>=', '<=')):
                    if temp_termino_busqueda[1] == '=': # Es '>=' o '<='
                        temp_termino_busqueda = temp_termino_busqueda[2:]
                    else: # Es '>' o '<'
                        temp_termino_busqueda = temp_termino_busqueda[1:]
                
                # Ahora temp_termino_busqueda debería ser algo como "1000W" o "50 Vatios" o solo "1000"
                match_termino_busqueda = patron_num_unidad_df.match(temp_termino_busqueda.strip())
                
                num_busqueda_efectivo: Optional[float] = None
                unidad_busqueda_efectiva: Optional[str] = None

                if match_termino_busqueda: # Si encontramos "número [unidad]" en el término de búsqueda
                    try:
                        num_busqueda_efectivo = float(match_termino_busqueda.group(1).replace(',', '.'))
                        if match_termino_busqueda.group(2): # ¿Había unidad?
                            unidad_busqueda_efectiva = self.extractor_magnitud._quitar_diacronicos_y_acentos(match_termino_busqueda.group(2).upper().strip())
                    except ValueError:
                        # Esto no debería pasar si _analizar_terminos hizo bien su trabajo con el número,
                        # pero por si acaso, usamos el número que ya teníamos.
                        logging.warning(f"No se pudo re-parsear el número del término de búsqueda '{temp_termino_busqueda}', usando valor de _analizar_terminos.")
                        num_busqueda_efectivo = self._parse_numero(str(valor_busqueda_original_parseado))
                        # Intentamos pillar la unidad del resto del string, de forma un poco heurística
                        str_valor_original_num = str(valor_busqueda_original_parseado)
                        if temp_termino_busqueda.startswith(str_valor_original_num):
                            posible_unidad_str = temp_termino_busqueda[len(str_valor_original_num):].strip()
                            if posible_unidad_str and not any(c.isdigit() for c in posible_unidad_str):
                                unidad_busqueda_efectiva = self.extractor_magnitud._quitar_diacronicos_y_acentos(posible_unidad_str.upper())
                else: # Si el término de búsqueda era solo un número (ej. ">100"), usamos el número de _analizar_terminos
                    num_busqueda_efectivo = self._parse_numero(str(valor_busqueda_original_parseado))
                
                # Si después de todo esto no tenemos un número para buscar, algo fue muy mal.
                if num_busqueda_efectivo is None:
                    logging.warning(f"Número de búsqueda inválido para el término '{termino_analizado.get('original', 'N/A')}'. Se omite esta parte de la búsqueda.")
                    continue # Pasamos a la siguiente columna o término

                # Ahora vamos fila por fila, celda por celda, en la columna actual del DataFrame
                for idx, val_celda_str in col_series.items():
                    if pd.isna(val_celda_str): continue # Saltamos los NaNs

                    # Buscamos todos los "número [unidad]" que haya en esta celda
                    matches_en_celda = patron_num_unidad_df.finditer(val_celda_str)
                    coincidencia_encontrada_en_celda = False
                    for match_celda in matches_en_celda: # Puede haber varios, ej "Potencia 100W, Voltaje 220V"
                        try:
                            num_celda = float(match_celda.group(1).replace(',', '.'))
                            unidad_celda: Optional[str] = None
                            if match_celda.group(2): # ¿Tiene unidad este valor de la celda?
                                unidad_celda = self.extractor_magnitud._quitar_diacronicos_y_acentos(match_celda.group(2).upper().strip())

                            # Comprobamos la condición numérica (ej. num_celda > num_busqueda_efectivo)
                            condicion_numerica_ok = False
                            if tipo == 'gt': condicion_numerica_ok = num_celda > num_busqueda_efectivo
                            elif tipo == 'lt': condicion_numerica_ok = num_celda < num_busqueda_efectivo
                            elif tipo == 'ge': condicion_numerica_ok = num_celda >= num_busqueda_efectivo
                            elif tipo == 'le': condicion_numerica_ok = num_celda <= num_busqueda_efectivo
                            
                            if condicion_numerica_ok:
                                # ¡El número cumple! Ahora veamos las unidades...
                                if unidad_busqueda_efectiva: 
                                    # Si el término de búsqueda SÍ especificaba una unidad (ej. ">100W")...
                                    # ...entonces la celda también DEBE tener esa misma unidad.
                                    if unidad_celda and unidad_celda == unidad_busqueda_efectiva:
                                        mask_col_item.loc[idx] = True # ¡Coincidencia total!
                                        coincidencia_encontrada_en_celda = True; break 
                                    # Si la búsqueda es ">100W" y la celda dice "120" (sin unidad), NO es match.
                                else: 
                                    # Si el término de búsqueda NO especificaba unidad (ej. ">100")...
                                    # ...entonces CUALQUIER valor numérico en la celda que cumpla la condición es un match,
                                    # independientemente de si la celda tiene unidad o no.
                                    # (Según la lógica anterior, si no hay unidad en búsqueda, SÍ coincidiría si el número cumple.
                                    #  Se ha cambiado el comportamiento para que si la busqueda no tiene unidad, la celda tampoco deba para match estricto)

                                    # Lógica actualizada según tus comentarios: si búsqueda no tiene unidad, celda tampoco debe tenerla.
                                    if not unidad_celda: # Solo hay match si la celda TAMPOCO tiene unidad
                                        mask_col_item.loc[idx] = True
                                        coincidencia_encontrada_en_celda = True; break
                        except ValueError:
                            continue # Si el "número" en la celda no es parseable, a otra cosa.
                    
                    # if coincidencia_encontrada_en_celda: continue # Optimización: si ya encontramos match en esta celda, pasamos a la siguiente fila del df.
                                                                # Esto estaba mal, porque si una fila ya tiene True, no debe ser sobreescrita por otra columna
                                                                # Lo correcto es que mask_item_total |= mask_col_item

            elif tipo == 'str': # Búsqueda de texto normal
                # Buscamos la palabra completa, insensible a mayúsculas/minúsculas
                mask_col_item = col_series.str.contains(
                    r"\b" + re.escape(str(valor_busqueda_original_parseado)) + r"\b", 
                    case=False, na=False, regex=True
                )
            elif tipo == 'range': # Búsqueda por rango numérico (ej. 10-20)
                min_v, max_v = valor_busqueda_original_parseado # Ya está parseado y ordenado
                # Para rangos, asumimos que son solo números, no manejamos unidades explícitamente aquí.
                # Convertimos la columna a números (lo que no sea número se vuelve NaN)
                col_num_range = pd.to_numeric(col_series, errors='coerce')
                mask_col_item = (col_num_range >= min_v) & (col_num_range <= max_v)
                mask_col_item = mask_col_item.fillna(False) # Los NaN no cumplen el rango
            
            # Combinamos la máscara de esta columna con la máscara total para ESTE TÉRMINO.
            # Si una fila cumple en CUALQUIER columna válida, es un True para este término.
            mask_item_total |= mask_col_item.fillna(False) 

        return mask_item_total

    def _aplicar_mascara_diccionario(self, df: pd.DataFrame, cols_nombres: List[str], terms_analizados: List[Dict[str, Any]], op_principal: str) -> pd.Series:
        # Aquí combinamos los resultados de _generar_mascara_para_un_termino para TODOS los términos de la búsqueda.
        # Tiene en cuenta si el operador principal es AND o OR, y las negaciones (#).
        if df is None or df.empty or not cols_nombres or not terms_analizados:
            return pd.Series(False, index=df.index if df is not None else None) # Si no hay chicha, no hay paraiso

        cols_ok = [c for c in cols_nombres if c in df.columns] # Columnas que realmente existen
        if not cols_ok:
            logging.error(f"Ninguna de las columnas de búsqueda ({cols_nombres}) existe en el DataFrame. ¡Houston, tenemos un problema!")
            return pd.Series(False, index=df.index)

        terms_pos = [item for item in terms_analizados if not item.get('negate', False)] # Términos "normales"
        terms_neg = [item for item in terms_analizados if item.get('negate', False)] # Términos con "#"

        op_es_and = op_principal.upper() == 'AND' # El usuario usó '+' o solo un término

        # Empezamos con la máscara para los términos positivos
        if not terms_pos: # Si no hay términos positivos...
            # ...si la operación era AND, ninguna fila cumple (necesitaría algo para ser True).
            # ...si era OR (o por defecto si solo hay negativos), todas las filas "podrían" cumplir inicialmente.
            mask_pos_final = pd.Series(op_es_and, index=df.index) # True si AND y no hay positivos (raro), False si OR. OJO REVISAR ESTA LÓGICA.
                                                                # Si no hay terms_pos, y es AND, debería ser False. Si es OR, también False.
                                                                # La idea es que si no hay positivos, solo cuentan los negativos.
                                                                # Si op_es_and y no hay terms_pos, el resultado final de positivos es "todo cumple", y luego se niega.
                                                                # Si es OR y no hay terms_pos, el resultado de positivos es "nada cumple".
                                                                # Si no hay términos positivos, la máscara de términos positivos es `True` para `AND` (vacuamente verdadero) 
                                                                # y `False` para `OR` (no hay nada que pueda hacerla verdadera).
                                                                # Esto se corregirá con los negativos.
            mask_pos_final = pd.Series(True, index=df.index) # Asumimos que todo cumple y luego los negativos filtran.

        else: # Si hay términos positivos
            # Si es AND, empezamos con todo True y vamos quitando con &. Si es OR, empezamos con todo False y vamos añadiendo con |.
            mask_pos_final = pd.Series(True, index=df.index) if op_es_and else pd.Series(False, index=df.index)
            for item in terms_pos:
                mask_item = self._generar_mascara_para_un_termino(df, cols_ok, item)
                if op_es_and:
                    mask_pos_final &= mask_item # Acumulamos con AND
                else:
                    mask_pos_final |= mask_item # Acumulamos con OR

        # Ahora procesamos los términos negativos
        mask_neg_combinada = pd.Series(False, index=df.index) # Empezamos diciendo que nada está negado
        if terms_neg:
            for item in terms_neg:
                # OJO: El 'item' para _generar_mascara_para_un_termino no debería llevar el 'negate' implícito.
                # El _generar_mascara_para_un_termino busca coincidencias. La negación se aplica DESPUÉS.
                # El 'negate' en 'item' solo lo usamos aquí para separarlos.
                mask_item_neg = self._generar_mascara_para_un_termino(df, cols_ok, item) # Qué filas cumplen con el término que queremos negar
                mask_neg_combinada |= mask_item_neg # Si CUALQUIER término negado coincide, esa fila se marca para ser excluida

        # El resultado final: filas que cumplen los positivos Y NO cumplen NINGUNO de los negativos.
        mask_final = mask_pos_final & (~mask_neg_combinada) 

        logging.debug(f"Máscara Positivos: {mask_pos_final.sum()} filas. Máscara Negativos (antes de invertir): {mask_neg_combinada.sum()} filas. Máscara Final: {mask_final.sum()} filas.")
        return mask_final

    def buscar_en_descripciones_directo(self, term_buscar: str) -> pd.DataFrame:
        # Búsqueda simple y directa en las descripciones, sin pasar por el diccionario.
        # Útil si el término no está en el diccionario o el usuario quiere ir al grano.
        logging.info(f"Búsqueda Directa en Descripciones, buscando: '{term_buscar}'")
        if self.datos_descripcion is None or self.datos_descripcion.empty:
            logging.warning("Intentando búsqueda directa pero no hay descripciones cargadas o están vacías.")
            messagebox.showwarning("Datos Faltantes", 
                                 "No hay datos de descripciones cargados para realizar la búsqueda directa.\n\n"
                                 "Por favor, cargue primero el archivo de descripciones.")
            return pd.DataFrame() # Nada que hacer aquí

        term_limpio = term_buscar.strip().upper() # Limpiamos y a mayúsculas
        if not term_limpio: # Si el usuario no escribió nada (o solo espacios)
            logging.info("Término de búsqueda directa vacío. Devolvemos todas las descripciones.")
            return self.datos_descripcion.copy() # Mostramos todo

        df_desc = self.datos_descripcion.copy() # Trabajamos sobre una copia
        res = pd.DataFrame(columns=df_desc.columns) # Preparamos un DataFrame vacío para los resultados
        try:
            # Juntamos todas las columnas de cada fila en un solo string, a mayúsculas.
            txt_filas = df_desc.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            mask = pd.Series(False, index=df_desc.index) # Empezamos sin coincidencias

            # Lógica simple para AND (+) y OR (| o /) en búsqueda directa
            if '+' in term_limpio: # Modo AND
                palabras = [p.strip() for p in term_limpio.split('+') if p.strip()]
                if not palabras: return res # Si solo eran '+', no hay nada que buscar
                mask = pd.Series(True, index=df_desc.index) # Empezamos con todo True
                for p in palabras:
                    # Buscamos cada palabra como palabra completa (\b) e insensible a mayúsculas
                    mask &= txt_filas.str.contains(r"\b" + re.escape(p) + r"\b", regex=True, na=False)
            elif '|' in term_limpio or '/' in term_limpio: # Modo OR
                sep = '|' if '|' in term_limpio else '/'
                palabras = [p.strip() for p in term_limpio.split(sep) if p.strip()]
                if not palabras: return res
                for p in palabras:
                    mask |= txt_filas.str.contains(r"\b" + re.escape(p) + r"\b", regex=True, na=False)
            else: # Búsqueda de un solo término (o frase)
                mask = txt_filas.str.contains(r"\b" + re.escape(term_limpio) + r"\b", regex=True, na=False)

            res = df_desc[mask] # Aplicamos la máscara para obtener los resultados
            logging.info(f"Búsqueda directa completada. Resultados: {len(res)}.")
        except Exception as e:
            logging.exception("Problemón en la búsqueda directa en descripciones.")
            messagebox.showerror("Error de Búsqueda Directa", 
                                 f"Ocurrió un error durante la búsqueda directa:\n{e}\n\n"
                                 "Por favor, intente nuevamente o contacte al soporte técnico si el problema persiste.")
        return res

# --- Clase ExtractorMagnitud ---
# Esta clase es una pequeña utilidad para encontrar números junto a unidades
# como "100W", "220V", "5A" dentro de un texto más largo. ¡Magia con regex!
class ExtractorMagnitud:
    MAGNITUDES_PREDEFINIDAS: List[str] = [ # Lista de unidades comunes que podríamos querer buscar
        "A","AMP","AMPS","AH","ANTENNA","BASE","BIT","ETH","FE","G","GB",
        "GBE","GE","GIGABIT","GBASE","GBASEWAN","GBIC","GBIT","GBPS","GH",
        "GHZ","HZ","KHZ","KM","KVA","KW","LINEAS","LINES","MHZ","NM","PORT",
        "PORTS","PTOS","PUERTO","PUERTOS","P","V","VA","VAC","VC","VCC",
        "VCD","VDC","W","WATTS","E","FE","GBE","GE","POTS","STM"
    ]
    def __init__(self, magnitudes: Optional[List[str]] = None):
        self.magnitudes = magnitudes if magnitudes is not None else self.MAGNITUDES_PREDEFINIDAS

    @staticmethod
    def _quitar_diacronicos_y_acentos(texto: str) -> str: # Adios acentos y cosas raras
        if not isinstance(texto, str) or not texto: return ""
        try:
            # Normalizamos el texto a su forma NFKD y quitamos los caracteres combinados (acentos, etc.)
            forma_normalizada = unicodedata.normalize('NFKD', texto)
            return ''.join(c for c in forma_normalizada if not unicodedata.combining(c))
        except TypeError: return "" # Por si acaso

    def buscar_cantidad_para_magnitud(self, mag: str, descripcion: str) -> Optional[str]:
        # Dada una magnitud (ej. "W") y un texto, busca si aparece un número antes de esa magnitud.
        if not isinstance(mag, str) or not mag: return None
        if not isinstance(descripcion, str) or not descripcion: return None
        
        mag_upper = mag.upper() # Todo a mayúsculas para comparar sin problemas
        texto_limpio = self._quitar_diacronicos_y_acentos(descripcion.upper()) # Limpiamos el texto también
        if not texto_limpio: return None

        mag_escapada = re.escape(mag_upper) # Escapamos la magnitud por si tiene caracteres especiales de regex
        
        # El patrón busca: número (con decimales opcionales) -> opcionalmente un espacio o 'X' -> la magnitud como palabra completa.
        patron_principal = re.compile(
            r"(\d+([.,]\d+)?)[ X]{0,1}(\b" + mag_escapada + r"\b)(?![a-zA-Z0-9])"
        )
        # (?![a-zA-Z0-9]) es para evitar que pille "10 WATTS" si buscamos "W" y luego hay "WATTS". Queremos "W" solo.
        
        for match in patron_principal.finditer(texto_limpio):
            return match.group(1).strip() # Devolvemos el número encontrado
        return None # No se encontró

# --- Clase Interfaz Gráfica ---
class InterfazGrafica(tk.Tk): # Heredamos de tk.Tk para crear nuestra ventana principal
    CONFIG_FILE = "config_buscador.json" # Nombre del archivo donde guardamos cositas

    def __init__(self):
        super().__init__() # Importante llamar al constructor de la clase padre
        self.title("Buscador Avanzado (con Salvar Regla)") # Título de la ventana, ¡que se vea profesional!
        self.geometry("1250x800") # Tamaño inicial de la ventana

        # Cargamos la configuración (últimos archivos usados, etc.)
        self.config = self._cargar_configuracion()
        indices_cfg = self.config.get("indices_columnas_busqueda_dic", [0, 3]) # Por defecto, columnas 0 y 3
        
        # Creamos nuestros motores de lógica
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg)
        self.extractor_magnitud = ExtractorMagnitud() # Este no lo usamos mucho directamente aquí, más bien el motor

        self.resultados_actuales: Optional[pd.DataFrame] = None # Para guardar los resultados de la última búsqueda
        
        # Esta variable mágica de Tkinter se actualiza sola cuando cambia el texto de búsqueda.
        # Y cuando cambia, llama a nuestra función _on_texto_busqueda_change. ¡Automagia!
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        
        self.ultimo_termino_buscado: Optional[str] = None # Para recordar qué buscó el usuario

        self.reglas_guardadas: List[Dict[str, Any]] = [] # Aquí guardaremos las "reglas" que el usuario salve
        self.df_candidato_diccionario: Optional[pd.DataFrame] = None # Resultados del diccionario
        self.df_candidato_descripcion: Optional[pd.DataFrame] = None # Resultados de las descripciones
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO # De dónde vienen los resultados principales

        self.color_fila_par = "white"
        self.color_fila_impar = "#f0f0f0" # Para que las tablas se vean bonitas a rayas

        # Montamos todos los cacharritos de la interfaz
        self._configurar_estilo_ttk() # Para que se vea un poco más moderno
        self._crear_widgets() # Creamos botones, etiquetas, tablas...
        self._configurar_grid() # Los colocamos en su sitio
        self._configurar_eventos() # Configuramos qué pasa cuando se pulsa Enter, etc.
        self._configurar_tags_treeview() # Colores para las filas de las tablas
        self._configurar_orden_tabla(self.tabla_resultados) # Para poder ordenar las tablas por columna
        self._actualizar_estado("Listo. Cargue Diccionario y Descripciones.") # Mensaje inicial en la barra de estado
        
        # Al principio, los botones de operadores (+, |, #, >, <, etc.) están desactivados
        self._deshabilitar_botones_operadores()
        
        # Y actualizamos el estado general de todos los botones (Buscar, Exportar, etc.)
        self._actualizar_botones_estado_general()
        
        logging.info("Interfaz Gráfica inicializada y lista para la acción.")

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str):
        """Cada vez que el usuario escribe algo en la caja de búsqueda, esta función se entera."""
        # No usamos los argumentos var_name, index, mode, pero Tkinter los pasa.
        # Lo importante es que actualizamos el estado de los botones de operadores (+, #, >, etc.)
        # para que se activen o desactiven según lo que haya escrito el usuario. ¡Inteligente!
        self._actualizar_estado_botones_operadores()
        # NOTA: La validación de no poner "++" o "||" seguidos ya la hacemos
        # al insertar los operadores con los botones, así que aquí no hace falta
        # preocuparse por eso para el color de fondo (que no está implementado de todas formas).

    def _cargar_configuracion(self) -> Dict:
        # Cargamos la configuración desde un archivo JSON. Si no existe, creamos una por defecto.
        config = {}
        if os.path.exists(self.CONFIG_FILE): # ¿Existe el archivo?
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f) # ¡Cargado!
                logging.info(f"Configuración cargada desde: {self.CONFIG_FILE}")
            except Exception as e:
                logging.error(f"Error al cargar el archivo de configuración '{self.CONFIG_FILE}': {e}")
                messagebox.showwarning("Error Configuración", f"No se pudo cargar la configuración:\n{e}")
        else:
            # Si no existe, no pasa nada, se creará uno nuevo al cerrar la app.
            logging.info("Archivo de configuración no encontrado. Se creará uno nuevo al cerrar.")
        
        # Nos aseguramos de que las claves importantes existan en la config, aunque sea con None o valores por defecto.
        config.setdefault("last_dic_path", None)
        config.setdefault("last_desc_path", None)
        config.setdefault("indices_columnas_busqueda_dic", [0, 3]) # Columnas 0 y 3 por defecto para el diccionario
        return config

    def _guardar_configuracion(self):
        # Guardamos la configuración actual en el archivo JSON.
        # Así la próxima vez que se abra la app, recordará los últimos archivos y configuraciones.
        self.config["last_dic_path"] = self.motor.archivo_diccionario_actual
        self.config["last_desc_path"] = self.motor.archivo_descripcion_actual
        self.config["indices_columnas_busqueda_dic"] = self.motor.indices_columnas_busqueda_dic
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4) # indent=4 para que el JSON se vea bonito y legible
            logging.info(f"Configuración guardada en: {self.CONFIG_FILE}")
        except Exception as e:
            logging.error(f"¡Oh, no! Error al guardar la configuración en '{self.CONFIG_FILE}': {e}")
            messagebox.showerror("Error Configuración", f"No se pudo guardar la configuración:\n{e}")

    def _configurar_estilo_ttk(self):
        # Intentamos que la app se vea un poco más nativa o al menos no tan "Tk básico".
        # Buscamos un tema TTK que se adapte al sistema operativo.
        style = ttk.Style(self)
        themes = style.theme_names() # ¿Qué temas hay disponibles?
        os_name = platform.system() # ¿Estamos en Windows, Mac, Linux?
        
        # Preferencias de temas por S.O.
        prefs = {"Windows":["vista","xpnative","clam"],"Darwin":["aqua","clam"],"Linux":["clam","alt","default"]}
        theme_to_use = next((t for t in prefs.get(os_name, ["clam","default"]) if t in themes), None)
        
        if not theme_to_use: # Si no encontramos uno preferido, usamos el que esté activo o el primero de la lista
            theme_to_use = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))
        
        if theme_to_use:
            logging.info(f"Intentando aplicar el tema TTK: {theme_to_use}")
            try:
                style.theme_use(theme_to_use) # ¡A ver si funciona!
            except tk.TclError as e:
                logging.warning(f"No se pudo aplicar el tema '{theme_to_use}': {e}. Usando el tema por defecto.")
        else:
            logging.warning("No se encontró ningún tema TTK disponible. Se verá... clásico.")

    def _crear_widgets(self):
        # Aquí es donde nacen todos los botones, etiquetas, y demás elementos de la interfaz.
        # Es como montar un puzzle visual.

        # --- Marco de Controles (arriba) ---
        self.marco_controles = ttk.LabelFrame(self, text="Controles") # Un marquito para agrupar cosas

        # Botones para cargar archivos
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.lbl_dic_cargado = ttk.Label(self.marco_controles, text="Dic: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion)
        self.lbl_desc_cargado = ttk.Label(self.marco_controles, text="Desc: Ninguno", width=20, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)

        # Marco para los botoncitos de operadores (+, |, #, >, <, etc.)
        self.frame_ops = ttk.Frame(self.marco_controles)
        # Cada botón llama a _insertar_operador_validado para añadir el operador al texto de búsqueda
        self.btn_and = ttk.Button(self.frame_ops, text="+", width=3, command=lambda: self._insertar_operador_validado("+"))
        self.btn_or = ttk.Button(self.frame_ops, text="|", width=3, command=lambda: self._insertar_operador_validado("|"))
        self.btn_not = ttk.Button(self.frame_ops, text="#", width=3, command=lambda: self._insertar_operador_validado("#"))
        self.btn_gt = ttk.Button(self.frame_ops, text=">", width=3, command=lambda: self._insertar_operador_validado(">"))
        self.btn_lt = ttk.Button(self.frame_ops, text="<", width=3, command=lambda: self._insertar_operador_validado("<"))
        self.btn_ge = ttk.Button(self.frame_ops, text="≥", width=3, command=lambda: self._insertar_operador_validado(">=")) # Símbolo chulo de mayor o igual
        self.btn_le = ttk.Button(self.frame_ops, text="≤", width=3, command=lambda: self._insertar_operador_validado("<=")) # Ídem para menor o igual
        self.btn_range = ttk.Button(self.frame_ops, text="-", width=3, command=lambda: self._insertar_operador_validado("-"))
        
        # Colocamos los botoncitos de operadores en fila
        self.btn_and.grid(row=0, column=0, padx=1); self.btn_or.grid(row=0, column=1, padx=1)
        self.btn_not.grid(row=0, column=2, padx=1); self.btn_gt.grid(row=0, column=3, padx=1)
        self.btn_lt.grid(row=0, column=4, padx=1); self.btn_ge.grid(row=0, column=5, padx=1)
        self.btn_le.grid(row=0, column=6, padx=1); self.btn_range.grid(row=0, column=7, padx=1)

        # La caja de texto donde el usuario escribe lo que quiere buscar
        # Usamos textvariable para que se conecte con nuestra self.texto_busqueda_var
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50, textvariable=self.texto_busqueda_var)
        
        # Botones principales de acción
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda)
        self.btn_salvar_regla = ttk.Button(self.marco_controles, text="Salvar Regla", command=self._salvar_regla_actual)
        self.btn_ayuda = ttk.Button(self.marco_controles, text="?", command=self._mostrar_ayuda, width=3) # El botón del pánico
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar", command=self._exportar_resultados)

        # --- Área de Tablas (abajo) ---
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario:")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados / Descripciones:")

        # Tabla para la vista previa del diccionario
        self.frame_tabla_diccionario = ttk.Frame(self) # Un marco para la tabla y sus scrollbars
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings", height=8) # El Treeview es el widget tabla de ttk
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)

        # Tabla para los resultados de la búsqueda (o descripciones cargadas)
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings")
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)

        # --- Barra de Estado (abajo del todo) ---
        self.barra_estado = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W, borderwidth=1) # Para mostrar mensajes al usuario
        self._actualizar_etiquetas_archivos() # Mostramos los nombres de los archivos cargados (o "Ninguno")


    def _configurar_grid(self):
        # Aquí definimos cómo se expanden y contraen las filas y columnas de la ventana principal
        # y dónde va cada widget grande. Es como jugar al Tetris con los componentes.
        self.grid_rowconfigure(2, weight=1) # La fila de la tabla del diccionario se puede expandir
        self.grid_rowconfigure(4, weight=3) # La fila de la tabla de resultados se expande más
        self.grid_columnconfigure(0, weight=1) # La única columna principal se expande

        # Colocamos el marco de controles arriba
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5)) # new = North, East, West (se expande horizontalmente)

        # Dentro del marco de controles, también usamos grid
        self.marco_controles.grid_columnconfigure(1, weight=1) # La columna de la etiqueta del diccionario se expande
        self.marco_controles.grid_columnconfigure(3, weight=1) # Ídem para descripciones
        
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w") # w = West (alineado a la izquierda)
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew") # ew = East, West (se expande horizontalmente)
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew")

        # El marco de los operadores
        self.frame_ops.grid(row=1, column=0, columnspan=6, padx=5, pady=(5,0), sticky="w") # columnspan=6 para que ocupe todo el ancho de los botones de abajo

        # La entrada de búsqueda y sus botones
        self.entrada_busqueda.grid(row=2, column=0, columnspan=2, padx=5, pady=(0,5), sticky="ew")
        self.btn_buscar.grid(row=2, column=2, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_salvar_regla.grid(row=2, column=3, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_ayuda.grid(row=2, column=4, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_exportar.grid(row=2, column=5, padx=(10, 5), pady=(0,5), sticky="e") # e = East (alineado a la derecha)

        # Etiquetas de las tablas
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0)) # sw = South, West (abajo a la izquierda)
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))

        # Marcos de las tablas (con sus scrollbars)
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10)) # nsew = North, South, East, West (se expande en todas direcciones)
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1); self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1) # Para que la tabla dentro del marco se expanda
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew")
        self.scrolly_diccionario.grid(row=0, column=1, sticky="ns") # ns = North, South (se expande verticalmente)
        self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")

        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1); self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.scrolly_resultados.grid(row=0, column=1, sticky="ns")
        self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # La barra de estado abajo del todo
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0)) # sew = South, East, West


    def _configurar_eventos(self):
        # Aquí conectamos acciones del usuario (como pulsar Enter) con funciones nuestras.
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda()) # Pulsar Enter en la búsqueda lanza la búsqueda
        self.protocol("WM_DELETE_WINDOW", self.on_closing) # Qué hacer cuando el usuario cierra la ventana (la X)

    def _actualizar_estado(self, mensaje: str):
        # Muestra un mensaje en la barra de estado y lo guarda en el log. Práctico.
        self.barra_estado.config(text=mensaje)
        logging.info(f"Estado UI: {mensaje}")
        self.update_idletasks() # Fuerza a Tkinter a redibujar, para que el mensaje aparezca ya

    def _mostrar_ayuda(self):
        # Un tochaco de ayuda para el usuario. ¡Que no se diga que no explicamos las cosas!
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
        # Ponemos colores alternos a las filas de las tablas, para que sea más fácil leerlas.
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par) # Filas pares
            tabla.tag_configure('impar', background=self.color_fila_impar) # Filas impares

    def _configurar_orden_tabla(self, tabla: ttk.Treeview):
        # Hacemos que se pueda hacer clic en las cabeceras de las columnas para ordenar.
        cols = tabla["columns"] # Obtenemos las columnas actuales de la tabla
        if cols: # Si hay columnas...
            for col in cols:
                # Al hacer clic, llamamos a _ordenar_columna. La lambda es para pasar el nombre de la columna.
                tabla.heading(col, text=col, anchor=tk.W,
                              command=lambda c=col: self._ordenar_columna(tabla, c, False))

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool):
        # Esta función se encarga de la magia de ordenar la tabla de resultados.
        df_para_ordenar = self.resultados_actuales # Usamos los resultados actuales que tenemos en memoria
        
        # Solo ordenamos la tabla de resultados y si tiene datos.
        if tabla != self.tabla_resultados or df_para_ordenar is None or df_para_ordenar.empty:
            logging.debug("Intentando ordenar una tabla que no es la de resultados, o está vacía. No hacemos nada.")
            return

        logging.info(f"Ordenando resultados por columna '{col}', descendente={reverse}")
        try:
            # Intentamos convertir la columna a números para ordenar numéricamente si es posible.
            # Si no, Pandas ordenará como texto. 'na_position' manda los NaN (vacíos) al final.
            df_ordenado = df_para_ordenar.sort_values(
                by=col, 
                ascending=not reverse, # Si reverse es True (queremos descendente), ascending debe ser False.
                na_position='last',
                key=lambda x: pd.to_numeric(x, errors='coerce') if x.dtype == object or pd.api.types.is_string_dtype(x) else x
                # El 'key' es un truquito: si la columna es texto (object) o string, intenta convertirla a números para ordenar.
                # Si ya es numérica, la deja tal cual. 'coerce' hace que los errores de conversión se vuelvan NaN.
            )
            self.resultados_actuales = df_ordenado # Actualizamos nuestros datos con la versión ordenada
            self._actualizar_tabla(tabla, self.resultados_actuales) # Volvemos a pintar la tabla, ya ordenada
            # Actualizamos el comando de la cabecera para que la próxima vez ordene al revés (asc <-> desc)
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, not reverse))
            self._actualizar_estado(f"Resultados ordenados por '{col}' ({'Ascendente' if not reverse else 'Descendente'}).")
        except Exception as e:
            logging.exception(f"¡Zas! Error al intentar ordenar por la columna '{col}'")
            messagebox.showerror("Error al Ordenar", f"No se pudo ordenar por '{col}':\n{e}")
            # Si falla, reseteamos el comando para que no entre en un bucle de errores
            tabla.heading(col, command=lambda c=col: self._ordenar_columna(tabla, c, False))


    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None, columnas_a_mostrar: Optional[List[str]] = None):
        # Esta función es la encargada de "pintar" los datos de un DataFrame en una tabla de la GUI.
        # Es bastante genérica y la usamos para la tabla del diccionario y la de resultados.
        es_tabla_dic = tabla == self.tabla_diccionario
        logging.debug(f"Actualizando la tabla de {'Diccionario' if es_tabla_dic else 'Resultados'}.")
        
        # 1. Limpiamos la tabla de lo que tuviera antes
        try:
            for i in tabla.get_children(): tabla.delete(i) # Borramos todas las filas viejas
        except tk.TclError as e: # A veces Tkinter se queja si la tabla ya no existe o algo raro
            logging.warning(f"Pequeño error de Tcl al limpiar la tabla (probablemente no sea grave): {e}"); pass
        tabla["columns"] = () # Quitamos las columnas viejas también

        # 2. Si no hay datos o el DataFrame está vacío, no hay nada que pintar.
        if datos is None or datos.empty:
            logging.debug("No hay datos para mostrar en la tabla. Dejándola vacía.")
            return

        # 3. Definimos qué columnas vamos a mostrar
        cols_originales = list(datos.columns)
        if columnas_a_mostrar: # Si nos han dicho explícitamente qué columnas...
            cols_ok = [c for c in columnas_a_mostrar if c in cols_originales] # ...nos quedamos solo con las que existan de verdad
            if not cols_ok: # Si ninguna de las pedidas existe, mostramos todas las originales
                logging.warning(f"Las columnas especificadas {columnas_a_mostrar} no se encontraron. Mostrando todas las columnas del DataFrame.")
                cols_ok = cols_originales
        else: # Si no nos dicen nada, mostramos todas las columnas del DataFrame
            cols_ok = cols_originales
        
        if not cols_ok: # Si después de todo esto no tenemos columnas para mostrar...
            logging.warning("El DataFrame no tiene columnas para mostrar. Algo raro pasa.")
            return

        df_mostrar = datos[cols_ok] # Nos quedamos solo con las columnas que vamos a pintar
        tabla["columns"] = cols_ok # Configuramos las columnas en el widget Treeview

        # 4. Configuramos cada columna (cabecera, ancho)
        for col in cols_ok:
            tabla.heading(col, text=str(col), anchor=tk.W) # Ponemos el nombre de la columna en la cabecera
            try:
                # Intentamos calcular un ancho apañado para la columna basado en el contenido
                muestra = min(len(df_mostrar), 100) # Miramos unas pocas filas para no tardar mucho
                if col in df_mostrar.columns:
                    # Longitud máxima del texto en la muestra de esta columna (ignorando NaN)
                    sub_df = df_mostrar.iloc[:muestra][col].dropna().astype(str)
                    ancho_contenido = sub_df.str.len().max() if not sub_df.empty else 0
                else: # Esto no debería pasar si cols_ok está bien
                    ancho_contenido = 0 
                    logging.warning(f"La columna '{col}' no se encontró al calcular su ancho. Raro, raro.")
                
                ancho_cabecera = len(str(col)) # Ancho del título de la columna
                # El ancho final es un max/min para que no sea ni muy estrecha ni exageradamente ancha
                ancho = max(70, min(int(max(ancho_cabecera * 9, ancho_contenido * 7) + 20), 400)) 
                tabla.column(col, anchor=tk.W, width=ancho, minwidth=70)
            except Exception as e:
                logging.warning(f"No se pudo calcular bien el ancho para la columna '{col}': {e}. Usando ancho por defecto.")
                tabla.column(col, anchor=tk.W, width=100, minwidth=50) # Un ancho por si todo falla

        # 5. Metemos las filas en la tabla
        # Si nos han puesto un límite de filas, lo aplicamos
        df_final = df_mostrar.head(limite_filas) if limite_filas is not None and len(df_mostrar) > limite_filas else df_mostrar
        logging.debug(f"Mostrando {len(df_final)} filas en la tabla.")
        
        for i, (_, row) in enumerate(df_final.iterrows()): # Iteramos por cada fila del DataFrame
            # Convertimos todos los valores de la fila a string (si son NaN, ponemos "")
            vals = [str(v) if pd.notna(v) else "" for v in row.values]
            tag = 'par' if i % 2 == 0 else 'impar' # Para los colores alternos
            try:
                tabla.insert("", "end", values=vals, tags=(tag,)) # ¡A pintar la fila!
            except tk.TclError as e: # A veces, caracteres raros dan problemas en Tkinter
                logging.warning(f"Error de Tcl al insertar la fila {i}: {e}. Intentando con caracteres ASCII.")
                try: # Intentamos convertir los valores a ASCII puro, a ver si así cuela
                    vals_ascii = [v.encode('ascii', 'ignore').decode('ascii') for v in vals]
                    tabla.insert("", "end", values=vals_ascii, tags=(tag,))
                except Exception as e_inner:
                    logging.error(f"Ni siquiera el fallback a ASCII funcionó para la fila {i}: {e_inner}")
        
        # Si estamos actualizando la tabla de resultados, nos aseguramos de que se pueda ordenar
        if tabla == self.tabla_resultados:
            self._configurar_orden_tabla(tabla)


    def _actualizar_etiquetas_archivos(self):
        # Ponemos los nombres de los archivos cargados en las etiquetas, o "Ninguno".
        # Si el nombre es muy largo, lo acortamos para que quepa.
        dic_name = os.path.basename(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else "Ninguno"
        desc_name = os.path.basename(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else "Ninguno"
        max_len_label = 25 # Máximo de caracteres para el nombre del archivo en la etiqueta
        
        dic_display = f"Dic: {dic_name}"
        if len(dic_name) > max_len_label: # Si es muy largo...
            dic_display = f"Dic: ...{dic_name[-(max_len_label-4):]}" # ...mostramos "...ultimos_caracteres"
        
        desc_display = f"Desc: {desc_name}"
        if len(desc_name) > max_len_label:
            desc_display = f"Desc: ...{desc_name[-(max_len_label-4):]}"
            
        self.lbl_dic_cargado.config(text=dic_display)
        self.lbl_desc_cargado.config(text=desc_display)

    def _actualizar_botones_estado_general(self):
        """Aquí se decide qué botones están activos y cuáles no, según lo que haya pasado en la app."""
        dic_cargado = self.motor.datos_diccionario is not None
        desc_cargado = self.motor.datos_descripcion is not None

        # Botones de operadores (+, #, >, <, etc.): solo activos si hay diccionario cargado
        estado_base_operadores = 'normal' if dic_cargado else 'disabled'
        
        self.btn_and['state'] = estado_base_operadores
        self.btn_or['state'] = estado_base_operadores
        self.btn_not['state'] = estado_base_operadores
        self.btn_gt['state'] = estado_base_operadores
        self.btn_lt['state'] = estado_base_operadores
        self.btn_ge['state'] = estado_base_operadores
        self.btn_le['state'] = estado_base_operadores
        self.btn_range['state'] = estado_base_operadores

        # Si el diccionario está cargado, aplicamos la lógica más fina de activar/desactivar
        # los operadores según el texto de búsqueda actual.
        if dic_cargado:
            self._actualizar_estado_botones_operadores()
        else:
            # Si no hay diccionario, nos aseguramos de que estén todos desactivados.
            self._deshabilitar_botones_operadores()

        # Otros botones:
        # Cargar descripciones: solo si ya hay un diccionario
        self.btn_cargar_descripciones['state'] = 'normal' if dic_cargado else 'disabled'
        # Buscar: solo si ambos archivos están cargados
        self.btn_buscar['state'] = 'normal' if dic_cargado and desc_cargado else 'disabled'
        
        # Botón "Salvar Regla": ¡este es un poco más lioso!
        puede_salvar_algo_del_diccionario = self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty
        puede_salvar_algo_de_descripciones = self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty
        estado_salvar = 'disabled' # Por defecto, no se puede salvar

        if self.origen_principal_resultados != OrigenResultados.NINGUNO: # Si hubo alguna búsqueda...
            if self.origen_principal_resultados.es_directo_descripcion: # Si fue búsqueda directa (o vacía)...
                if puede_salvar_algo_de_descripciones: estado_salvar = 'normal' # ...podemos salvar si hay resultados de descripciones
            elif self.origen_principal_resultados.es_via_diccionario: # Si pasamos por el diccionario...
                # Podemos salvar si hay algo del diccionario, O si hay algo de descripciones Y estas vienen del diccionario
                if puede_salvar_algo_del_diccionario or \
                   (self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC and puede_salvar_algo_de_descripciones):
                    estado_salvar = 'normal'
        self.btn_salvar_regla['state'] = estado_salvar
        
        # Exportar: solo si hay reglas guardadas
        self.btn_exportar['state'] = 'normal' if self.reglas_guardadas else 'disabled'

    def _cargar_diccionario(self):
        """El usuario quiere cargar el archivo 'diccionario'. ¡Manos a la obra!"""
        try:
            # El botón de cargar diccionario siempre debe estar activo, ¡por si quiere cambiarlo!
            self.btn_cargar_diccionario['state'] = 'normal' 
            
            # Recordamos la última carpeta que usó para que sea más cómodo
            last_dir = os.path.dirname(self.config.get("last_dic_path", "") or "") or None
            ruta = filedialog.askopenfilename(
                title="Seleccionar Archivo Diccionario", 
                filetypes=[("Archivos Excel", "*.xlsx *.xls")], 
                initialdir=last_dir
            )
            if not ruta: # Si el usuario cancela, no hacemos nada
                logging.info("Carga de diccionario cancelada por el usuario.")
                return

            nombre_archivo = os.path.basename(ruta)
            self._actualizar_estado(f"Cargando diccionario: {nombre_archivo}...")
            
            # Limpiamos cosas de búsquedas anteriores
            self._actualizar_tabla(self.tabla_diccionario, None) 
            self.resultados_actuales = None
            self.df_candidato_diccionario = None
            self.df_candidato_descripcion = None
            self.origen_principal_resultados = OrigenResultados.NINGUNO
            self._actualizar_tabla(self.tabla_resultados, None) 

            if self.motor.cargar_excel_diccionario(ruta): # Intentamos cargar con nuestro motor
                self._guardar_configuracion() # Si va bien, guardamos la ruta para la próxima vez
                df_dic = self.motor.datos_diccionario
                if df_dic is not None:
                    num_filas = len(df_dic)
                    # Mostramos en qué columnas del diccionario vamos a buscar
                    cols_busqueda_nombres = self.motor._obtener_nombres_columnas_busqueda(df_dic)
                    indices_str = ', '.join(map(str, self.motor.indices_columnas_busqueda_dic))
                    lbl_text = f"Vista Previa Diccionario (Índices: {indices_str})"
                    if cols_busqueda_nombres: 
                        lbl_text = f"Vista Previa Dic ({', '.join(cols_busqueda_nombres)} - Índices: {indices_str})"
                    self.lbl_tabla_diccionario.config(text=lbl_text)
                    
                    # Mostramos una vista previa del diccionario (primeras 100 filas)
                    self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols_busqueda_nombres)
                    self.title(f"Buscador - Dic: {nombre_archivo}") # Ponemos el nombre en el título de la ventana
                    self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado. ¡A jugar!")
            else:
                # Si hubo un error al cargar (ya se habrá mostrado un messagebox desde el motor)
                self._actualizar_estado("Error al cargar el diccionario.")
                self.title("Buscador Avanzado (con Salvar Regla)") # Título por defecto

            self._actualizar_etiquetas_archivos() # Actualizamos las etiquetas con los nombres de archivo
            self._actualizar_botones_estado_general() # Y el estado de los botones
        except Exception as e: # Por si algo muy gordo falla aquí
            logging.error(f"Error catastrófico al cargar el diccionario: {e}")
            messagebox.showerror("Error al Cargar Diccionario", f"No se pudo cargar el archivo del diccionario:\n{e}")

    def _cargar_excel_descripcion(self):
        # Similar a cargar diccionario, pero para el archivo de descripciones.
        last_dir = os.path.dirname(self.config.get("last_desc_path", "") or "") or None
        ruta = filedialog.askopenfilename(
            title="Seleccionar Archivo de Descripciones", 
            filetypes=[("Archivos Excel", "*.xlsx *.xls")], 
            initialdir=last_dir
        )
        if not ruta: logging.info("Carga de descripciones cancelada."); return

        nombre_archivo = os.path.basename(ruta)
        self._actualizar_estado(f"Cargando descripciones: {nombre_archivo}...")
        
        # Reseteamos resultados de búsquedas anteriores
        self.resultados_actuales = None
        self.df_candidato_diccionario = None
        self.df_candidato_descripcion = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO
        
        if self.motor.cargar_excel_descripcion(ruta):
            self._guardar_configuracion()
            df_desc = self.motor.datos_descripcion
            if df_desc is not None:
                num_filas = len(df_desc)
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando datos...")
                # Al cargar las descripciones, las mostramos directamente en la tabla de resultados.
                self._actualizar_tabla(self.tabla_resultados, df_desc) 

                if self.motor.archivo_diccionario_actual: # Actualizamos el título de la ventana
                    dic_n = os.path.basename(self.motor.archivo_diccionario_actual)
                    self.title(f"Buscador - Dic: {dic_n} | Desc: {nombre_archivo}")
                self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas.")
        else:
            self._actualizar_estado("Error al cargar las descripciones.")
        
        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _buscar_y_enfocar_en_preview(self):
        # Cuando se hace una búsqueda que pasa por el diccionario, intentamos
        # encontrar el término en la tabla de vista previa del diccionario y hacer scroll hasta él.
        # ¡Un detallito de UX!
        termino_buscar = self.texto_busqueda_var.get().strip()
        if not termino_buscar: return # Si no hay término, no hay nada que enfocar

        items_preview = self.tabla_diccionario.get_children('') # Todos los items de la tabla
        if not items_preview: return # Si la tabla está vacía

        termino_upper = termino_buscar.upper() # Para buscar sin importar mayús/minús
        logging.info(f"Intentando enfocar '{termino_buscar}' en la vista previa del diccionario...")

        found_item_id = None
        for item_id in items_preview: # Recorremos cada fila de la tabla
            try:
                valores_fila = self.tabla_diccionario.item(item_id, 'values') # Sacamos los valores de la fila
                # Si el término está en alguno de los valores de la fila...
                if any(termino_upper in str(val).upper() for val in valores_fila):
                    found_item_id = item_id; break # ¡Lo encontramos! Salimos del bucle.
            except Exception as e: # Por si acaso algo falla al leer un item
                logging.warning(f"Error procesando item {item_id} en la vista previa del diccionario: {e}"); continue
        
        if found_item_id:
            logging.info(f"Término '{termino_buscar}' encontrado en vista previa (item ID: {found_item_id}). Haciendo scroll...")
            try:
                # Seleccionamos y hacemos visible el item encontrado
                current_selection = self.tabla_diccionario.selection()
                if current_selection: self.tabla_diccionario.selection_remove(current_selection) # Quitamos selección anterior
                self.tabla_diccionario.selection_set(found_item_id)
                self.tabla_diccionario.see(found_item_id) # ¡Magia de Tkinter para hacer scroll!
            except Exception as e:
                logging.error(f"Error al intentar enfocar el item {found_item_id} en la vista previa: {e}")
        else:
            logging.info(f"El término '{termino_buscar}' no se encontró en la vista previa del diccionario.")


    def _parsear_termino_busqueda_inicial(self, termino_raw: str) -> Tuple[str, List[str]]:
        # Divide el término de búsqueda inicial en partes según los operadores principales (+, |, /)
        # y devuelve el operador principal ('AND' o 'OR') y la lista de sub-términos.
        termino_limpio = termino_raw.strip()
        op_principal = 'OR' # Por defecto, si no hay '+' es OR (o un solo término)
        terminos_brutos = []

        if not termino_limpio: return op_principal, [] # Búsqueda vacía

        if '+' in termino_limpio: # Si hay un '+', es un AND
            op_principal = 'AND'
            terminos_brutos = [p.strip() for p in termino_limpio.split('+') if p.strip()]
        elif '|' in termino_limpio: # Si hay un '|', es un OR
            op_principal = 'OR'
            terminos_brutos = [p.strip() for p in termino_limpio.split('|') if p.strip()]
        elif '/' in termino_limpio: # '/' también es OR
            op_principal = 'OR'
            terminos_brutos = [p.strip() for p in termino_limpio.split('/') if p.strip()]
        else: # Si no hay operadores principales, es un solo término (o una frase)
            terminos_brutos = [termino_limpio]
        
        # Si después de trocear no queda nada (ej. el usuario escribió " + "), devolvemos vacío.
        if not any(terminos_brutos): 
            logging.warning(f"El término de búsqueda '{termino_raw}' quedó vacío después del parseo inicial.")
            return 'OR', [] 
        return op_principal, terminos_brutos
    
    def _procesar_busqueda_via_diccionario(self, termino_original: str, terminos_analizados: List[Dict[str, Any]], op_principal: str,
                                          df_dic_original: pd.DataFrame, df_desc_original: pd.DataFrame,
                                          cols_nombres_dic: List[str]) -> bool:
        # Este es el flujo cuando la búsqueda intenta primero pasar por el diccionario.
        # 1. Busca en el diccionario.
        # 2. Si encuentra algo, extrae términos clave de esas coincidencias.
        # 3. Busca esos términos clave en el archivo de descripciones.
        
        # Aplicamos la máscara al diccionario con los términos analizados
        mascara_diccionario = self.motor._aplicar_mascara_diccionario(df_dic_original, cols_nombres_dic, terminos_analizados, op_principal)
        
        if mascara_diccionario.any(): # ¿Hubo alguna coincidencia en el diccionario?
            self.df_candidato_diccionario = df_dic_original[mascara_diccionario].copy() # Guardamos estas filas
            logging.info(f"Para '{termino_original}', se encontraron {len(self.df_candidato_diccionario)} filas candidatas en el Diccionario (FCD).")
            
            # Ahora extraemos los términos de búsqueda de estas filas del diccionario
            terminos_extraidos_del_diccionario = self.motor._extraer_terminos_diccionario(self.df_candidato_diccionario, cols_nombres_dic)
            
            if terminos_extraidos_del_diccionario: # ¿Conseguimos sacar palabras clave?
                # ¡Sí! Ahora buscamos estas palabras clave en el archivo de descripciones
                resultados_en_desc = self.motor._buscar_terminos_en_descripciones(df_desc_original, terminos_extraidos_del_diccionario, require_all=False) # require_all=False significa OR
                self.df_candidato_descripcion = resultados_en_desc.copy() if resultados_en_desc is not None else pd.DataFrame(columns=df_desc_original.columns)
                self.resultados_actuales = self.df_candidato_descripcion # Estos son los resultados finales que mostraremos
                
                if self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty:
                    # ¡Éxito total! Encontramos en dicc y luego en descripciones
                    self.origen_principal_resultados = OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
                    self._actualizar_estado(f"'{termino_original}': {len(self.df_candidato_diccionario)} en Diccionario, {len(self.df_candidato_descripcion)} resultados en Descripciones.")
                else:
                    # Encontramos en dicc, pero luego nada en descripciones
                    self.origen_principal_resultados = OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                    self._actualizar_estado(f"'{termino_original}': {len(self.df_candidato_diccionario)} en Diccionario, pero 0 resultados en Descripciones.")
                    messagebox.showinfo("Información", 
                                      f"Se encontraron {len(self.df_candidato_diccionario)} filas en el Diccionario para '{termino_original}', "
                                      f"pero no se encontraron coincidencias de esos términos en las Descripciones.\n\n"
                                      f"Esto puede pasar si los términos clave del diccionario no están tal cual en las descripciones.")
            else:
                # Encontramos en dicc, pero no pudimos sacar palabras clave válidas de ahí (raro)
                self.origen_principal_resultados = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                self.df_candidato_descripcion = pd.DataFrame(columns=df_desc_original.columns) # No hay resultados en descripciones
                self.resultados_actuales = self.df_candidato_descripcion
                self._actualizar_estado(f"'{termino_original}': {len(self.df_candidato_diccionario)} en Diccionario, pero no se extrajeron términos válidos para buscar en Descripciones.")
                messagebox.showinfo("Información", 
                                  f"Se encontraron {len(self.df_candidato_diccionario)} filas en el Diccionario para '{termino_original}', "
                                  f"pero no se pudieron extraer términos válidos de ellas para luego buscar en las Descripciones.\n\n"
                                  f"Revisa si las columnas de búsqueda del diccionario están vacías o tienen contenido extraño.")
            return True # Indica que sí hubo coincidencias iniciales en el diccionario
        
        return False # No hubo ninguna coincidencia en el diccionario

    def _procesar_busqueda_directa_descripcion(self, termino_original: str, df_desc_original: pd.DataFrame):
        # Si la búsqueda en el diccionario no dio frutos o se omitió, vamos directos a las descripciones.
        self.ultimo_termino_buscado = f"{termino_original} (Directo)" # Marcamos que fue búsqueda directa
        self._actualizar_estado(f"Buscando '{termino_original}' (Directo) en todas las descripciones...")
        
        res_directos = self.motor.buscar_en_descripciones_directo(termino_original)
        self.df_candidato_descripcion = res_directos.copy() if res_directos is not None else pd.DataFrame(columns=df_desc_original.columns)
        self.resultados_actuales = self.df_candidato_descripcion # Estos son los resultados a mostrar
        self.origen_principal_resultados = OrigenResultados.DIRECTO_DESCRIPCION
        num_rdd = len(self.df_candidato_descripcion) if self.df_candidato_descripcion is not None else 0
        self._actualizar_estado(f"Búsqueda directa de '{termino_original}': {num_rdd} resultados encontrados.")
        if num_rdd == 0:
            messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_original}' en la búsqueda directa en descripciones.")

    def _ejecutar_busqueda(self):
        # ¡El botón de BUSCAR ha sido pulsado! Esta función orquesta todo el proceso.
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", 
                                 "Por favor, cargue el Diccionario y las Descripciones antes de realizar una búsqueda.\n\n"
                                 "El Diccionario es como tu chuleta de términos, y las Descripciones es donde realmente buscamos.")
            return

        termino = self.texto_busqueda_var.get() # Lo que escribió el usuario
        self.ultimo_termino_buscado = termino # Lo guardamos por si quiere salvar la regla

        # Limpiamos el estado de la búsqueda anterior
        self.resultados_actuales = None
        self._actualizar_tabla(self.tabla_resultados, None) 
        self.df_candidato_diccionario = None
        self.df_candidato_descripcion = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        if not termino.strip(): # Si el usuario no escribió nada (o solo espacios)...
            logging.info("Búsqueda vacía. El usuario quiere ver todas las descripciones.")
            df_desc_all = self.motor.datos_descripcion # ...mostramos todas las descripciones.
            self._actualizar_tabla(self.tabla_resultados, df_desc_all)
            self.resultados_actuales = df_desc_all.copy() if df_desc_all is not None else None
            self.df_candidato_descripcion = self.resultados_actuales # Consideramos esto como "candidato" para salvar
            self.origen_principal_resultados = OrigenResultados.DIRECTO_DESCRIPCION_VACIA
            num_filas = len(df_desc_all) if df_desc_all is not None else 0
            self._actualizar_estado(f"Mostrando todas las {num_filas} descripciones disponibles.")
            self._actualizar_botones_estado_general()
            return

        self._actualizar_estado(f"Buscando '{termino}'... ¡Agárrense los machos!")
        df_dic_original = self.motor.datos_diccionario.copy() # Copias para no modificar los originales
        df_desc_original = self.motor.datos_descripcion.copy()
        
        cols_nombres_dic = self.motor._obtener_nombres_columnas_busqueda(df_dic_original)
        if cols_nombres_dic is None: # Si hay un problema con las columnas del diccionario (ya se habrá mostrado error)
            self._actualizar_estado("Error: Configuración de columnas del diccionario inválida. No se puede buscar.")
            self._actualizar_botones_estado_general()
            return

        # 1. Parseamos el término de búsqueda para ver si usa '+', '|', '/'
        op_principal_busqueda, terminos_brutos_busqueda = self._parsear_termino_busqueda_inicial(termino)

        if not terminos_brutos_busqueda: # Si el término es inválido (ej. " + ")
            messagebox.showwarning("Término Inválido", 
                                 "El término de búsqueda está vacío o no es válido.\n\n"
                                 "Asegúrate de que los operadores (+, |, /) tengan algo antes y después, "
                                 "y que no haya cosas raras como '++' o '||'.")
            self._actualizar_estado("Término de búsqueda inválido.")
            self._actualizar_botones_estado_general()
            return

        # 2. Analizamos cada sub-término para ver si es negación, comparación, rango o texto
        terminos_analizados = self.motor._analizar_terminos(terminos_brutos_busqueda)
        if not terminos_analizados: # Si no se pudo analizar ningún término
            messagebox.showwarning("Término Inválido", 
                                 f"No se pudieron analizar los términos en '{termino}'.\n\n"
                                 f"Comprueba la sintaxis de comparaciones (ej. >100, >=50W), rangos (ej. 10-20) "
                                 f"y negaciones (#termino).")
            self._actualizar_estado(f"El análisis de los términos de '{termino}' falló.")
            self._actualizar_botones_estado_general()
            return

        # 3. Intentamos la búsqueda pasando primero por el diccionario
        fcd_encontrados_via_diccionario = self._procesar_busqueda_via_diccionario(
            termino, terminos_analizados, op_principal_busqueda, 
            df_dic_original, df_desc_original, cols_nombres_dic
        )
        self._actualizar_tabla(self.tabla_resultados, self.resultados_actuales) # Mostramos lo que se haya encontrado

        # 4. Si no se encontró NADA en el diccionario que coincidiera...
        if not fcd_encontrados_via_diccionario:
            logging.info(f"El término '{termino}' no produjo coincidencias (o fue completamente negado) en el Diccionario.")
            self._actualizar_estado(f"'{termino}' no encontrado/negado en el Diccionario.")
            
            # Restauramos la vista previa del diccionario a su estado original (todas las filas)
            logging.info("Restaurando la vista previa del diccionario a su estado completo.")
            df_dic_preview = self.motor.datos_diccionario
            if df_dic_preview is not None:
                cols_preview = self.motor._obtener_nombres_columnas_busqueda(df_dic_preview)
                self._actualizar_tabla(self.tabla_diccionario, df_dic_preview, limite_filas=100, columnas_a_mostrar=cols_preview)

            # Preguntamos al usuario si quiere buscar directamente en las descripciones
            if messagebox.askyesno("Término no en Diccionario", 
                                 f"'{termino}' no se encontró en el Diccionario (o las coincidencias fueron negadas).\n\n"
                                 f"¿Desea buscar directamente este término en todas las Descripciones?\n\n"
                                 f"(La búsqueda directa es más simple: busca el texto tal cual, con opción de '+' para AND y '|' o '/' para OR)."):
                self._procesar_busqueda_directa_descripcion(termino, df_desc_original)
                self._actualizar_tabla(self.tabla_resultados, self.resultados_actuales) # Mostramos los resultados directos
            else: # El usuario no quiso buscar directamente
                self._actualizar_estado(f"Búsqueda de '{termino}' cancelada después de no encontrar en diccionario.")
                # Limpiamos todo para que no haya confusión
                self.origen_principal_resultados = OrigenResultados.NINGUNO
                self._actualizar_tabla(self.tabla_resultados, None)
                self.resultados_actuales = None
                self.df_candidato_descripcion = None
                self.df_candidato_diccionario = None
        
        self._actualizar_botones_estado_general() # Actualizamos el estado de los botones (Salvar, Exportar, etc.)
        
        # Si al final tenemos resultados, hacemos una pequeña demo del extractor de magnitudes con la primera celda
        if self.resultados_actuales is not None and not self.resultados_actuales.empty:
            origen_nombre = self.origen_principal_resultados.name if self.origen_principal_resultados != OrigenResultados.NINGUNO else "DESCONOCIDO"
            self._demo_extractor(self.resultados_actuales, origen_nombre)

        # Si la búsqueda fue vía diccionario y tenemos datos, intentamos enfocar en la vista previa
        if self.origen_principal_resultados.es_via_diccionario and \
           self.motor.datos_diccionario is not None and not self.motor.datos_diccionario.empty:
            self._buscar_y_enfocar_en_preview()


    def _demo_extractor(self, df_res: pd.DataFrame, tipo_busqueda: str):
        # Pequeña demo para ver si nuestro ExtractorMagnitud funciona.
        # Coge la primera celda de los resultados y busca magnitudes conocidas.
        # Los resultados se ven en el log, no en la GUI.
        if df_res is None or df_res.empty or len(df_res.columns) == 0: return # Si no hay resultados, no hay demo
        try:
            texto_primera_celda = str(df_res.iloc[0, 0]) # La celda de la primera fila, primera columna
            logging.info(f"--- INICIO DEMO Extractor de Magnitudes (desde búsqueda tipo: {tipo_busqueda}) ---")
            logging.info(f"Texto analizado (primeros 100 chars): '{texto_primera_celda[:100]}...'")
            encontrado_algo = False
            for mag in self.extractor_magnitud.magnitudes: # Probamos con todas nuestras magnitudes predefinidas
                cantidad = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag, texto_primera_celda)
                if cantidad is not None:
                    logging.info(f"  -> Para magnitud '{mag}', se encontró la cantidad: '{cantidad}'")
                    encontrado_algo = True
            if not encontrado_algo:
                logging.info("  (No se encontraron magnitudes predefinidas en este texto de ejemplo)")
            logging.info("--- FIN DEMO Extractor de Magnitudes ---")
        except IndexError: # Por si el DataFrame de resultados está vacío o tiene una estructura rara
            logging.warning(f"Error en la demo del extractor ({tipo_busqueda}): No se pudo acceder a la celda [0,0]. ¿Resultados vacíos o sin columnas?")
        except Exception as e:
            logging.warning(f"Error inesperado durante la demo del extractor ({tipo_busqueda}): {e}")

    def _sanitizar_nombre_archivo(self, texto: str, max_len: int = 50) -> str:
        # Limpia un texto para que se pueda usar como nombre de archivo sin problemas.
        # Quita caracteres raros, espacios extra, y lo acorta si es muy largo.
        if not texto: return "resultados" # Nombre por defecto si no hay texto
        
        # Quitamos caracteres que los sistemas operativos odian en nombres de archivo
        texto_limpio = re.sub(r'[<>:"/\\|?*#]', '_', texto) 
        # Quitamos caracteres de control (como saltos de línea, tabs, etc. que no se ven)
        texto_limpio = "".join(c for c in texto_limpio if c not in string.control) 
        texto_limpio = re.sub(r'\s+', ' ', texto_limpio).strip() # Normalizamos espacios múltiples a uno solo y quitamos de los extremos
        
        texto_cortado = texto_limpio[:max_len] # Lo cortamos si es más largo de max_len
        # A veces al cortar queda un '_' o '.' al final, lo quitamos para que quede más limpio.
        texto_final = texto_cortado.rstrip('._- ') 
        
        if not texto_final: return "resultados" # Si después de limpiar no queda nada, nombre por defecto
        return texto_final

    def _mostrar_dialogo_seleccion_salvado_via_diccionario(self) -> Dict[str, bool]:
        # Cuando los resultados vienen de una búsqueda por diccionario, puede haber dos "sets" de datos:
        # 1. Las filas que coincidieron en el propio diccionario (FCD).
        # 2. Las filas de las descripciones que coincidieron con los términos extraídos de las FCD (RFD).
        # Este diálogo pregunta al usuario cuál de los dos (o ambos) quiere salvar.
        decision = {'confirmed': False, 'save_fcd': False, 'save_rfd': False}
        
        choice_window = tk.Toplevel(self) # Ventanita nueva
        choice_window.title("Elegir Datos a Salvar")
        choice_window.geometry("450x200") 
        choice_window.resizable(False, False) # Que no se pueda cambiar el tamaño
        choice_window.transient(self) # Se mantiene encima de la ventana principal
        choice_window.grab_set() # Bloquea la interacción con la ventana principal hasta que esta se cierre
        
        tk.Label(choice_window, text=f"Para la búsqueda '{self.ultimo_termino_buscado}', elija qué datos desea salvar:").pack(pady=10, padx=10)
        
        var_salvar_fcd = tk.BooleanVar(value=False) # Variable para el checkbox de FCD
        var_salvar_rfd = tk.BooleanVar(value=False) # Variable para el checkbox de RFD
        
        frame_checkboxes = ttk.Frame(choice_window)
        frame_checkboxes.pack(fill=tk.X, padx=20)
        
        # Solo habilitamos los checkboxes si realmente hay datos que salvar para cada tipo
        puede_salvar_fcd = self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty
        # Para RFD, además, el origen tiene que ser el específico que genera RFD válidos.
        puede_salvar_rfd = self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty and \
                           self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
        
        num_fcd = len(self.df_candidato_diccionario or [])
        num_rfd = len(self.df_candidato_descripcion or [])

        chk_fcd_widget = ttk.Checkbutton(frame_checkboxes, 
                                         text=f"Coincidencias del Diccionario ({num_fcd} filas)", 
                                         variable=var_salvar_fcd, 
                                         state="normal" if puede_salvar_fcd else "disabled")
        chk_fcd_widget.pack(anchor=tk.W, pady=2)
        
        chk_rfd_widget = ttk.Checkbutton(frame_checkboxes, 
                                         text=f"Resultados en Descripciones (vía Dic, {num_rfd} filas)", 
                                         variable=var_salvar_rfd, 
                                         state="normal" if puede_salvar_rfd else "disabled")
        chk_rfd_widget.pack(anchor=tk.W, pady=2)
        
        # Si solo una opción es viable, la preseleccionamos y la deshabilitamos para que el usuario no se líe.
        if puede_salvar_fcd and not puede_salvar_rfd: 
            var_salvar_fcd.set(True); chk_fcd_widget.configure(state="disabled")
        elif not puede_salvar_fcd and puede_salvar_rfd: 
            var_salvar_rfd.set(True); chk_rfd_widget.configure(state="disabled")
        
        def on_confirm_choice(): # Cuando el usuario pulsa "Confirmar"
            if not var_salvar_fcd.get() and not var_salvar_rfd.get(): # Si no ha marcado nada...
                messagebox.showwarning("Ninguna Selección", "Por favor, seleccione al menos una opción para salvar.", parent=choice_window)
                return
            decision.update({'confirmed': True, 'save_fcd': var_salvar_fcd.get(), 'save_rfd': var_salvar_rfd.get()})
            choice_window.destroy() # Cerramos la ventanita
            
        ttk.Button(choice_window, text="Confirmar y Salvar Selección", command=on_confirm_choice).pack(pady=15)
        
        self.wait_window(choice_window) # Esperamos a que el usuario interactúe y se cierre la ventana
        return decision

    def _salvar_regla_actual(self):
        # El usuario quiere guardar los resultados de la búsqueda actual como una "regla".
        # Esto simplemente añade los DataFrames relevantes a una lista en memoria.
        origen_nombre = self.origen_principal_resultados.name if self.origen_principal_resultados != OrigenResultados.NINGUNO else "NINGUNO"
        logging.info(f"Intentando salvar regla. Origen principal de resultados: {origen_nombre}, Último término buscado: '{self.ultimo_termino_buscado}'")

        if not self.ultimo_termino_buscado: # Debería haberse guardado en _ejecutar_busqueda
            messagebox.showerror("Error Interno", "No hay un término de búsqueda asociado a estos resultados. No se puede salvar la regla.")
            return
        
        # Comprobamos si realmente hay algo que salvar
        puede_salvar_fcd = self.df_candidato_diccionario is not None and not self.df_candidato_diccionario.empty
        puede_salvar_rfd_o_rdd = self.df_candidato_descripcion is not None and not self.df_candidato_descripcion.empty

        if not puede_salvar_fcd and not puede_salvar_rfd_o_rdd:
            messagebox.showwarning("Sin Datos Salvables", "No hay datos de la búsqueda actual que se puedan salvar. Intente otra búsqueda.")
            return

        salvo_algo_en_esta_operacion = False
        timestamp_actual = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S") # Para todas las reglas de esta "salvada"

        if self.origen_principal_resultados.es_via_diccionario:
            # Si la búsqueda pasó por el diccionario, preguntamos al usuario qué quiere salvar
            decision = self._mostrar_dialogo_seleccion_salvado_via_diccionario()
            if decision['confirmed']:
                if decision['save_fcd'] and puede_salvar_fcd:
                    self.reglas_guardadas.append({
                        'termino_busqueda': self.ultimo_termino_buscado,
                        'tipo_fuente': "DICCIONARIO_COINCIDENCIAS", # Un nombre especial para identificar estas
                        'datos_relevantes': self.df_candidato_diccionario.copy(), # ¡Importante copiarlo!
                        'timestamp': timestamp_actual 
                    })
                    salvo_algo_en_esta_operacion = True
                    logging.info(f"Regla salvada (Coincidencias del Diccionario) para: '{self.ultimo_termino_buscado}'")
                
                # Solo salvamos RFD si realmente vienen de una búsqueda exitosa en diccionario Y luego en descripciones
                if decision['save_rfd'] and (self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC and puede_salvar_rfd_o_rdd):
                    self.reglas_guardadas.append({
                        'termino_busqueda': self.ultimo_termino_buscado,
                        'tipo_fuente': OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC.name, # Usamos el nombre del Enum
                        'datos_relevantes': self.df_candidato_descripcion.copy(),
                        'timestamp': timestamp_actual 
                    })
                    salvo_algo_en_esta_operacion = True
                    logging.info(f"Regla salvada (Resultados en Descripciones vía Dic) para: '{self.ultimo_termino_buscado}'")
        
        elif self.origen_principal_resultados.es_directo_descripcion: # Si fue búsqueda directa (o vacía)
            if puede_salvar_rfd_o_rdd: # Solo hay un tipo de datos que salvar aquí (las descripciones)
                self.reglas_guardadas.append({
                    'termino_busqueda': self.ultimo_termino_buscado,
                    'tipo_fuente': self.origen_principal_resultados.name,
                    'datos_relevantes': self.df_candidato_descripcion.copy(),
                    'timestamp': timestamp_actual
                })
                salvo_algo_en_esta_operacion = True
                logging.info(f"Regla salvada (Búsqueda Directa/Vacía en Descripciones) para: '{self.ultimo_termino_buscado}'")
            else: # Esto no debería pasar si el botón de salvar estaba bien habilitado
                messagebox.showwarning("Sin Datos", "No hay resultados de la búsqueda directa o vacía para salvar.")
        
        else: # Origen desconocido o no manejado para salvar
            if self.origen_principal_resultados != OrigenResultados.NINGUNO: # Si no es NINGUNO, es un caso raro
                messagebox.showerror("Error Inesperado", f"No se puede determinar qué salvar para el origen de resultados: {self.origen_principal_resultados.name}.")

        if salvo_algo_en_esta_operacion:
            num_total_reglas = len(self.reglas_guardadas)
            self._actualizar_estado(f"¡Regla(s) nueva(s) guardada(s)! Total acumulado: {num_total_reglas}.")
        else: # Si no se confirmó el diálogo o no había nada que salvar bajo las condiciones
            self._actualizar_estado("Ninguna regla fue salvada en esta operación.")
        
        # Deshabilitamos el botón de salvar para evitar doble click o confusión,
        # se reactivará si es posible en la próxima actualización de botones.
        self.btn_salvar_regla["state"] = "disabled" 
        self._actualizar_botones_estado_general() # Reevaluamos el estado de todos los botones


    def _exportar_resultados(self):
        # El usuario quiere exportar todas las reglas que ha ido guardando a un Excel.
        if not self.reglas_guardadas:
            messagebox.showwarning("Sin Reglas", "Aún no has guardado ninguna regla para exportar. Usa 'Salvar Regla' después de una búsqueda.")
            return
        
        timestamp_export = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S") # Para un nombre de archivo único
        nombre_sugerido_base = f"exportacion_reglas_{timestamp_export}"
        tipos_archivo = [("Archivo Excel (*.xlsx)", "*.xlsx")]

        # Preguntamos dónde quiere guardar el archivo
        ruta_guardar = filedialog.asksaveasfilename(
            title="Exportar Reglas Guardadas Como...",
            initialfile=f"{nombre_sugerido_base}.xlsx", # Nombre sugerido
            defaultextension=".xlsx", 
            filetypes=tipos_archivo
        )
        if not ruta_guardar: # Si cancela...
            logging.info("Exportación de reglas cancelada por el usuario."); self._actualizar_estado("Exportación de reglas cancelada."); return

        self._actualizar_estado("Exportando reglas guardadas... Un momentito..."); num_reglas = len(self.reglas_guardadas)
        logging.info(f"A punto de exportar {num_reglas} regla(s) a: {ruta_guardar}")
        
        try:
            extension = ruta_guardar.split('.')[-1].lower()
            if extension == 'xlsx': # Solo soportamos Excel por ahora
                with pd.ExcelWriter(ruta_guardar, engine='openpyxl') as writer: # Usamos ExcelWriter para meter varias hojas
                    # Hoja 1: Un índice de todas las reglas exportadas
                    datos_indice = []
                    for i, regla in enumerate(self.reglas_guardadas):
                        df_datos_regla = regla.get('datos_relevantes')
                        num_filas_datos = len(df_datos_regla) if df_datos_regla is not None else 0
                        datos_indice.append({
                            "ID_Regla_Hoja": f"R{i+1}", # Para que sepa a qué hoja ir
                            "Termino_Busqueda": regla.get('termino_busqueda', 'N/A'),
                            "Fuente_Datos": regla.get('tipo_fuente', 'N/A'),
                            "Filas_Resultado": num_filas_datos,
                            "Timestamp_Guardado": regla.get('timestamp', 'N/A')
                        })
                    df_indice = pd.DataFrame(datos_indice)
                    if not df_indice.empty:
                        df_indice.to_excel(writer, sheet_name="Indice_Reglas", index=False)

                    # Ahora, una hoja por cada regla guardada
                    for i, regla in enumerate(self.reglas_guardadas):
                        df_regla_datos = regla.get('datos_relevantes')
                        if df_regla_datos is not None and isinstance(df_regla_datos, pd.DataFrame) and not df_regla_datos.empty:
                            # Creamos un nombre de hoja cortito y descriptivo
                            term_sanitizado = self._sanitizar_nombre_archivo(regla.get('termino_busqueda','SinTermino'), max_len=15)
                            fuente_abbr = regla.get('tipo_fuente','FuenteDesc')[:3].upper() # Abreviatura de la fuente
                            nombre_hoja_base = f"R{i+1}_{fuente_abbr}_{term_sanitizado}"
                            nombre_hoja = nombre_hoja_base[:31] # Excel tiene un límite de 31 chars para nombres de hoja

                            # Por si hay nombres de hoja duplicados (poco probable pero posible)
                            original_nombre_hoja = nombre_hoja; count = 1
                            # writer.book.sheetnames nos da los nombres de las hojas ya creadas en este ExcelWriter
                            while nombre_hoja in writer.book.sheetnames: 
                                nombre_hoja = f"{original_nombre_hoja[:28]}_{count}"; count +=1 
                                if count > 50 : nombre_hoja = f"ErrorNombreHoja{i}"; break # Evitar bucle infinito
                            
                            try:
                                df_regla_datos.to_excel(writer, sheet_name=nombre_hoja, index=False)
                            except Exception as e_sheet:
                                logging.error(f"Error al escribir la hoja '{nombre_hoja}' para la regla R{i+1}: {e_sheet}")
                        else:
                            logging.warning(f"La Regla R{i+1} (buscando '{regla.get('termino_busqueda')}') no tenía datos para exportar a una hoja.")
            else: # Si no es .xlsx
                messagebox.showerror("Formato No Soportado", "Lo sentimos, solo se soporta la exportación de reglas a formato Excel (.xlsx).")
                self._actualizar_estado("Exportación fallida: formato no soportado.")
                return

            logging.info(f"¡Exportación de {num_reglas} regla(s) completada con éxito a {ruta_guardar}!")
            messagebox.showinfo("Exportación Exitosa", f"{num_reglas} regla(s) han sido exportadas a:\n{ruta_guardar}")
            self._actualizar_estado(f"Reglas exportadas a {os.path.basename(ruta_guardar)}.")

            # Preguntamos si quiere limpiar las reglas guardadas ahora que ya las exportó
            if messagebox.askyesno("Limpiar Reglas Guardadas", "Exportación exitosa.\n¿Desea limpiar las reglas guardadas internamente ahora?"):
                self.reglas_guardadas.clear()
                self._actualizar_estado("Reglas guardadas internamente han sido limpiadas.")
                logging.info("Reglas guardadas limpiadas por el usuario después de la exportación.")
            self._actualizar_botones_estado_general() # Actualizamos botones (ej. Exportar se desactivará)

        except Exception as e:
            logging.exception("Error inesperado durante la exportación de reglas.")
            messagebox.showerror("Error al Exportar", f"Ocurrió un error inesperado al exportar:\n{e}")
            self._actualizar_estado("Error exportando reglas.")

    def _actualizar_estado_botones_operadores(self):
        """
        Esta función es clave para la UX: activa o desactiva los botones de operadores (+, #, >, etc.)
        dependiendo de lo que el usuario ya haya escrito en el campo de búsqueda y dónde esté el cursor.
        Intenta evitar que el usuario escriba cosas sin sentido como "##termino" o ">>100".
        """
        texto_completo = self.texto_busqueda_var.get() # Lo que hay escrito
        cursor_pos = self.entrada_busqueda.index(tk.INSERT) # Dónde está el cursor

        # Intentamos aislar el "segmento lógico" actual donde está el cursor.
        # Un segmento es lo que hay entre operadores principales como '+' o '|', o desde el inicio.
        inicio_segmento = 0
        # Buscamos el último '+' o '|' ANTES del cursor
        temp_pos_plus = texto_completo.rfind('+', 0, cursor_pos)
        if temp_pos_plus != -1: inicio_segmento = temp_pos_plus + 1
        
        temp_pos_pipe = texto_completo.rfind('|', 0, cursor_pos)
        if temp_pos_pipe != -1: inicio_segmento = max(inicio_segmento, temp_pos_pipe + 1)
        
        # El segmento actual, quitando espacios al principio (ej. si escribió "termino1 +     #termino2")
        segmento_actual_con_espacios_delante = texto_completo[inicio_segmento:cursor_pos]
        segmento_actual_limpio = segmento_actual_con_espacios_delante.lstrip()

        # Analizamos este segmento para ver qué tipo de término es (comparación, rango, etc.)
        # Esto nos ayuda a decidir si podemos añadir otro operador de comparación o rango.
        terminos_analizados_segmento = self.motor._analizar_terminos([segmento_actual_limpio.strip()])
        es_termino_comparativo_o_rango = False
        if terminos_analizados_segmento: # Si se pudo analizar...
            tipo_termino_segmento = terminos_analizados_segmento[0].get('tipo')
            if tipo_termino_segmento in ['gt', 'lt', 'ge', 'le', 'range']:
                es_termino_comparativo_o_rango = True
        
        # También miramos si visualmente ya hay un operador de comparación o rango, por si el parseo aún no lo pilla bien
        # (ej. si el usuario está a medio escribir ">10").
        tiene_comparacion_o_rango_visible_en_segmento = any(op in segmento_actual_limpio for op in ['>', '<', '>=', '<=', '-'])
        
        tiene_negacion_en_segmento = segmento_actual_limpio.startswith('#')

        # --- Ahora, decidimos el estado de cada botón ---

        # Botón NOT (#): Solo se puede poner uno al principio del segmento.
        self.btn_not['state'] = 'disabled' if tiene_negacion_en_segmento else 'normal'
        
        # Botones de comparación (>, <, >=, <=): Desactivados si ya hay uno en el segmento.
        estado_comparacion = 'disabled' if es_termino_comparativo_o_rango or tiene_comparacion_o_rango_visible_en_segmento else 'normal'
        for btn in [self.btn_gt, self.btn_lt, self.btn_ge, self.btn_le]:
            btn.config(state=estado_comparacion)
        
        # Botón de rango (-): Desactivado si no hay nada escrito antes, o si ya hay comparación/rango.
        estado_rango = 'disabled'
        # Necesita algo antes del guion para formar "numero-numero"
        # Y no puede haber ya un operador de comparación/rango.
        caracter_antes_cursor_es_digito = False
        if cursor_pos > inicio_segmento + len(segmento_actual_con_espacios_delante) - len(segmento_actual_limpio): # Estamos después de los espacios iniciales
            if segmento_actual_limpio and segmento_actual_limpio[-1].isdigit():
                 caracter_antes_cursor_es_digito = True

        if caracter_antes_cursor_es_digito and not (es_termino_comparativo_o_rango or tiene_comparacion_o_rango_visible_en_segmento):
            estado_rango = 'normal'
        self.btn_range.config(state=estado_rango)
        
        # Botones lógicos (+, |): Desactivados si la caja está vacía o si lo último escrito es ya un '+' o '|'.
        # (Para evitar "termino ++" o "termino ||")
        texto_limpio_al_final = texto_completo.rstrip() # Quitamos espacios del final del TODO el texto
        estado_logico = 'normal'
        if not texto_completo.strip(): # Si no hay nada escrito (o solo espacios)
            estado_logico = 'disabled'
        elif texto_limpio_al_final and texto_limpio_al_final[-1] in ['+', '|']: # Si lo último es un operador lógico...
            estado_logico = 'disabled' # ...no podemos poner otro seguido.
        
        self.btn_and.config(state=estado_logico)
        self.btn_or.config(state=estado_logico)

    def on_closing(self):
        # Esto se ejecuta cuando el usuario cierra la ventana.
        # Aprovechamos para guardar la configuración.
        logging.info("Cerrando la aplicación... ¡Hasta la próxima!")
        self._guardar_configuracion()
        self.destroy() # Cerramos la ventana de Tkinter

    def _insertar_operador_validado(self, operador: str):
        """Inserta un operador en la caja de búsqueda si tiene sentido hacerlo en ese momento."""
        # Primero, si no hay diccionario cargado, los operadores no deberían hacer nada.
        if self.motor.datos_diccionario is None:
            logging.warning("Intento de insertar operador sin diccionario cargado. No se hace nada.")
            return

        texto_actual = self.texto_busqueda_var.get()
        cursor_pos = self.entrada_busqueda.index(tk.INSERT) # Dónde está el cursor
        
        # Lógica similar a _actualizar_estado_botones_operadores para ver el contexto del segmento actual
        inicio_segmento = 0
        temp_pos_plus = texto_actual.rfind('+', 0, cursor_pos)
        if temp_pos_plus != -1: inicio_segmento = temp_pos_plus + 1
        temp_pos_pipe = texto_actual.rfind('|', 0, cursor_pos)
        if temp_pos_pipe != -1: inicio_segmento = max(inicio_segmento, temp_pos_pipe + 1)
        
        segmento_actual_cursor = texto_actual[inicio_segmento:cursor_pos].strip()
        
        puede_insertar = True # Por defecto, pensamos que sí se puede

        # --- Reglas para NO insertar ---
        if operador in ['>', '<', '>=', '<=']:
            # No si ya hay una comparación o rango en el segmento donde está el cursor
            if any(op in segmento_actual_cursor for op in ['>', '<', '>=', '<=', '-']):
                puede_insertar = False
        elif operador == '-':
            # No si no hay nada antes del cursor en este segmento, o si ya hay comparación/rango
            if not segmento_actual_cursor or any(op in segmento_actual_cursor for op in ['>', '<', '>=', '<=', '-']):
                puede_insertar = False
            # Y el carácter justo antes del cursor debe ser un dígito
            elif not (texto_actual[cursor_pos-1].isdigit() if cursor_pos > 0 else False):
                puede_insertar = False
        elif operador in ['+', '|']:
            # No si el texto total está vacío, o si justo antes del cursor (ignorando espacios) hay otro '+' o '|'.
            if not texto_actual.strip() or (texto_actual.rstrip() and texto_actual.rstrip()[-1] in ['+', '|']):
                puede_insertar = False
        elif operador == '#':
            # No si el segmento donde está el cursor ya empieza con '#'
            if segmento_actual_cursor.startswith('#'):
                puede_insertar = False
        
        if puede_insertar:
            # Si todas las validaciones pasan, insertamos el operador
            nuevo_texto = texto_actual[:cursor_pos] + operador + texto_actual[cursor_pos:]
            self.texto_busqueda_var.set(nuevo_texto)
            # Y movemos el cursor al final de lo que acabamos de insertar
            self.entrada_busqueda.icursor(cursor_pos + len(operador))
            # Forzamos la actualización del estado de los botones porque el texto ha cambiado
            # (aunque el trace_add ya lo haría, esto asegura inmediatez si el trace es asíncrono)
            self._actualizar_estado_botones_operadores() 
        else:
            logging.debug(f"Inserción del operador '{operador}' bloqueada por validación.")


    def _deshabilitar_botones_operadores(self):
        """Un pequeño ayudante para apagar todos los botones de operadores de golpe."""
        self.btn_and['state'] = 'disabled'
        self.btn_or['state'] = 'disabled'
        self.btn_not['state'] = 'disabled'
        self.btn_gt['state'] = 'disabled'
        self.btn_lt['state'] = 'disabled'
        self.btn_ge['state'] = 'disabled'
        self.btn_le['state'] = 'disabled'
        self.btn_range['state'] = 'disabled'

# --- Bloque Principal (`if __name__ == "__main__":`) ---
if __name__ == "__main__":
    # Esto solo se ejecuta si corremos este archivo directamente (no si lo importamos como módulo)
    
    log_file = 'buscador_app.log' # Guardaremos un log de lo que pasa
    logging.basicConfig( # Configuramos el logging
        level=logging.DEBUG, # Nivel de detalle: DEBUG es el más alto (más cotilla)
        format='%(asctime)s - %(filename)s:%(lineno)d - %(levelname)s - %(message)s', # Cómo se ven los mensajes
        handlers=[ # Dónde van los mensajes
            logging.FileHandler(log_file, encoding='utf-8', mode='w'), # A un archivo (modo 'w' borra el anterior)
            logging.StreamHandler() # Y a la consola
        ]
    )
    logging.info("=============================================")
    logging.info("=== ¡Arrancando el Buscador Avanzado Molón! ===")
    logging.info(f"Plataforma: {platform.system()} {platform.release()}")
    logging.info(f"Versión de Python: {platform.python_version()}")
    logging.info("=============================================")

    # Comprobamos si tenemos las librerías que necesitamos (Pandas y openpyxl)
    missing_deps = []
    try: import pandas as pd; logging.info(f"Pandas versión: {pd.__version__}")
    except ImportError: missing_deps.append("pandas"); logging.critical("¡Falta Pandas! No podemos seguir sin él.")
    try: import openpyxl; logging.info(f"openpyxl versión: {openpyxl.__version__}")
    except ImportError: missing_deps.append("openpyxl"); logging.critical("¡Falta openpyxl! Necesario para los .xlsx.")

    if missing_deps: # Si falta alguna...
        error_msg = f"¡Oh, no! Faltan librerías importantes: {', '.join(missing_deps)}.\nPor favor, instálalas con: pip install {' '.join(missing_deps)}"
        logging.critical(error_msg)
        try: # Intentamos mostrar un error bonito en una ventana
            root_temp = tk.Tk(); root_temp.withdraw(); messagebox.showerror("Dependencias Faltantes", error_msg); root_temp.destroy()
        except tk.TclError: # Si ni Tkinter funciona, pues a la consola
            print(f"ERROR CRÍTICO: {error_msg}")
        exit(1) # Cerramos la aplicación
    
    # ¡Todo listo! Lanzamos la aplicación.
    try:
        app = InterfazGrafica() # Creamos nuestra ventana
        app.mainloop() # Y la ponemos en marcha (esto bloquea hasta que se cierra la ventana)
    except Exception as main_error: # Si algo explota de forma inesperada y no lo hemos pillado antes...
        logging.critical("¡¡¡ERROR FATAL NO CAPTURADO EN LA APLICACIÓN!!!", exc_info=True)
        try: # Intentamos mostrarlo en una ventana de error
            root_err = tk.Tk(); root_err.withdraw()
            messagebox.showerror("Error Fatal Inesperado", f"Ha ocurrido un error crítico que no esperábamos:\n{main_error}\n\nPor favor, consulta el archivo '{log_file}' para más detalles técnicos.")
            root_err.destroy()
        except Exception as fallback_error: # Si ni eso funciona...
            logging.error(f"No se pudo ni mostrar el mensaje de error fatal en una ventana: {fallback_error}")
            print(f"ERROR FATAL INESPERADO: {main_error}. Revisa el archivo {log_file} para más detalles.")
    finally:
        logging.info("=== Aplicación Buscador Finalizada. ¡Apagando luces! ===")
