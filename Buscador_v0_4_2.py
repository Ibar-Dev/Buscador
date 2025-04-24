import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Tuple, Union, Set, Callable
import traceback
import platform
import unicodedata # >>> Import necesario para ExtractorMagnitud <<<

# --- Clases de Lógica ---

class ManejadorExcel:
    @staticmethod
    def cargar_excel(ruta: str) -> Optional[pd.DataFrame]:
        try:
            return pd.read_excel(ruta, engine='openpyxl' if ruta.endswith('.xlsx') else None)
        except Exception as e:
            messagebox.showerror("Error al Cargar Archivo",
                                 f"No se pudo cargar el archivo:\n{ruta}\n\n"
                                 f"Error: {e}\n\n"
                                 "Asegúrese de que el archivo no esté corrupto y "
                                 "tenga instalado 'openpyxl' (`pip install openpyxl`) para archivos .xlsx.")
            return None

class MotorBusqueda:
    # >>> (Sin cambios en MotorBusqueda) <<<
    def __init__(self):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[str] = None
        self.archivo_descripcion_actual: Optional[str] = None

    def cargar_excel_diccionario(self, ruta: str) -> bool:
        self.datos_diccionario = ManejadorExcel.cargar_excel(ruta)
        self.archivo_diccionario_actual = ruta if self.datos_diccionario is not None else None
        return self.datos_diccionario is not None

    def cargar_excel_descripcion(self, ruta: str) -> bool:
        self.datos_descripcion = ManejadorExcel.cargar_excel(ruta)
        self.archivo_descripcion_actual = ruta if self.datos_descripcion is not None else None
        return self.datos_descripcion is not None

    def _obtener_columnas_busqueda(self, df: pd.DataFrame) -> Optional[List[str]]:
        if df is None:
            messagebox.showerror("Error Interno", "Se intentó obtener columnas de un DataFrame nulo (diccionario).")
            return None
        columnas_disponibles = df.columns
        if len(columnas_disponibles) == 0:
            messagebox.showerror("Error de Diccionario", "El archivo de diccionario está vacío o no tiene columnas.")
            return None
        cols_encontradas = []
        if len(columnas_disponibles) > 0:
            cols_encontradas.append(columnas_disponibles[0])
        if len(columnas_disponibles) > 3:
            cols_encontradas.append(columnas_disponibles[3])
        if len(cols_encontradas) == 1 and len(columnas_disponibles) <= 3:
            messagebox.showwarning("Advertencia de Diccionario",
                                   f"El diccionario tiene menos de 4 columnas. Se buscará solo en la columna: '{cols_encontradas[0]}'.")
        return cols_encontradas

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, columnas_busqueda: List[str]) -> Set[str]:
        terminos_encontrados: Set[str] = set()
        if df_coincidencias is None or df_coincidencias.empty:
            return terminos_encontrados
        columnas_validas = [col for col in columnas_busqueda if col in df_coincidencias.columns]
        if not columnas_validas:
            print(f"Advertencia: Ninguna de las columnas de búsqueda {columnas_busqueda} encontrada en df_coincidencias.")
            return terminos_encontrados
        for col in columnas_validas:
            try:
                terminos_col = df_coincidencias[col].dropna().astype(str).str.upper().unique()
                terminos_encontrados.update(terminos_col)
            except Exception as e:
                print(f"Error extrayendo términos de columna '{col}': {e}. Intentando conversión simple.")
                try:
                    terminos_col = df_coincidencias[col].astype(str).str.upper().unique()
                    terminos_encontrados.update(terminos_col)
                    terminos_encontrados.discard('NAN')
                except Exception as e_fallback:
                    print(f"Error en fallback extrayendo términos de columna '{col}': {e_fallback}")
        return terminos_encontrados

    def _buscar_terminos_en_descripciones(self,
                                          df_descripcion: pd.DataFrame,
                                          terminos_a_buscar: Set[str],
                                          require_all: bool = False) -> pd.DataFrame:
        if df_descripcion is None or df_descripcion.empty or not terminos_a_buscar:
            columnas = df_descripcion.columns if df_descripcion is not None else []
            return pd.DataFrame(columns=columnas)
        agg_func: Callable = all if require_all else any
        try:
            texto_filas = df_descripcion.apply(lambda fila: ' '.join(fila.astype(str).values).upper(), axis=1)
            terminos_validos = {t for t in terminos_a_buscar if t}
            if not terminos_validos:
                print("Advertencia: No hay términos válidos para buscar en descripciones.")
                return pd.DataFrame(columns=df_descripcion.columns)
            mascara_descripcion = texto_filas.apply(
                lambda texto_fila: agg_func(termino in texto_fila for termino in terminos_validos)
            )
            resultados = df_descripcion[mascara_descripcion]
        except Exception as e:
            messagebox.showerror("Error en Búsqueda de Descripciones", f"Ocurrió un error al filtrar descripciones:\n{e}")
            traceback.print_exc()
            return pd.DataFrame(columns=df_descripcion.columns)
        return resultados

    def buscar(self, termino_buscado: str) -> Union[None, pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        if self.datos_diccionario is None:
            messagebox.showwarning("Diccionario No Cargado", "Por favor, cargue primero el archivo 'Diccionario'.")
            return None
        if self.datos_descripcion is None:
            messagebox.showwarning("Descripciones No Cargadas", "Por favor, cargue primero el archivo 'Descripciones'.")
            return None
        termino_limpio = termino_buscado.strip().upper()
        if not termino_limpio:
            return self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame()
        df_diccionario = self.datos_diccionario.copy() if self.datos_diccionario is not None else pd.DataFrame()
        df_descripcion = self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame()
        if df_diccionario.empty:
            messagebox.showerror("Error", "El DataFrame del diccionario está vacío.")
            return None
        if df_descripcion.empty:
            messagebox.showerror("Error", "El DataFrame de descripciones está vacío.")
            return None
        try:
            if '+' in termino_limpio:
                return self._busqueda_con_and(df_diccionario, df_descripcion, termino_limpio)
            elif '-' in termino_limpio:
                return self._busqueda_con_or(df_diccionario, df_descripcion, termino_limpio)
            else:
                return self._busqueda_simple(df_diccionario, df_descripcion, termino_limpio)
        except Exception as e:
            messagebox.showerror("Error Inesperado en Búsqueda", f"Ocurrió un error durante la búsqueda:\n{e}")
            traceback.print_exc()
            return (df_diccionario, df_descripcion)

    def _aplicar_mascara_diccionario(self, df: pd.DataFrame, columnas: List[str], palabras: List[str], operador: str) -> pd.Series:
        if df is None or df.empty or not columnas or not palabras:
            return pd.Series(False, index=df.index if df is not None else None)
        columnas_validas = [col for col in columnas if col in df.columns]
        if not columnas_validas:
            print(f"Advertencia: Ninguna de las columnas {columnas} existe en el DataFrame del diccionario.")
            return pd.Series(False, index=df.index)
        try:
            df_str = df[columnas_validas].astype(str)
        except Exception as e:
            print(f"Error convirtiendo columnas {columnas_validas} a string: {e}")
            return pd.Series(False, index=df.index)
        if operador == 'AND':
            mascara_total = pd.Series(True, index=df.index)
        else:
            mascara_total = pd.Series(False, index=df.index)
        palabras_escapadas = [re.escape(p) for p in palabras]
        if operador == 'SIMPLE':
            palabra_regex = rf"\b{palabras_escapadas[0]}\b"
            for col in columnas_validas:
                try:
                    mascara_total |= df_str[col].str.contains(palabra_regex, regex=True, na=False, case=False)
                except Exception as e:
                    print(f"Error aplicando regex '{palabra_regex}' en columna '{col}': {e}")
        elif operador == 'OR':
            for palabra_esc in palabras_escapadas:
                palabra_regex = rf"\b{palabra_esc}\b"
                mascara_palabra = pd.Series(False, index=df.index)
                for col in columnas_validas:
                    try:
                        mascara_palabra |= df_str[col].str.contains(palabra_regex, regex=True, na=False, case=False)
                    except Exception as e:
                        print(f"Error aplicando regex OR '{palabra_regex}' en columna '{col}': {e}")
                mascara_total |= mascara_palabra
        elif operador == 'AND':
            for palabra_esc in palabras_escapadas:
                palabra_regex = rf"\b{palabra_esc}\b"
                mascara_palabra = pd.Series(False, index=df.index)
                for col in columnas_validas:
                    try:
                        mascara_palabra |= df_str[col].str.contains(palabra_regex, regex=True, na=False, case=False)
                    except Exception as e:
                        print(f"Error aplicando regex AND '{palabra_regex}' en columna '{col}': {e}")
                mascara_total &= mascara_palabra
        return mascara_total

    def _busqueda_base(self, df_diccionario: pd.DataFrame, df_descripcion: pd.DataFrame, termino: str, operador: str, require_all_desc: bool) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame, str], Tuple[pd.DataFrame, pd.DataFrame]]:
        columnas_busqueda = self._obtener_columnas_busqueda(df_diccionario)
        if columnas_busqueda is None:
            return (df_diccionario, df_descripcion, "ErrorColumnas")
        if operador == 'AND': palabras = [p.strip() for p in termino.split('+') if p.strip()]
        elif operador == 'OR': palabras = [p.strip() for p in termino.split('-') if p.strip()]
        else: palabras = [termino.strip()]
        if not palabras or all(not p for p in palabras):
            return (df_diccionario, df_descripcion, "ErrorTermino")
        mascara_diccionario = self._aplicar_mascara_diccionario(df_diccionario, columnas_busqueda, palabras, operador)
        if not mascara_diccionario.any():
            return (df_diccionario, df_descripcion)
        coincidencias_diccionario = df_diccionario[mascara_diccionario]
        terminos_encontrados = self._extraer_terminos_diccionario(coincidencias_diccionario, columnas_busqueda)
        if not terminos_encontrados:
            messagebox.showinfo("Aviso", f"Se encontraron {len(coincidencias_diccionario)} fila(s) en el diccionario para '{termino}', pero no términos extraíbles en las columnas de búsqueda (posiblemente celdas vacías o NaN). No se buscará en descripciones.")
            return pd.DataFrame(columns=df_descripcion.columns if df_descripcion is not None else None)
        resultados_descripcion = self._buscar_terminos_en_descripciones(df_descripcion, terminos_encontrados, require_all=require_all_desc)
        return resultados_descripcion

    def _busqueda_simple(self, df_diccionario: pd.DataFrame, df_descripcion: pd.DataFrame, termino: str) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        return self._busqueda_base(df_diccionario, df_descripcion, termino, 'SIMPLE', require_all_desc=False)

    def _busqueda_con_and(self, df_diccionario: pd.DataFrame, df_descripcion: pd.DataFrame, termino: str) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        return self._busqueda_base(df_diccionario, df_descripcion, termino, 'AND', require_all_desc=False)

    def _busqueda_con_or(self, df_diccionario: pd.DataFrame, df_descripcion: pd.DataFrame, termino: str) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        return self._busqueda_base(df_diccionario, df_descripcion, termino, 'OR', require_all_desc=False)

    def buscar_en_descripciones_directo(self, termino_buscado: str) -> pd.DataFrame:
        if self.datos_descripcion is None or self.datos_descripcion.empty:
            messagebox.showwarning("Descripciones No Cargadas", "No hay datos de descripciones cargados para buscar.")
            return pd.DataFrame()
        termino_limpio = termino_buscado.strip().upper()
        if not termino_limpio:
            return self.datos_descripcion.copy()
        df_descripcion = self.datos_descripcion.copy()
        resultados = pd.DataFrame(columns=df_descripcion.columns)
        try:
            texto_filas = df_descripcion.apply(lambda fila: ' '.join(fila.astype(str).values).upper(), axis=1)
            mascara = pd.Series(False, index=df_descripcion.index)
            if '+' in termino_limpio:
                palabras = [p.strip() for p in termino_limpio.split('+') if p.strip()]
                if not palabras: return resultados
                mascara = pd.Series(True, index=df_descripcion.index)
                for palabra in palabras:
                    palabra_regex = rf"\b{re.escape(palabra)}\b"
                    mascara &= texto_filas.str.contains(palabra_regex, regex=True, na=False)
            elif '-' in termino_limpio:
                palabras = [p.strip() for p in termino_limpio.split('-') if p.strip()]
                if not palabras: return resultados
                for palabra in palabras:
                    palabra_regex = rf"\b{re.escape(palabra)}\b"
                    mascara |= texto_filas.str.contains(palabra_regex, regex=True, na=False)
            else:
                palabra_regex = rf"\b{re.escape(termino_limpio)}\b"
                mascara = texto_filas.str.contains(palabra_regex, regex=True, na=False)
            resultados = df_descripcion[mascara]
        except Exception as e:
            messagebox.showerror("Error en Búsqueda Directa", f"Ocurrió un error al buscar directamente en descripciones:\n{e}")
            traceback.print_exc()
            return pd.DataFrame(columns=df_descripcion.columns)
        return resultados


# >>> INICIO: Definición de la Clase ExtractorMagnitud <<<
class ExtractorMagnitud:
    """
    Clase para buscar y extraer cantidades numéricas asociadas a magnitudes
    específicas dentro de un texto descriptivo.
    """
    MAGNITUDES_PREDEFINIDAS: List[str] = [
        "A","AMP","AMPS","AH","ANTENNA","BASE","BIT","ETH","FE","G","GB",
        "GBE","GE","GIGABIT","GBASE","GBASEWAN","GBIC","GBIT","GBPS","GH",
        "GHZ","HZ","KHZ","KM","KVA","KW","LINEAS","LINES","MHZ","NM","PORT",
        "PORTS","PTOS","PUERTO","PUERTOS","P","V","VA","VAC","VC","VCC",
        "VCD","VDC","W","WATTS","E","FE","GBE","GE","POTS","STM"
    ]

    def __init__(self, magnitudes: Optional[List[str]] = None):
        """
        Inicializa el extractor.
        Args:
            magnitudes: Una lista opcional de magnitudes a usar. Si es None, usa MAGNITUDES_PREDEFINIDAS.
        """
        self.magnitudes = magnitudes if magnitudes is not None else self.MAGNITUDES_PREDEFINIDAS

    @staticmethod
    def _quitar_diacronicos_y_acentos(texto: str) -> str:
        """Elimina acentos y marcas diacríticas del texto."""
        if not isinstance(texto, str): return ""
        try:
            return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
        except TypeError: return ""

    def buscar_cantidad_para_magnitud(self, mag: str, descripcion: str) -> Optional[str]:
        """Busca una magnitud específica precedida por un número válido en la descripción."""
        if not isinstance(mag, str) or not mag: return None
        if not isinstance(descripcion, str) or len(descripcion) < 1: return None

        mag_upper = mag.upper()
        texto_limpio = self._quitar_diacronicos_y_acentos(descripcion.upper())
        cantidad_encontrada: Optional[str] = None
        mag_escapada = re.escape(mag_upper)

        patron_completo = re.compile(
             r"([0-9]{1}[,.][\d]{1,3}|[1-9]{1}[\d]{0,1}[,.][\d]{0,2}|[1-9]{1}[\d]{0,2}[,.][\d]{1}|[1-9]{1}[\d]{0,3})"
             r"[ X]{0,1}"
             r"(" + mag_escapada + r")"
        )
        patron_simple_avance = re.compile(
             r"([0-9]+([,.][0-9]+)?[ X]{0,1})"
             r"(" + mag_escapada + r")"
             r"([\s\*\/\+\-\(\)\[\]\,\.\;\:]|$)"
        )

        posicion_actual = 0
        while posicion_actual < len(texto_limpio):
            match_completo = patron_completo.search(texto_limpio, posicion_actual)
            if match_completo:
                cantidad_encontrada = match_completo.group(1).strip()
                # print(f"DEBUG: Magnitud encontrada -> Cantidad: '{cantidad_encontrada}', Mag: '{mag_upper}', Match: '{match_completo.group(0)}'")
                return cantidad_encontrada
            else:
                match_simple = patron_simple_avance.search(texto_limpio, posicion_actual)
                if match_simple:
                    posicion_actual = match_simple.end()
                    # print(f"DEBUG: Match simple encontrado pero no válido, avanzando a {posicion_actual}. Match: {match_simple.group(0)}")
                else:
                    # print(f"DEBUG: No se encontró ni match simple a partir de {posicion_actual}. Terminando búsqueda para '{mag_upper}'.")
                    break
        return None
# >>> FIN: Definición de la Clase ExtractorMagnitud <<<


# --- Clase de la Interfaz Gráfica ---

class InterfazGrafica(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado")
        self.geometry("1000x800")
        self.motor = MotorBusqueda()
        # >>> INSTANCIACIÓN: Crear un objeto ExtractorMagnitud <<<
        # Creamos una instancia del extractor aquí para que esté disponible
        # en toda la clase InterfazGrafica a través de self.extractor_magnitud
        self.extractor_magnitud = ExtractorMagnitud()
        # <<< FIN INSTANCIACIÓN >>>
        self.resultados_actuales: Optional[pd.DataFrame] = None

        self.color_fila_par = "white"
        self.color_fila_impar = "#f0f0f0"

        # --- (Lógica de selección de tema TTK - Sin cambios) ---
        style = ttk.Style(self)
        available_themes = style.theme_names()
        current_os = platform.system()
        chosen_theme = None
        print(f"--- Theme Debug ---")
        print(f"Sistema Operativo: {current_os}")
        print(f"Temas TTK Disponibles: {available_themes}")
        if current_os == "Windows":
            if "vista" in available_themes: chosen_theme = "vista"
            elif "xpnative" in available_themes: chosen_theme = "xpnative"
        elif current_os == "Darwin":
            if "aqua" in available_themes: chosen_theme = "aqua"
        if not chosen_theme:
            if "clam" in available_themes: chosen_theme = "clam"
            else:
                try: chosen_theme = style.theme_use()
                except tk.TclError: chosen_theme = "default"
        if chosen_theme:
            print(f"Tema Seleccionado: {chosen_theme}")
            try: style.theme_use(chosen_theme)
            except tk.TclError as e: print(f"Advertencia: No se pudo aplicar el tema '{chosen_theme}'. Error: {e}")
        else: print("Advertencia: No se pudo determinar un tema TTK adecuado.")
        print(f"-------------------")
        # --- FIN Lógica de selección de tema TTK ---

        self._crear_widgets()
        self._configurar_grid()
        self._configurar_eventos()
        self._configurar_tags_treeview()
        self._actualizar_estado("Listo. Cargue el Diccionario y las Descripciones.")

    # >>> (Métodos _crear_widgets, _configurar_tags_treeview, _configurar_grid,
    #      _configurar_eventos, _actualizar_estado, _actualizar_tabla,
    #      _cargar_diccionario, _cargar_excel_descripcion - Sin cambios) <<<
    def _crear_widgets(self):
        self.marco_controles = ttk.LabelFrame(self, text="Controles")
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion, state="disabled")
        self.lbl_busqueda = ttk.Label(self.marco_controles, text="Término de Búsqueda (use '+' para AND, '-' para OR):")
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50)
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda, state="disabled")
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar Resultados", command=self._exportar_resultados, state="disabled")
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario (Columnas 0 y 3):")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados de Búsqueda / Descripciones Cargadas:")
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings")
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings")
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)
        self.barra_estado = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W)

    def _configurar_tags_treeview(self):
        self.tabla_diccionario.tag_configure('par', background=self.color_fila_par)
        self.tabla_diccionario.tag_configure('impar', background=self.color_fila_impar)
        self.tabla_resultados.tag_configure('par', background=self.color_fila_par)
        self.tabla_resultados.tag_configure('impar', background=self.color_fila_impar)

    def _configurar_grid(self):
        self.grid_rowconfigure(2, weight=1); self.grid_rowconfigure(4, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        self.marco_controles.grid_columnconfigure(2, weight=1)
        self.marco_controles.grid_columnconfigure(4, weight=0)
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.btn_cargar_descripciones.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.lbl_busqueda.grid(row=1, column=0, columnspan=2, padx=5, pady=(5,0), sticky="w")
        self.entrada_busqueda.grid(row=1, column=2, padx=5, pady=(5,5), sticky="ew")
        self.btn_buscar.grid(row=1, column=3, padx=5, pady=(5,5))
        self.btn_exportar.grid(row=0, column=4, rowspan=2, padx=(10, 5), pady=5, sticky="e")
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0))
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1); self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1); self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew"); self.scrolly_diccionario.grid(row=0, column=1, sticky="ns"); self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew"); self.scrolly_resultados.grid(row=0, column=1, sticky="ns"); self.scrollx_resultados.grid(row=1, column=0, sticky="ew")
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0))

    def _configurar_eventos(self):
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())

    def _actualizar_estado(self, mensaje: str):
        self.barra_estado.config(text=mensaje)
        self.update_idletasks()

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None):
        try:
            for item in tabla.get_children(): tabla.delete(item)
            if tabla["columns"]: tabla["columns"] = ()
        except tk.TclError as e: print(f"Advertencia: Error Tcl al limpiar tabla: {e}")
        if datos is None or datos.empty:
            tabla["columns"] = ()
            return
        columnas = list(datos.columns)
        tabla["columns"] = columnas
        for col in columnas:
            tabla.heading(col, text=str(col), anchor=tk.W)
            try:
                sample_size = min(len(datos), 100)
                col_data_subset = datos[col].head(sample_size).dropna().astype(str)
                content_width = col_data_subset.str.len().max() if not col_data_subset.empty else 0
                header_width = len(str(col))
                width_factor = 7; padding = 20
                calculated_width = max(header_width * (width_factor + 1), content_width * width_factor) + padding
                final_width = max(50, min(calculated_width, 400))
                tabla.column(col, width=final_width, minwidth=50, anchor=tk.W)
            except Exception as e:
                print(f"Advertencia: Error calculando ancho para columna '{col}': {e}")
                tabla.column(col, width=len(str(col)) * 8 + 15, minwidth=50, anchor=tk.W)
        datos_a_mostrar = datos.head(limite_filas) if limite_filas is not None and len(datos) > limite_filas else datos
        for i, (_, fila) in enumerate(datos_a_mostrar.iterrows()):
            valores = [str(v) if pd.notna(v) else "" for v in fila.values]
            tag_actual = 'par' if i % 2 == 0 else 'impar'
            try: tabla.insert("", "end", values=valores, tags=(tag_actual,))
            except tk.TclError as e: print(f"Advertencia: Error Tcl al insertar fila {i}: {e}")

    def _cargar_diccionario(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Archivo de Diccionario", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not ruta: return
        self._actualizar_estado("Cargando diccionario...")
        if self.motor.cargar_excel_diccionario(ruta):
            df_dic_original = self.motor.datos_diccionario
            if df_dic_original is not None and not df_dic_original.empty:
                num_filas = len(df_dic_original)
                self._actualizar_estado(f"Procesando diccionario ({num_filas} filas)...")
                df_dic_display = pd.DataFrame()
                cols_a_mostrar_nombres = []
                if len(df_dic_original.columns) > 0: cols_a_mostrar_nombres.append(df_dic_original.columns[0])
                if len(df_dic_original.columns) > 3: cols_a_mostrar_nombres.append(df_dic_original.columns[3])
                if cols_a_mostrar_nombres:
                    cols_validas = [c for c in cols_a_mostrar_nombres if c in df_dic_original.columns]
                    if cols_validas: df_dic_display = df_dic_original[cols_validas].copy()
                    else: print("Advertencia: Columnas 0 y/o 3 no encontradas para mostrar en vista previa.")
                self._actualizar_tabla(self.tabla_diccionario, df_dic_display, limite_filas=100)
                self.title(f"Buscador - Diccionario: {ruta.split('/')[-1]}")
                self.btn_cargar_descripciones["state"] = "normal"
                if self.motor.datos_descripcion is not None: self.btn_buscar["state"] = "normal"
                mensaje_columnas = f"Columnas '{', '.join(cols_a_mostrar_nombres)}' mostradas." if cols_a_mostrar_nombres else "No se pudieron mostrar columnas 0 y 3."
                self._actualizar_estado(f"Diccionario '{ruta.split('/')[-1]}' cargado ({num_filas} filas). {mensaje_columnas} Cargue descripciones.")
                messagebox.showinfo("Éxito", f"Diccionario cargado ({num_filas} filas).\n{mensaje_columnas}\n(Vista previa muestra máx. 100 filas).")
            else:
                self._actualizar_tabla(self.tabla_diccionario, None)
                self._actualizar_estado("Diccionario cargado pero vacío o inválido.")
                self.btn_cargar_descripciones["state"] = "disabled"; self.btn_buscar["state"] = "disabled"
                messagebox.showwarning("Archivo Vacío", "El archivo de diccionario se cargó pero está vacío.")
        else:
            self._actualizar_tabla(self.tabla_diccionario, None)
            self._actualizar_estado("Error al cargar diccionario.")
            self.btn_cargar_descripciones["state"] = "disabled"; self.btn_buscar["state"] = "disabled"

    def _cargar_excel_descripcion(self):
        ruta_descripciones = filedialog.askopenfilename(title="Seleccionar archivo de descripciones", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not ruta_descripciones: return
        self._actualizar_estado("Cargando descripciones...")
        if self.motor.cargar_excel_descripcion(ruta_descripciones):
            df_desc = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc)
            self.resultados_actuales = df_desc.copy() if df_desc is not None else None
            num_filas = len(df_desc) if df_desc is not None else 0
            self.btn_exportar["state"] = "normal" if num_filas > 0 else "disabled"
            if self.motor.datos_diccionario is not None: self.btn_buscar["state"] = "normal"
            self._actualizar_estado(f"Descripciones '{ruta_descripciones.split('/')[-1]}' cargadas ({num_filas} filas). Listo para buscar.")
            messagebox.showinfo("Éxito", f"Descripciones cargadas ({num_filas} filas). Mostrando en la tabla de resultados.")
        else:
            self._actualizar_estado("Error al cargar descripciones.")
            self.btn_buscar["state"] = "disabled"; self.btn_exportar["state"] = "disabled"
            self._actualizar_tabla(self.tabla_resultados, None); self.resultados_actuales = None

    def _ejecutar_busqueda(self):
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", "Debe cargar tanto el archivo de Diccionario como el de Descripciones antes de buscar.")
            return

        termino = self.entrada_busqueda.get()

        if not termino.strip():
            messagebox.showinfo("Búsqueda Vacía", "Mostrando todas las descripciones cargadas.")
            df_desc = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc)
            self.resultados_actuales = df_desc.copy() if df_desc is not None else None
            num_filas = len(df_desc) if df_desc is not None else 0
            self.btn_exportar["state"] = "normal" if num_filas > 0 else "disabled"
            self._actualizar_estado(f"Mostrando todas las {num_filas} descripciones.")
            return

        self._actualizar_estado(f"Buscando '{termino}'...")
        resultados = self.motor.buscar(termino)

        self._actualizar_tabla(self.tabla_resultados, None)
        self.resultados_actuales = None
        self.btn_exportar["state"] = "disabled"

        if resultados is None:
            self._actualizar_estado(f"Búsqueda de '{termino}' no pudo iniciarse o falló gravemente.")

        elif isinstance(resultados, tuple):
            # ... (manejo de caso 'no encontrado en diccionario' o error específico - sin cambios) ...
            df_dic, df_desc = resultados[0], resultados[1]
            error_flag = resultados[2] if len(resultados) == 3 else None
            if error_flag:
                if error_flag == "ErrorColumnas": self._actualizar_estado("Error: Problema con columnas del diccionario.")
                elif error_flag == "ErrorTermino": self._actualizar_estado(f"Error: Término de búsqueda '{termino}' inválido o vacío tras procesar.")
                else: self._actualizar_estado(f"Error inesperado durante la búsqueda inicial: {error_flag}")
            else:
                respuesta = messagebox.askyesno("Sin Coincidencias en Diccionario",
                                                f"El término '{termino}' no se encontró en las columnas de búsqueda del diccionario.\n\n"
                                                "¿Desea buscar directamente en todas las columnas de las descripciones?")
                if respuesta:
                    self._actualizar_estado(f"Buscando '{termino}' directamente en descripciones...")
                    resultados_directos = self.motor.buscar_en_descripciones_directo(termino)
                    self._actualizar_tabla(self.tabla_resultados, resultados_directos)
                    self.resultados_actuales = resultados_directos
                    num_res = len(resultados_directos) if resultados_directos is not None else 0
                    if num_res > 0:
                        self.btn_exportar["state"] = "normal"
                        self._actualizar_estado(f"Búsqueda directa de '{termino}' completada: {num_res} resultados.")
                        # >>> INICIO: Demostración de uso del ExtractorMagnitud <<<
                        # Esto es solo un ejemplo para mostrar cómo llamar al extractor
                        # con el texto de la primera celda del primer resultado.
                        if num_res > 0 and len(resultados_directos.columns) > 0:
                            try:
                                texto_primer_resultado = str(resultados_directos.iloc[0, 0]) # Texto de fila 0, columna 0
                                print("\n--- DEMO Extractor Magnitudes (1er resultado) ---")
                                for mag_demo in ["V", "A", "KM", "GB", "PORT"]: # Probar algunas magnitudes
                                    cantidad_extraida = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag_demo, texto_primer_resultado)
                                    if cantidad_extraida is not None:
                                        print(f"  -> Magnitud '{mag_demo}' encontrada: Cantidad = {cantidad_extraida}")
                                    else:
                                        print(f"  -> Magnitud '{mag_demo}' NO encontrada en texto: '{texto_primer_resultado[:50]}...'")
                                print("--- FIN DEMO ---")
                            except Exception as e_demo:
                                print(f"Error en la demostración del extractor: {e_demo}")
                        # <<< FIN: Demostración de uso del ExtractorMagnitud <<<
                    else:
                        messagebox.showinfo("Sin Coincidencias", f"La búsqueda directa de '{termino}' tampoco encontró resultados.")
                        self._actualizar_estado(f"Búsqueda directa de '{termino}' completada: 0 resultados.")
                else:
                    self._actualizar_estado(f"Búsqueda de '{termino}' cancelada (sin coincidencias en diccionario).")

        elif isinstance(resultados, pd.DataFrame):
            self.resultados_actuales = resultados
            num_res = len(resultados)
            self._actualizar_tabla(self.tabla_resultados, resultados)
            if num_res > 0:
                self.btn_exportar["state"] = "normal"
                self._actualizar_estado(f"Búsqueda de '{termino}' completada: {num_res} resultados encontrados.")
                # >>> INICIO: Demostración de uso del ExtractorMagnitud <<<
                # Repetimos la demostración aquí para el caso de búsqueda normal exitosa
                if len(resultados.columns) > 0:
                    try:
                        texto_primer_resultado = str(resultados.iloc[0, 0]) # Texto de fila 0, columna 0
                        print("\n--- DEMO Extractor Magnitudes (1er resultado) ---")
                        for mag_demo in ["V", "A", "KM", "GB", "PORT"]: # Probar algunas magnitudes
                             # Usamos self.extractor_magnitud que creamos en __init__
                            cantidad_extraida = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag_demo, texto_primer_resultado)
                            if cantidad_extraida is not None:
                                print(f"  -> Magnitud '{mag_demo}' encontrada: Cantidad = {cantidad_extraida}")
                            else:
                                print(f"  -> Magnitud '{mag_demo}' NO encontrada en texto: '{texto_primer_resultado[:50]}...'")
                        print("--- FIN DEMO ---")
                    except Exception as e_demo:
                        print(f"Error en la demostración del extractor: {e_demo}")
                # <<< FIN: Demostración de uso del ExtractorMagnitud <<<
            else:
                messagebox.showinfo("Sin Coincidencias Finales",
                                    f"Se encontraron términos en el diccionario para '{termino}', pero no se hallaron coincidencias finales en las descripciones.")
                self._actualizar_estado(f"Búsqueda de '{termino}' completada: 0 resultados en descripciones.")
        else:
            self._actualizar_estado(f"Error: Tipo de resultado inesperado ({type(resultados)}) tras buscar '{termino}'.")
            print(f"Resultado inesperado de búsqueda: {resultados}")

    def _exportar_resultados(self):
        # >>> (Sin cambios en _exportar_resultados) <<<
        if self.resultados_actuales is None or self.resultados_actuales.empty:
            messagebox.showwarning("Sin Resultados", "No hay resultados para exportar.")
            return
        file_types = [("Archivo Excel", "*.xlsx"), ("Archivo CSV (UTF-8)", "*.csv"), ("Excel 97-2003", "*.xls")]
        ruta = filedialog.asksaveasfilename(title="Guardar Resultados Como", defaultextension=".xlsx", filetypes=file_types)
        if not ruta: return
        self._actualizar_estado("Exportando resultados...")
        try:
            extension = ruta.split('.')[-1].lower()
            if extension == 'csv': self.resultados_actuales.to_csv(ruta, index=False, encoding='utf-8-sig')
            elif extension == 'xlsx': self.resultados_actuales.to_excel(ruta, index=False, engine='openpyxl')
            elif extension == 'xls':
                try:
                    import xlwt
                    if len(self.resultados_actuales) > 65535:
                        messagebox.showwarning("Límite de Filas Excedido", f"El formato .xls solo soporta hasta 65,536 filas.\nSus resultados ({len(self.resultados_actuales)} filas) serán truncados.")
                        df_to_export = self.resultados_actuales.head(65535)
                    else: df_to_export = self.resultados_actuales
                    df_to_export.to_excel(ruta, index=False, engine='xlwt')
                except ImportError:
                    messagebox.showerror("Librería Faltante", "Para exportar a formato .xls, necesita instalar 'xlwt':\n`pip install xlwt`")
                    self._actualizar_estado("Error al exportar: Falta 'xlwt'."); return
                except Exception as ex_xls: raise ex_xls
            else:
                messagebox.showerror("Extensión Inválida", f"Extensión de archivo no soportada: {extension}")
                self._actualizar_estado("Error al exportar: Extensión inválida."); return
            messagebox.showinfo("Éxito", f"Resultados ({len(self.resultados_actuales)} filas) exportados correctamente a:\n{ruta}")
            self._actualizar_estado(f"Resultados exportados a {ruta.split('/')[-1]}")
        except ImportError as imp_err:
            if 'openpyxl' in str(imp_err) and ruta.endswith('.xlsx'):
                messagebox.showerror("Librería Faltante", "Para exportar a formato .xlsx, necesita instalar 'openpyxl':\n`pip install openpyxl`")
                self._actualizar_estado("Error al exportar: Falta 'openpyxl'.")
            else:
                messagebox.showerror("Error de Importación", f"Falta una librería necesaria para exportar:\n{imp_err}")
                self._actualizar_estado("Error al exportar: Librería faltante."); print(traceback.format_exc())
        except Exception as e:
            error_detallado = traceback.format_exc()
            print(f"Error detallado de exportación:\n{error_detallado}")
            messagebox.showerror("Error de Exportación", f"No se pudo guardar el archivo:\n{e}")
            self._actualizar_estado("Error al exportar resultados.")


if __name__ == "__main__":
    # >>> (Sin cambios en el bloque __main__) <<<
    missing_libs = []
    try: import pandas as pd
    except ImportError: missing_libs.append("pandas")
    try: import openpyxl
    except ImportError: missing_libs.append("openpyxl (necesario para .xlsx)")
    # >>> Añadir unicodedata a la comprobación inicial podría ser buena idea <<<
    # try: import unicodedata
    # except ImportError: missing_libs.append("unicodedata (parte de la biblioteca estándar, debería estar)")

    if missing_libs:
        root_error = tk.Tk()
        root_error.withdraw()
        messagebox.showerror("Dependencias Faltantes",
                             f"Error: Faltan librerías necesarias:\n - {chr(10)} - ".join(missing_libs) +
                             f"\n\nPor favor, instálalas (ej: pip install pandas openpyxl) y reinicia la aplicación.")
        root_error.destroy()
        exit(1)

    app = InterfazGrafica()
    app.mainloop()
