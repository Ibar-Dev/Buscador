# -*- coding: utf-8 -*- # Buena práctica añadir codificación
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Tuple, Union, Set, Callable
import traceback
import platform
# >>> INICIO: Import necesario para ExtractorMagnitud <<<
# unicodedata es parte de la biblioteca estándar de Python, así que normalmente no necesita instalación extra.
import unicodedata
# <<< FIN: Import necesario para ExtractorMagnitud <<<

# Ya no necesitamos esta lista global aquí, la definimos dentro de la clase ExtractorMagnitud
# magnitudes_predefinidas: List[str] = [ ... ]


# --- Clases de Lógica ---

class ManejadorExcel:
    """Clase estática para manejar la carga de archivos Excel."""
    @staticmethod
    def cargar_excel(ruta: str) -> Optional[pd.DataFrame]:
        """
        Carga un archivo Excel (.xlsx o .xls) en un DataFrame de pandas.

        Args:
            ruta: La ruta completa al archivo Excel.

        Returns:
            Un DataFrame de pandas con los datos o None si ocurre un error.
        """
        try:
            # Determina el motor adecuado según la extensión
            engine = 'openpyxl' if ruta.endswith('.xlsx') else None
            df = pd.read_excel(ruta, engine=engine)
            print(f"Archivo '{ruta.split('/')[-1]}' cargado con éxito ({len(df)} filas).") # Info útil en consola
            return df
        except FileNotFoundError:
            messagebox.showerror("Error al Cargar Archivo", f"El archivo no se encontró en la ruta:\n{ruta}")
            return None
        except Exception as e:
            messagebox.showerror("Error al Cargar Archivo",
                                 f"No se pudo cargar el archivo:\n{ruta}\n\n"
                                 f"Error: {e}\n\n"
                                 "Asegúrese de que el archivo no esté corrupto, cerrado y "
                                 "tenga instalado 'openpyxl' (`pip install openpyxl`) para archivos .xlsx.")
            traceback.print_exc() # Imprime el error completo en consola para debugging
            return None

class MotorBusqueda:
    """Gestiona los datos y la lógica de búsqueda entre diccionario y descripciones."""
    def __init__(self):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[str] = None
        self.archivo_descripcion_actual: Optional[str] = None
        # Columnas que se usarán para la búsqueda en el diccionario (ej. 0 y 3)
        self.indices_columnas_busqueda_dic: List[int] = [0, 3]

    def cargar_excel_diccionario(self, ruta: str) -> bool:
        """Carga el archivo Excel que actúa como diccionario."""
        self.datos_diccionario = ManejadorExcel.cargar_excel(ruta)
        self.archivo_diccionario_actual = ruta if self.datos_diccionario is not None else None
        # Validar columnas después de cargar
        if self.datos_diccionario is not None and not self._validar_columnas_diccionario():
             self.datos_diccionario = None # Invalidar si las columnas no son adecuadas
             self.archivo_diccionario_actual = None
             return False
        return self.datos_diccionario is not None

    def cargar_excel_descripcion(self, ruta: str) -> bool:
        """Carga el archivo Excel que contiene las descripciones donde buscar."""
        self.datos_descripcion = ManejadorExcel.cargar_excel(ruta)
        self.archivo_descripcion_actual = ruta if self.datos_descripcion is not None else None
        return self.datos_descripcion is not None

    def _validar_columnas_diccionario(self) -> bool:
        """Verifica si las columnas esperadas existen en el diccionario."""
        if self.datos_diccionario is None: return False
        num_cols = len(self.datos_diccionario.columns)
        max_indice_requerido = max(self.indices_columnas_busqueda_dic) if self.indices_columnas_busqueda_dic else -1

        if num_cols == 0:
            messagebox.showerror("Error de Diccionario", "El archivo de diccionario está vacío o no tiene columnas.")
            return False
        elif num_cols <= max_indice_requerido:
            messagebox.showerror("Error de Diccionario",
                                 f"El diccionario necesita al menos {max_indice_requerido + 1} columnas "
                                 f"para buscar en los índices {self.indices_columnas_busqueda_dic}, pero solo tiene {num_cols}.")
            return False
        # Opcional: Advertir si solo se usará la primera columna porque no hay suficientes
        elif len(self.indices_columnas_busqueda_dic) > 1 and num_cols <= self.indices_columnas_busqueda_dic[1]:
             messagebox.showwarning("Advertencia de Diccionario",
                                    f"El diccionario tiene {num_cols} columnas. Se esperaba buscar en la columna índice "
                                    f"{self.indices_columnas_busqueda_dic[1]} pero no existe. Se buscará solo en la columna índice "
                                    f"{self.indices_columnas_busqueda_dic[0]}.")
        return True

    def _obtener_nombres_columnas_busqueda(self, df: pd.DataFrame) -> Optional[List[str]]:
        """Obtiene los nombres de las columnas de búsqueda basados en los índices."""
        if df is None:
            messagebox.showerror("Error Interno", "Se intentó obtener columnas de un DataFrame nulo (diccionario).")
            return None
        if not self._validar_columnas_diccionario(): # Re-validar por si acaso
            return None

        columnas_disponibles = df.columns
        cols_encontradas_nombres = []
        num_cols_df = len(columnas_disponibles)

        for indice in self.indices_columnas_busqueda_dic:
            if indice < num_cols_df:
                cols_encontradas_nombres.append(columnas_disponibles[indice])
            # Si un índice posterior no existe, podemos decidir parar o continuar con los que sí existen.
            # Por ahora, solo añadimos las que existen dentro de los índices definidos.

        if not cols_encontradas_nombres:
             messagebox.showerror("Error de Diccionario", f"No se pudieron obtener nombres de columnas para los índices: {self.indices_columnas_busqueda_dic}")
             return None

        return cols_encontradas_nombres

    def _extraer_terminos_diccionario(self, df_coincidencias: pd.DataFrame, columnas_busqueda_nombres: List[str]) -> Set[str]:
        """Extrae términos únicos de las columnas especificadas en las filas coincidentes."""
        terminos_encontrados: Set[str] = set()
        if df_coincidencias is None or df_coincidencias.empty or not columnas_busqueda_nombres:
            return terminos_encontrados

        # Asegurarse de que las columnas realmente existen en este sub-dataframe
        columnas_validas = [col for col in columnas_busqueda_nombres if col in df_coincidencias.columns]
        if not columnas_validas:
            print(f"Advertencia: Ninguna de las columnas de búsqueda {columnas_busqueda_nombres} se encontró en las coincidencias del diccionario.")
            return terminos_encontrados

        for col in columnas_validas:
            try:
                # 1. Selecciona la columna, quita NaNs, convierte a string, pasa a mayúsculas, obtiene únicos
                terminos_col = df_coincidencias[col].dropna().astype(str).str.upper().unique()
                # 2. Añade los términos encontrados al conjunto general
                terminos_encontrados.update(terminos_col)
            except Exception as e:
                print(f"Advertencia: Error extrayendo términos de la columna '{col}': {e}. Se intentará conversión simple.")
                # Intento más simple si falla la cadena de métodos (poco probable con .astype(str))
                try:
                    terminos_col = set(df_coincidencias[col].astype(str).str.upper().unique())
                    terminos_col.discard('NAN') # Pandas a veces convierte NaN a 'nan' o 'NAN' en string
                    terminos_encontrados.update(terminos_col)
                except Exception as e_fallback:
                    print(f"Error crítico en fallback extrayendo términos de columna '{col}': {e_fallback}")

        # Filtrar términos vacíos o que solo sean espacios
        terminos_encontrados = {t for t in terminos_encontrados if t and not t.isspace()}
        return terminos_encontrados

    def _buscar_terminos_en_descripciones(self,
                                         df_descripcion: pd.DataFrame,
                                         terminos_a_buscar: Set[str],
                                         require_all: bool = False) -> pd.DataFrame:
        """Filtra el DataFrame de descripciones buscando filas que contengan los términos."""
        columnas_originales = list(df_descripcion.columns) if df_descripcion is not None else []

        if df_descripcion is None or df_descripcion.empty or not terminos_a_buscar:
            return pd.DataFrame(columns=columnas_originales) # Devuelve DF vacío con columnas originales

        # Función agregadora: all para AND, any para OR (para los *términos encontrados en diccionario*)
        agg_func: Callable = all if require_all else any

        try:
            # 1. Crear una serie donde cada elemento es el texto concatenado de toda la fila en mayúsculas
            #    Usamos 'fillna' para manejar NaNs antes de convertir a str y unir
            texto_filas = df_descripcion.fillna('').astype(str).agg(' '.join, axis=1).str.upper()

            # 2. Asegurarse de que los términos a buscar son válidos (no vacíos)
            terminos_validos = {t for t in terminos_a_buscar if t}
            if not terminos_validos:
                print("Advertencia: No hay términos válidos (extraídos del diccionario) para buscar en descripciones.")
                return pd.DataFrame(columns=columnas_originales)

            # 3. Crear la máscara: para cada fila, comprobar si contiene los términos
            #    Usamos expresión regular con límites de palabra (\b) para buscar términos completos
            #    y re.escape para manejar caracteres especiales en los términos.
            terminos_escapados = [r"\b" + re.escape(t) + r"\b" for t in terminos_validos]
            patron_terminos = '|'.join(terminos_escapados) # Unir con OR (|) para buscar cualquiera

            if require_all: # Si todos los términos deben estar presentes
                 mascara_descripcion = texto_filas.apply(lambda texto: all(re.search(r"\b" + re.escape(t) + r"\b", texto) for t in terminos_validos))
            else: # Si al menos uno debe estar presente
                 mascara_descripcion = texto_filas.str.contains(patron_terminos, regex=True, na=False)


            # 4. Aplicar la máscara para obtener las filas resultado
            resultados = df_descripcion[mascara_descripcion]

        except Exception as e:
            messagebox.showerror("Error en Búsqueda de Descripciones", f"Ocurrió un error al filtrar descripciones:\n{e}")
            traceback.print_exc()
            return pd.DataFrame(columns=columnas_originales) # Devuelve DF vacío con columnas originales

        return resultados

    def buscar(self, termino_buscado: str) -> Union[None, pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        """
        Método principal de búsqueda. Orquesta la búsqueda en diccionario y luego en descripciones.

        Args:
            termino_buscado: El texto introducido por el usuario.

        Returns:
            - pd.DataFrame: Si la búsqueda es exitosa, devuelve el DataFrame de descripciones filtrado.
            - Tuple[pd.DataFrame, pd.DataFrame]: Si el término no se encuentra en el diccionario,
              devuelve los DataFrames originales para la opción de búsqueda directa.
            - None: Si ocurre un error grave o faltan datos.
        """
        if self.datos_diccionario is None:
            messagebox.showwarning("Diccionario No Cargado", "Por favor, cargue primero el archivo 'Diccionario'.")
            return None
        if self.datos_descripcion is None:
            messagebox.showwarning("Descripciones No Cargadas", "Por favor, cargue primero el archivo 'Descripciones'.")
            return None

        termino_limpio = termino_buscado.strip().upper()

        # Si no hay término de búsqueda, devolvemos todas las descripciones
        if not termino_limpio:
            return self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame()

        # Hacemos copias para no modificar los originales
        df_diccionario = self.datos_diccionario.copy()
        df_descripcion = self.datos_descripcion.copy()

        if df_diccionario.empty:
            messagebox.showerror("Error", "El DataFrame del diccionario está vacío.")
            return None # O quizás devolver el tuple para búsqueda directa? Depende del flujo deseado.
        if df_descripcion.empty:
            messagebox.showerror("Error", "El DataFrame de descripciones está vacío.")
            return pd.DataFrame(columns=df_descripcion.columns) # Devolver DF vacío si no hay dónde buscar

        try:
            # Determinar el tipo de búsqueda y delegar
            if '+' in termino_limpio:
                # Búsqueda AND: todas las palabras deben estar en el diccionario
                return self._busqueda_compuesta(df_diccionario, df_descripcion, termino_limpio, '+', 'AND', require_all_desc=False)
            elif '-' in termino_limpio:
                # Búsqueda OR: al menos una palabra debe estar en el diccionario
                return self._busqueda_compuesta(df_diccionario, df_descripcion, termino_limpio, '-', 'OR', require_all_desc=False)
            else:
                # Búsqueda Simple: la palabra/frase debe estar en el diccionario
                return self._busqueda_simple(df_diccionario, df_descripcion, termino_limpio)

        except Exception as e:
            messagebox.showerror("Error Inesperado en Búsqueda", f"Ocurrió un error durante la búsqueda:\n{e}")
            traceback.print_exc()
            # Devolver los originales podría ser una opción, o None
            return None

    def _aplicar_mascara_diccionario(self, df: pd.DataFrame, columnas_nombres: List[str], palabras: List[str], operador: str) -> pd.Series:
        """Aplica la lógica de búsqueda (SIMPLE, AND, OR) sobre las columnas del diccionario."""
        if df is None or df.empty or not columnas_nombres or not palabras:
            return pd.Series(False, index=df.index if df is not None else None)

        # Validar que las columnas existan en el DataFrame actual
        columnas_validas = [col for col in columnas_nombres if col in df.columns]
        if not columnas_validas:
            print(f"Advertencia: Ninguna de las columnas {columnas_nombres} existe en el DataFrame del diccionario para aplicar máscara.")
            return pd.Series(False, index=df.index)

        # Convertir columnas relevantes a string para búsqueda segura
        try:
            df_str = df[columnas_validas].fillna('').astype(str)
        except Exception as e:
            print(f"Error convirtiendo columnas {columnas_validas} a string en _aplicar_mascara_diccionario: {e}")
            return pd.Series(False, index=df.index)

        # Escapar términos para regex y añadir límites de palabra
        palabras_escapadas = [r"\b" + re.escape(p) + r"\b" for p in palabras]

        mascara_total = pd.Series(False, index=df.index) # Iniciar en Falso por defecto

        if operador == 'SIMPLE':
            # Buscar la palabra (como regex) en CUALQUIERA de las columnas válidas
            patron = palabras_escapadas[0]
            for col in columnas_validas:
                try:
                    mascara_total |= df_str[col].str.contains(patron, regex=True, na=False, case=False)
                except Exception as e:
                    print(f"Error aplicando regex SIMPLE '{patron}' en columna '{col}': {e}")

        elif operador == 'OR':
            # Para cada palabra, buscarla en cualquier columna válida. Luego unir con OR.
            for patron in palabras_escapadas:
                mascara_palabra_actual = pd.Series(False, index=df.index)
                for col in columnas_validas:
                    try:
                        mascara_palabra_actual |= df_str[col].str.contains(patron, regex=True, na=False, case=False)
                    except Exception as e:
                        print(f"Error aplicando regex OR '{patron}' en columna '{col}': {e}")
                mascara_total |= mascara_palabra_actual # Acumular con OR

        elif operador == 'AND':
            # TODAS las palabras deben encontrarse (cada una en al menos una columna válida)
            mascara_total = pd.Series(True, index=df.index) # Empezar en True para AND
            for patron in palabras_escapadas:
                mascara_palabra_actual = pd.Series(False, index=df.index) # Dónde se encuentra esta palabra
                for col in columnas_validas:
                    try:
                        mascara_palabra_actual |= df_str[col].str.contains(patron, regex=True, na=False, case=False)
                    except Exception as e:
                        print(f"Error aplicando regex AND '{patron}' en columna '{col}': {e}")
                mascara_total &= mascara_palabra_actual # Acumular con AND

        return mascara_total

    def _busqueda_simple(self, df_diccionario: pd.DataFrame, df_descripcion: pd.DataFrame, termino: str) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        """Realiza una búsqueda simple de un término."""
        columnas_busqueda_nombres = self._obtener_nombres_columnas_busqueda(df_diccionario)
        if columnas_busqueda_nombres is None:
            # Si no se pueden obtener las columnas, podría ser un error o simplemente no hay
            # Devolvemos el tuple para permitir búsqueda directa
            return (df_diccionario, df_descripcion)

        palabras = [termino.strip()] # Solo una palabra/frase en búsqueda simple
        if not palabras or not palabras[0]:
             return (df_diccionario, df_descripcion) # Término vacío

        mascara_diccionario = self._aplicar_mascara_diccionario(df_diccionario, columnas_busqueda_nombres, palabras, 'SIMPLE')

        if not mascara_diccionario.any():
            # No encontrado en diccionario -> devolver tuple para opción directa
            return (df_diccionario, df_descripcion)

        # Encontrado en diccionario, extraer términos y buscar en descripciones
        coincidencias_diccionario = df_diccionario[mascara_diccionario]
        terminos_para_descripcion = self._extraer_terminos_diccionario(coincidencias_diccionario, columnas_busqueda_nombres)

        if not terminos_para_descripcion:
            messagebox.showinfo("Aviso", f"Se encontraron {len(coincidencias_diccionario)} fila(s) en el diccionario para '{termino}', pero no se pudieron extraer términos válidos de las columnas de búsqueda (podrían estar vacías o ser NaN). No se buscará en descripciones.")
            # Devolver DF vacío de descripciones
            return pd.DataFrame(columns=df_descripcion.columns)

        # Buscar los términos extraídos en las descripciones
        resultados_descripcion = self._buscar_terminos_en_descripciones(df_descripcion, terminos_para_descripcion, require_all=False) # require_all=False para búsqueda simple/OR
        return resultados_descripcion

    def _busqueda_compuesta(self, df_diccionario: pd.DataFrame, df_descripcion: pd.DataFrame, termino: str, separador: str, operador: str, require_all_desc: bool) -> Union[pd.DataFrame, Tuple[pd.DataFrame, pd.DataFrame]]:
        """Realiza una búsqueda compuesta (AND/OR) de múltiples términos."""
        columnas_busqueda_nombres = self._obtener_nombres_columnas_busqueda(df_diccionario)
        if columnas_busqueda_nombres is None:
            return (df_diccionario, df_descripcion)

        palabras = [p.strip() for p in termino.split(separador) if p.strip()]
        if not palabras:
            messagebox.showwarning("Término Inválido", f"La búsqueda '{termino}' no contiene términos válidos separados por '{separador}'.")
            return (df_diccionario, df_descripcion) # Término inválido

        mascara_diccionario = self._aplicar_mascara_diccionario(df_diccionario, columnas_busqueda_nombres, palabras, operador)

        if not mascara_diccionario.any():
             # No encontrado en diccionario -> devolver tuple para opción directa
            return (df_diccionario, df_descripcion)

        # Encontrado en diccionario, extraer términos y buscar en descripciones
        coincidencias_diccionario = df_diccionario[mascara_diccionario]
        terminos_para_descripcion = self._extraer_terminos_diccionario(coincidencias_diccionario, columnas_busqueda_nombres)

        if not terminos_para_descripcion:
            messagebox.showinfo("Aviso", f"Se encontraron {len(coincidencias_diccionario)} fila(s) en el diccionario para '{termino}', pero no se pudieron extraer términos válidos de las columnas de búsqueda. No se buscará en descripciones.")
            return pd.DataFrame(columns=df_descripcion.columns)

        # Buscar los términos extraídos en las descripciones
        # Para require_all_desc: normalmente False, pero podría ser True si quisiéramos que *todos* los términos extraídos aparezcan en la descripción
        resultados_descripcion = self._buscar_terminos_en_descripciones(df_descripcion, terminos_para_descripcion, require_all=require_all_desc)
        return resultados_descripcion


    def buscar_en_descripciones_directo(self, termino_buscado: str) -> pd.DataFrame:
        """Busca el término directamente en todas las columnas del DataFrame de descripciones."""
        if self.datos_descripcion is None or self.datos_descripcion.empty:
            messagebox.showwarning("Descripciones No Cargadas", "No hay datos de descripciones cargados para buscar directamente.")
            return pd.DataFrame()

        termino_limpio = termino_buscado.strip().upper()
        if not termino_limpio:
            return self.datos_descripcion.copy() # Devuelve todo si la búsqueda directa es vacía

        df_descripcion = self.datos_descripcion.copy()
        resultados = pd.DataFrame(columns=df_descripcion.columns) # Empezar con DF vacío

        try:
            # Crear el texto unificado por fila
            texto_filas = df_descripcion.fillna('').astype(str).agg(' '.join, axis=1).str.upper()
            mascara = pd.Series(False, index=df_descripcion.index)

            # Aplicar lógica AND/OR/SIMPLE directamente
            if '+' in termino_limpio:
                palabras = [p.strip() for p in termino_limpio.split('+') if p.strip()]
                if not palabras: return resultados # Búsqueda vacía
                mascara = pd.Series(True, index=df_descripcion.index) # Empezar en True para AND
                for palabra in palabras:
                    palabra_regex = r"\b" + re.escape(palabra) + r"\b"
                    mascara &= texto_filas.str.contains(palabra_regex, regex=True, na=False)
            elif '-' in termino_limpio:
                palabras = [p.strip() for p in termino_limpio.split('-') if p.strip()]
                if not palabras: return resultados # Búsqueda vacía
                for palabra in palabras:
                    palabra_regex = r"\b" + re.escape(palabra) + r"\b"
                    mascara |= texto_filas.str.contains(palabra_regex, regex=True, na=False) # Acumular con OR
            else: # Búsqueda simple directa
                palabra_regex = r"\b" + re.escape(termino_limpio) + r"\b"
                mascara = texto_filas.str.contains(palabra_regex, regex=True, na=False)

            resultados = df_descripcion[mascara]

        except Exception as e:
            messagebox.showerror("Error en Búsqueda Directa", f"Ocurrió un error al buscar directamente en descripciones:\n{e}")
            traceback.print_exc()
            return pd.DataFrame(columns=df_descripcion.columns) # Devolver DF vacío en caso de error

        return resultados


# >>> INICIO: Clase ExtractorMagnitud MEJORADA <<<
class ExtractorMagnitud:
    """
    Clase para buscar y extraer cantidades numéricas asociadas a magnitudes
    específicas dentro de un texto descriptivo.
    """
    # Es buena práctica tener las constantes dentro de la clase si le pertenecen
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
        # Usamos la lista de la clase (self.MAGNITUDES_PREDEFINIDAS) o la que nos pasen
        self.magnitudes = magnitudes if magnitudes is not None else self.MAGNITUDES_PREDEFINIDAS

    @staticmethod
    def _quitar_diacronicos_y_acentos(texto: str) -> str:
        """Elimina acentos y marcas diacríticas del texto."""
        if not isinstance(texto, str): return ""
        if not texto: return "" # Más directo para string vacío
        try:
            # NFKD suele ser más compatible para descomponer caracteres
            forma_normalizada = unicodedata.normalize('NFKD', texto)
            # Quitamos los caracteres combinados (acentos, etc.)
            return ''.join(c for c in forma_normalizada if not unicodedata.combining(c))
        except TypeError:
            return ""

    def buscar_cantidad_para_magnitud(self, mag: str, descripcion: str) -> Optional[str]:
        """
        Busca una magnitud específica precedida por un número válido en la descripción.
        Devuelve la primera cantidad encontrada que coincida con el patrón.
        """
        if not isinstance(mag, str) or not mag: return None
        if not isinstance(descripcion, str) or not descripcion: return None

        mag_upper = mag.upper()
        texto_limpio = self._quitar_diacronicos_y_acentos(descripcion.upper())
        if not texto_limpio: return None

        mag_escapada = re.escape(mag_upper)

        # Patrón Principal Mejorado:
        # 1. (\d+([.,]\d+)?) : Captura el número (entero o decimal con . o ,) - Grupo 1
        # 2. [ X]{0,1}       : Permite un espacio o 'X' opcional entre número y magnitud
        # 3. (\b{mag_escapada}\b) : Captura la magnitud EXACTA (como palabra completa) - Grupo 3
        # 4. (?![a-zA-Z0-9]) : Asegura que después de la magnitud no siga inmediatamente
        #                      otra letra o número (evita falsos positivos como '12VDC' para 'V')
        patron_principal = re.compile(
            r"(\d+([.,]\d+)?)"  # Grupo 1: El número
            r"[ X]{0,1}"
            r"(\b" + mag_escapada + r"\b)" # Grupo 3: La magnitud exacta
            r"(?![a-zA-Z0-9])" # No seguido por letra/número
        )

        # Usamos finditer para encontrar TODAS las coincidencias no solapadas
        # Es más eficiente que buscar en un bucle while con search y slicing
        for match in patron_principal.finditer(texto_limpio):
            cantidad = match.group(1) # El número completo (grupo 1)
            # No necesitamos verificar la magnitud (grupo 3) porque \b ya lo hace,
            # pero podríamos si quisiéramos ser extra seguros.

            # Devolvemos la *primera* cantidad válida que encontremos.
            # Podríamos añadir lógica si quisiéramos la última, la más grande, etc.
            # print(f"DEBUG: Encontrado -> Cantidad: '{cantidad}', Mag: '{mag_upper}', Contexto: '{match.group(0)}'")
            return cantidad.strip() # Quitamos posibles espacios extra

        # Si el bucle termina sin encontrar nada
        # print(f"DEBUG: No se encontró patrón para '{mag_upper}' en '{texto_limpio[:50]}...'")
        return None
# >>> FIN: Clase ExtractorMagnitud MEJORADA <<<


# --- Clase de la Interfaz Gráfica ---

class InterfazGrafica(tk.Tk):
    """Clase principal para la interfaz gráfica del buscador."""
    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado")
        # Ajustar tamaño inicial si es necesario
        self.geometry("1100x800") # Un poco más ancho quizás

        self.motor = MotorBusqueda()
        # >>> INSTANCIACIÓN: Crear un objeto ExtractorMagnitud <<<
        # Creamos una instancia del extractor aquí para usarla luego
        self.extractor_magnitud = ExtractorMagnitud()
        # <<< FIN INSTANCIACIÓN >>>
        self.resultados_actuales: Optional[pd.DataFrame] = None # Guarda el DF de resultados mostrado

        # Colores para filas alternas en las tablas
        self.color_fila_par = "white"
        self.color_fila_impar = "#f0f0f0" # Gris muy claro

        # --- Configuración del estilo TTK ---
        self._configurar_estilo_ttk()

        # --- Creación y configuración de widgets ---
        self._crear_widgets()
        self._configurar_grid()
        self._configurar_eventos()
        self._configurar_tags_treeview()
        self._actualizar_estado("Listo. Cargue el Diccionario y las Descripciones.")

    def _configurar_estilo_ttk(self):
        """Intenta aplicar un tema TTK nativo o uno por defecto."""
        style = ttk.Style(self)
        available_themes = style.theme_names()
        current_os = platform.system()
        chosen_theme = None

        print(f"--- Theme Debug ---")
        print(f"Sistema Operativo: {current_os}")
        print(f"Temas TTK Disponibles: {available_themes}")

        # Preferencias de temas por OS
        theme_preferences = {
            "Windows": ["vista", "xpnative", "clam"],
            "Darwin": ["aqua", "clam"], # macOS
            "Linux": ["clam", "alt", "default"] # Opciones comunes en Linux
        }

        prefs = theme_preferences.get(current_os, ["clam", "default"])

        for theme in prefs:
            if theme in available_themes:
                chosen_theme = theme
                break

        if not chosen_theme:
             # Fallback al tema actual o 'default'
             try: chosen_theme = style.theme_use()
             except tk.TclError: chosen_theme = "default" if "default" in available_themes else available_themes[0]


        if chosen_theme:
            print(f"Tema Seleccionado: {chosen_theme}")
            try:
                style.theme_use(chosen_theme)
            except tk.TclError as e:
                print(f"Advertencia: No se pudo aplicar el tema '{chosen_theme}'. Error: {e}")
                # Intentar 'clam' como último recurso si falla el seleccionado
                if chosen_theme != 'clam' and 'clam' in available_themes:
                     try:
                         print("Intentando aplicar tema 'clam' como fallback.")
                         style.theme_use('clam')
                     except tk.TclError: print("No se pudo aplicar 'clam'.")
        else:
            print("Advertencia: No se pudo determinar un tema TTK adecuado.")
        print(f"-------------------")

    def _crear_widgets(self):
        """Crea todos los widgets de la interfaz."""
        # Marco para controles superiores
        self.marco_controles = ttk.LabelFrame(self, text="Controles")

        # Botones de carga
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=self._cargar_diccionario)
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=self._cargar_excel_descripcion, state="disabled") # Deshabilitado al inicio

        # Controles de búsqueda
        self.lbl_busqueda = ttk.Label(self.marco_controles, text="Buscar (use '+' para AND, '-' para OR):")
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50)
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=self._ejecutar_busqueda, state="disabled") # Deshabilitado al inicio

        # Botón de exportar
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar Resultados", command=self._exportar_resultados, state="disabled") # Deshabilitado al inicio

        # Etiquetas para las tablas
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario (Columnas de Búsqueda):")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados de Búsqueda / Descripciones Cargadas:")

        # --- Tabla Diccionario (Vista Previa) ---
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings") # 'headings' para no mostrar la columna #0 vacía
        # Scrollbars
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)

        # --- Tabla Resultados / Descripciones ---
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings")
        # Scrollbars
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)

        # Barra de estado inferior
        self.barra_estado = ttk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W)

    def _configurar_tags_treeview(self):
        """Configura los colores para las filas pares e impares en las tablas."""
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure('par', background=self.color_fila_par)
            tabla.tag_configure('impar', background=self.color_fila_impar)
            # Podríamos añadir más tags si quisiéramos resaltar algo

    def _configurar_grid(self):
        """Organiza los widgets en la ventana usando grid."""
        # Configuración general de filas/columnas de la ventana principal
        self.grid_rowconfigure(2, weight=1) # Fila para tabla diccionario expandible
        self.grid_rowconfigure(4, weight=3) # Fila para tabla resultados más expandible
        self.grid_columnconfigure(0, weight=1) # Columna principal expandible

        # --- Marco de Controles ---
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        # Configurar columnas dentro del marco de controles
        self.marco_controles.grid_columnconfigure(2, weight=1) # Hacer que la entrada de búsqueda se expanda

        # Widgets dentro del marco de controles
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.btn_cargar_descripciones.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.lbl_busqueda.grid(row=1, column=0, columnspan=2, padx=5, pady=(5,0), sticky="w")
        self.entrada_busqueda.grid(row=1, column=2, padx=5, pady=(5,5), sticky="ew") # ew para expandir horizontalmente
        self.btn_buscar.grid(row=1, column=3, padx=5, pady=(5,5))
        # Mover exportar a la misma fila para mejor alineación
        self.btn_exportar.grid(row=1, column=4, padx=(10, 5), pady=5, sticky="e")

        # --- Etiquetas de Tablas ---
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0))
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0))

        # --- Frames y Tablas ---
        # Tabla Diccionario
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1)
        self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew")
        self.scrolly_diccionario.grid(row=0, column=1, sticky="ns")
        self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")

        # Tabla Resultados
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.scrolly_resultados.grid(row=0, column=1, sticky="ns")
        self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # --- Barra de Estado ---
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5, 0))

    def _configurar_eventos(self):
        """Vincula eventos a funciones (ej. Enter en búsqueda)."""
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())
        # Podríamos añadir más eventos, como clics en tablas si quisiéramos más interacción

    def _actualizar_estado(self, mensaje: str):
        """Actualiza el texto de la barra de estado inferior."""
        self.barra_estado.config(text=mensaje)
        self.update_idletasks() # Fuerza la actualización de la UI

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], limite_filas: Optional[int] = None, columnas_a_mostrar: Optional[List[str]] = None):
        """Limpia y rellena una tabla Treeview con datos de un DataFrame."""
        # Limpiar tabla
        try:
            for item in tabla.get_children():
                tabla.delete(item)
            # Resetear columnas si ya tenía
            if tabla["columns"]:
                 tabla["columns"] = ()
        except tk.TclError as e:
            print(f"Advertencia: Error Tcl al limpiar tabla: {e}") # Puede pasar si la tabla es destruida

        if datos is None or datos.empty:
            tabla["columns"] = () # Asegurar que no hay columnas visibles
            return

        # Determinar columnas a mostrar
        if columnas_a_mostrar:
             # Filtrar para asegurar que existen en el DF
             cols_display = [col for col in columnas_a_mostrar if col in datos.columns]
             if not cols_display: # Si ninguna de las especificadas existe, mostrar todas
                  print(f"Advertencia: Ninguna de las columnas especificadas {columnas_a_mostrar} existe en el DataFrame. Mostrando todas.")
                  cols_display = list(datos.columns)
             df_a_mostrar = datos[cols_display]
        else:
             cols_display = list(datos.columns)
             df_a_mostrar = datos

        tabla["columns"] = cols_display

        # Configurar cabeceras y calcular anchos
        for col in cols_display:
            tabla.heading(col, text=str(col), anchor=tk.W)
            try:
                # Calcular ancho basado en contenido y cabecera (muestra limitada para eficiencia)
                sample_size = min(len(df_a_mostrar), 100) # Usar una muestra
                col_data_subset = df_a_mostrar[col].dropna().astype(str) # Convertir a string para medir longitud
                # A veces head(sample_size) puede ser lento en DFs muy grandes, usar iloc es más directo
                # col_data_subset = df_a_mostrar.iloc[:sample_size][col].dropna().astype(str)

                # Ancho del contenido (máximo en la muestra)
                content_width = col_data_subset.str.len().max() if not col_data_subset.empty else 0
                # Ancho de la cabecera
                header_width = len(str(col))

                # Estimar ancho en píxeles (ajustar factor y padding según fuente y gusto)
                width_factor = 7 # Píxeles por carácter (aproximado)
                padding = 20    # Espacio extra
                calculated_width = max(header_width * (width_factor + 2), content_width * width_factor) + padding

                # Limitar ancho máximo y mínimo
                final_width = max(70, min(int(calculated_width), 400)) # Mínimo 70px, máximo 400px

                tabla.column(col, width=final_width, minwidth=70, anchor=tk.W) # tk.W para alinear a la izquierda

            except Exception as e:
                # Fallback si el cálculo falla
                print(f"Advertencia: Error calculando ancho para columna '{col}': {e}. Usando ancho por defecto.")
                tabla.column(col, width=100, minwidth=50, anchor=tk.W)

        # Limitar filas si se especificó
        df_final = df_a_mostrar.head(limite_filas) if limite_filas is not None and len(df_a_mostrar) > limite_filas else df_a_mostrar

        # Insertar filas
        for i, (_, fila) in enumerate(df_final.iterrows()):
            # Convertir todos los valores a string, manejando None/NaN
            valores = [str(v) if pd.notna(v) else "" for v in fila.values]
            tag_actual = 'par' if i % 2 == 0 else 'impar'
            try:
                tabla.insert("", "end", values=valores, tags=(tag_actual,))
            except tk.TclError as e:
                # Puede ocurrir si los datos contienen caracteres inesperados por Tcl
                print(f"Advertencia: Error Tcl al insertar fila {i} con valores {valores[:5]}... : {e}")
                # Intentar insertar valores 'limpios' como fallback
                try:
                    valores_limpios = [str(v).encode('ascii', 'ignore').decode('ascii') for v in fila.values]
                    tabla.insert("", "end", values=valores_limpios, tags=(tag_actual,))
                except Exception as e_insert:
                    print(f"Error crítico al intentar insertar fila limpia {i}: {e_insert}")


    def _cargar_diccionario(self):
        """Maneja la selección y carga del archivo diccionario."""
        ruta = filedialog.askopenfilename(title="Seleccionar Archivo de Diccionario",
                                           filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not ruta: return # Usuario canceló

        self._actualizar_estado(f"Cargando diccionario: {ruta.split('/')[-1]}...")
        # Limpiar tablas anteriores
        self._actualizar_tabla(self.tabla_diccionario, None)
        # Deshabilitar botones mientras carga
        self.btn_cargar_descripciones["state"] = "disabled"
        self.btn_buscar["state"] = "disabled"
        self.btn_exportar["state"] = "disabled"

        if self.motor.cargar_excel_diccionario(ruta):
            df_dic_original = self.motor.datos_diccionario # Ya validado en cargar_excel_diccionario
            if df_dic_original is not None and not df_dic_original.empty:
                num_filas = len(df_dic_original)
                self._actualizar_estado(f"Diccionario cargado ({num_filas} filas). Procesando vista previa...")

                # Obtener nombres de columnas a mostrar (las de búsqueda)
                cols_a_mostrar_nombres = self.motor._obtener_nombres_columnas_busqueda(df_dic_original)
                if cols_a_mostrar_nombres:
                     self.lbl_tabla_diccionario.config(text=f"Vista Previa Diccionario (Columnas: {', '.join(cols_a_mostrar_nombres)}):")
                else: # Si falla obtener nombres (no debería si la carga fue exitosa), mostrar todas
                     cols_a_mostrar_nombres = list(df_dic_original.columns)
                     self.lbl_tabla_diccionario.config(text="Vista Previa Diccionario (Todas las Columnas):")


                # Actualizar tabla de diccionario con las columnas seleccionadas y límite de filas
                self._actualizar_tabla(self.tabla_diccionario, df_dic_original, limite_filas=100, columnas_a_mostrar=cols_a_mostrar_nombres)

                # Actualizar título de ventana y estado de botones
                self.title(f"Buscador - Dic: {ruta.split('/')[-1]}")
                self.btn_cargar_descripciones["state"] = "normal" # Habilitar carga de descripciones
                if self.motor.datos_descripcion is not None: # Si ya había descripciones cargadas
                    self.btn_buscar["state"] = "normal" # Habilitar búsqueda

                mensaje_estado = f"Diccionario '{ruta.split('/')[-1]}' ({num_filas} filas) cargado. Cargue descripciones."
                self._actualizar_estado(mensaje_estado)
                messagebox.showinfo("Éxito", f"Diccionario cargado ({num_filas} filas).\nVista previa muestra máx. 100 filas de columnas de búsqueda.\nAhora puede cargar el archivo de Descripciones.")
            else:
                # Caso: Carga OK pero DF vacío (raro si cargar_excel_diccionario no falló antes)
                 self._actualizar_estado("Diccionario cargado pero parece vacío o inválido.")
                 self.title("Buscador Avanzado")
                 messagebox.showwarning("Archivo Vacío", "El archivo de diccionario se cargó pero está vacío o no tiene las columnas esperadas.")
        else:
            # Caso: Error durante la carga (ya mostrado por ManejadorExcel o MotorBusqueda)
            self._actualizar_estado("Error al cargar diccionario. Verifique el archivo y los mensajes de error.")
            self.title("Buscador Avanzado")
            # Asegurarse de que todo sigue deshabilitado
            self.btn_cargar_descripciones["state"] = "disabled"
            self.btn_buscar["state"] = "disabled"
            self.btn_exportar["state"] = "disabled"

    def _cargar_excel_descripcion(self):
        """Maneja la selección y carga del archivo de descripciones."""
        ruta_descripciones = filedialog.askopenfilename(title="Seleccionar Archivo de Descripciones",
                                                        filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not ruta_descripciones: return # Usuario canceló

        self._actualizar_estado(f"Cargando descripciones: {ruta_descripciones.split('/')[-1]}...")
        # Limpiar tabla de resultados anterior
        self._actualizar_tabla(self.tabla_resultados, None)
        self.resultados_actuales = None
        # Deshabilitar botones mientras carga
        self.btn_buscar["state"] = "disabled"
        self.btn_exportar["state"] = "disabled"

        if self.motor.cargar_excel_descripcion(ruta_descripciones):
            df_desc = self.motor.datos_descripcion
            if df_desc is not None and not df_desc.empty:
                num_filas = len(df_desc)
                self._actualizar_estado(f"Descripciones cargadas ({num_filas} filas). Mostrando todas...")

                # Mostrar todas las descripciones cargadas inicialmente
                self._actualizar_tabla(self.tabla_resultados, df_desc)
                self.resultados_actuales = df_desc.copy() # Guardar para exportar

                # Habilitar botones
                self.btn_exportar["state"] = "normal" # Se pueden exportar las descripciones completas
                if self.motor.datos_diccionario is not None: # Si ya hay diccionario
                    self.btn_buscar["state"] = "normal" # Habilitar búsqueda

                # Actualizar título si tenemos ambos archivos
                if self.motor.archivo_diccionario_actual:
                    dic_name = self.motor.archivo_diccionario_actual.split('/')[-1]
                    desc_name = ruta_descripciones.split('/')[-1]
                    self.title(f"Buscador - Dic: {dic_name} | Desc: {desc_name}")

                mensaje_estado = f"Descripciones '{ruta_descripciones.split('/')[-1]}' ({num_filas} filas) cargadas. Listo para buscar."
                self._actualizar_estado(mensaje_estado)
                messagebox.showinfo("Éxito", f"Descripciones cargadas ({num_filas} filas).\nMostrando en la tabla de resultados.\nAhora puede buscar.")
            else:
                 # Caso: Carga OK pero DF vacío
                 self._actualizar_estado("Archivo de descripciones cargado pero parece vacío.")
                 messagebox.showwarning("Archivo Vacío", "El archivo de descripciones se cargó pero está vacío.")
                 # Mantener botones deshabilitados (excepto cargar diccionario)
                 self.btn_buscar["state"] = "disabled"
                 self.btn_exportar["state"] = "disabled"

        else:
            # Caso: Error durante la carga
            self._actualizar_estado("Error al cargar descripciones. Verifique el archivo y los mensajes de error.")
            # Mantener botones deshabilitados
            self.btn_buscar["state"] = "disabled"
            self.btn_exportar["state"] = "disabled"


    def _ejecutar_busqueda(self):
        """Ejecuta la búsqueda basada en el término introducido."""
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", "Debe cargar ambos archivos (Diccionario y Descripciones) antes de buscar.")
            return

        termino = self.entrada_busqueda.get()

        # --- Búsqueda Vacía: Mostrar todas las descripciones ---
        if not termino.strip():
            messagebox.showinfo("Búsqueda Vacía", "Mostrando todas las descripciones cargadas.")
            df_desc = self.motor.datos_descripcion
            self._actualizar_tabla(self.tabla_resultados, df_desc)
            self.resultados_actuales = df_desc.copy() if df_desc is not None else None
            num_filas = len(df_desc) if df_desc is not None else 0
            self.btn_exportar["state"] = "normal" if num_filas > 0 else "disabled"
            self._actualizar_estado(f"Mostrando todas las {num_filas} descripciones.")
            return

        # --- Ejecutar Búsqueda Normal ---
        self._actualizar_estado(f"Buscando '{termino}'...")
        # Limpiar resultados anteriores y deshabilitar exportar
        self._actualizar_tabla(self.tabla_resultados, None)
        self.resultados_actuales = None
        self.btn_exportar["state"] = "disabled"

        resultados = self.motor.buscar(termino) # Llamada al método principal del motor

        # --- Procesar Resultados ---
        if resultados is None:
            # Error grave o archivos faltantes (ya manejado antes, pero por si acaso)
            self._actualizar_estado(f"Error durante la búsqueda de '{termino}'.")
            # Mensaje de error ya debería haber sido mostrado por el motor

        elif isinstance(resultados, tuple):
            # Caso: Término no encontrado en el diccionario
            df_dic, df_desc = resultados # Los dataframes originales
            self._actualizar_estado(f"'{termino}' no encontrado en el diccionario. ¿Buscar directamente?")

            respuesta = messagebox.askyesno("Sin Coincidencias en Diccionario",
                                            f"El término '{termino}' no se encontró en las columnas de búsqueda del diccionario.\n\n"
                                            "¿Desea buscar directamente en todas las columnas de las descripciones?")
            if respuesta:
                # --- Ejecutar Búsqueda Directa ---
                self._actualizar_estado(f"Buscando '{termino}' directamente en descripciones...")
                resultados_directos = self.motor.buscar_en_descripciones_directo(termino)
                self._actualizar_tabla(self.tabla_resultados, resultados_directos)
                self.resultados_actuales = resultados_directos # Guardar para exportar
                num_res = len(resultados_directos) if resultados_directos is not None else 0

                if num_res > 0:
                    self.btn_exportar["state"] = "normal"
                    self._actualizar_estado(f"Búsqueda directa de '{termino}' completada: {num_res} resultados.")
                    # >>> INICIO: Demostración ExtractorMagnitud (para búsqueda directa) <<<
                    if len(resultados_directos.columns) > 0: # Asegurar que hay columnas
                        try:
                            # Intentar obtener texto de la primera celda del primer resultado
                            texto_primer_resultado = str(resultados_directos.iloc[0, 0])
                            print("\n--- DEMO Extractor Magnitudes (1er resultado directo) ---")
                            # >>> INICIO: Usar self.extractor_magnitud.magnitudes <<<
                            for mag_demo in self.extractor_magnitud.magnitudes: # Usar la lista de la instancia
                            # <<< FIN: Usar self.extractor_magnitud.magnitudes <<<
                                cantidad_extraida = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag_demo, texto_primer_resultado)
                                if cantidad_extraida is not None:
                                    print(f"  -> Magnitud '{mag_demo}' encontrada: Cantidad = {cantidad_extraida}")
                                # else: # Descomentar si quieres ver las que NO encuentra
                                #     print(f"  -> Magnitud '{mag_demo}' NO encontrada en: '{texto_primer_resultado[:50]}...'")
                            print("--- FIN DEMO ---")
                        except IndexError:
                             print("Advertencia: No se pudo acceder a la primera celda para la demo del extractor.")
                        except Exception as e_demo:
                            print(f"Error en la demostración del extractor (búsqueda directa): {e_demo}")
                    # <<< FIN: Demostración ExtractorMagnitud (para búsqueda directa) <<<
                else:
                    # Búsqueda directa sin resultados
                    messagebox.showinfo("Sin Coincidencias", f"La búsqueda directa de '{termino}' tampoco encontró resultados.")
                    self._actualizar_estado(f"Búsqueda directa de '{termino}' completada: 0 resultados.")
            else:
                # Usuario no quiso búsqueda directa
                self._actualizar_estado(f"Búsqueda de '{termino}' cancelada (sin coincidencias en diccionario).")

        elif isinstance(resultados, pd.DataFrame):
            # --- Búsqueda Normal Exitosa (puede tener 0 o más filas) ---
            self.resultados_actuales = resultados # Guardar para exportar
            num_res = len(resultados)
            self._actualizar_tabla(self.tabla_resultados, resultados)

            if num_res > 0:
                self.btn_exportar["state"] = "normal"
                self._actualizar_estado(f"Búsqueda de '{termino}' completada: {num_res} resultados encontrados.")
                # >>> INICIO: Demostración ExtractorMagnitud (para búsqueda normal) <<<
                if len(resultados.columns) > 0: # Asegurar que hay columnas
                    try:
                        # Intentar obtener texto de la primera celda del primer resultado
                        texto_primer_resultado = str(resultados.iloc[0, 0])
                        print("\n--- DEMO Extractor Magnitudes (1er resultado normal) ---")
                        # >>> INICIO: Usar self.extractor_magnitud.magnitudes <<<
                        for mag_demo in self.extractor_magnitud.magnitudes: # Usar la lista de la instancia
                        # <<< FIN: Usar self.extractor_magnitud.magnitudes <<<
                            cantidad_extraida = self.extractor_magnitud.buscar_cantidad_para_magnitud(mag_demo, texto_primer_resultado)
                            if cantidad_extraida is not None:
                                print(f"  -> Magnitud '{mag_demo}' encontrada: Cantidad = {cantidad_extraida}")
                            # else: # Descomentar si quieres ver las que NO encuentra
                            #      print(f"  -> Magnitud '{mag_demo}' NO encontrada en: '{texto_primer_resultado[:50]}...'")
                        print("--- FIN DEMO ---")
                    except IndexError:
                         print("Advertencia: No se pudo acceder a la primera celda para la demo del extractor.")
                    except Exception as e_demo:
                        print(f"Error en la demostración del extractor (búsqueda normal): {e_demo}")
                # <<< FIN: Demostración ExtractorMagnitud (para búsqueda normal) <<<
            else:
                # Búsqueda normal encontró 0 resultados en descripciones
                messagebox.showinfo("Sin Coincidencias Finales",
                                    f"Se encontraron términos en el diccionario para '{termino}', pero no se hallaron coincidencias finales en las descripciones.")
                self._actualizar_estado(f"Búsqueda de '{termino}' completada: 0 resultados en descripciones.")
        else:
            # Tipo de resultado inesperado (no debería ocurrir)
            self._actualizar_estado(f"Error: Tipo de resultado inesperado ({type(resultados)}) tras buscar '{termino}'.")
            print(f"Resultado inesperado de búsqueda: {resultados}")


    def _exportar_resultados(self):
        """Guarda el DataFrame de resultados actual en un archivo Excel o CSV."""
        if self.resultados_actuales is None or self.resultados_actuales.empty:
            messagebox.showwarning("Sin Resultados", "No hay resultados para exportar.")
            return

        # Tipos de archivo para guardar
        file_types = [("Archivo Excel (.xlsx)", "*.xlsx"),
                      ("Archivo CSV (UTF-8) (.csv)", "*.csv"),
                      ("Excel 97-2003 (.xls)", "*.xls")] # Ofrecer xls como opción

        # Pedir ruta al usuario
        ruta = filedialog.asksaveasfilename(title="Guardar Resultados Como",
                                            defaultextension=".xlsx", # Extensión por defecto
                                            filetypes=file_types)
        if not ruta: return # Usuario canceló

        self._actualizar_estado("Exportando resultados...")

        try:
            extension = ruta.split('.')[-1].lower()

            if extension == 'csv':
                # Guardar como CSV con codificación UTF-8-SIG (mejor compatibilidad con Excel para caracteres especiales)
                self.resultados_actuales.to_csv(ruta, index=False, encoding='utf-8-sig')
            elif extension == 'xlsx':
                # Guardar como XLSX usando openpyxl
                self.resultados_actuales.to_excel(ruta, index=False, engine='openpyxl')
            elif extension == 'xls':
                # Guardar como XLS (formato antiguo, requiere xlwt)
                try:
                    import xlwt # Necesita instalarse: pip install xlwt
                    # Advertir y truncar si excede el límite de filas de XLS
                    max_rows_xls = 65535
                    if len(self.resultados_actuales) > max_rows_xls:
                        messagebox.showwarning("Límite de Filas Excedido (.xls)",
                                               f"El formato .xls solo soporta hasta {max_rows_xls} filas.\n"
                                               f"Sus resultados ({len(self.resultados_actuales)} filas) serán truncados.")
                        df_to_export = self.resultados_actuales.head(max_rows_xls)
                    else:
                        df_to_export = self.resultados_actuales
                    df_to_export.to_excel(ruta, index=False, engine='xlwt')
                except ImportError:
                    messagebox.showerror("Librería Faltante",
                                         "Para exportar a formato .xls, necesita instalar 'xlwt':\n`pip install xlwt`")
                    self._actualizar_estado("Error al exportar: Falta 'xlwt'.")
                    return # No continuar si falta la librería
                except Exception as ex_xls:
                    # Capturar otros errores específicos de xlwt si ocurren
                    raise ex_xls # Relanzar para el manejo general de excepciones
            else:
                # Extensión no soportada
                messagebox.showerror("Extensión Inválida", f"Extensión de archivo no soportada para exportar: {extension}")
                self._actualizar_estado("Error al exportar: Extensión inválida.")
                return

            # Éxito en la exportación
            messagebox.showinfo("Éxito", f"Resultados ({len(self.resultados_actuales)} filas) exportados correctamente a:\n{ruta}")
            self._actualizar_estado(f"Resultados exportados a {ruta.split('/')[-1]}")

        except ImportError as imp_err:
            # Manejar error si falta openpyxl para .xlsx
            if 'openpyxl' in str(imp_err) and ruta.endswith('.xlsx'):
                messagebox.showerror("Librería Faltante",
                                     "Para exportar a formato .xlsx, necesita instalar 'openpyxl':\n`pip install openpyxl`")
                self._actualizar_estado("Error al exportar: Falta 'openpyxl'.")
            else:
                # Otro error de importación inesperado
                messagebox.showerror("Error de Importación", f"Falta una librería necesaria para exportar:\n{imp_err}")
                self._actualizar_estado("Error al exportar: Librería faltante.")
                print(traceback.format_exc())
        except Exception as e:
            # Error general durante la escritura del archivo
            error_detallado = traceback.format_exc()
            print(f"Error detallado de exportación:\n{error_detallado}")
            messagebox.showerror("Error de Exportación", f"No se pudo guardar el archivo:\n{e}")
            self._actualizar_estado("Error al exportar resultados.")


if __name__ == "__main__":
    # Comprobación inicial de dependencias críticas
    missing_libs = []
    try:
        import pandas as pd
    except ImportError:
        missing_libs.append("pandas")
    try:
        # openpyxl es necesario para leer/escribir .xlsx
        import openpyxl
    except ImportError:
        missing_libs.append("openpyxl (para archivos .xlsx)")
    # xlwt es opcional (solo para exportar a .xls) - no lo comprobamos aquí
    # unicodedata es estándar, no necesita comprobación

    if missing_libs:
        # Si faltan librerías, mostrar mensaje y salir (sin crear ventana principal)
        root_error = tk.Tk()
        root_error.withdraw() # Ocultar ventana raíz vacía
        messagebox.showerror("Dependencias Faltantes",
                             f"Error: Faltan librerías esenciales para ejecutar la aplicación:\n\n"
                             f"- {chr(10)}- ".join(missing_libs) + # chr(10) es salto de línea
                             f"\n\nPor favor, instálalas usando pip (ej: pip install pandas openpyxl) y reinicia la aplicación.")
        root_error.destroy()
        exit(1) # Salir del script

    # Si las dependencias están OK, iniciar la aplicación
    app = InterfazGrafica()
    app.mainloop()
