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
    # Callable, # No se usa explícitamente Callable como tipo directo de parámetro/retorno en las funciones principales
    Dict,
    Any,
    # Literal, # No se usa explícitamente Literal
)
from enum import Enum, auto
# import traceback # No se usa explícitamente traceback en el código final unificado
import platform
import unicodedata
import logging
import json
import os
from pathlib import Path
import string

# --- Configuración del Logging ---
logger = logging.getLogger(__name__)


# --- Enumeraciones ---
class OrigenResultados(Enum):
    NINGUNO = 0
    VIA_DICCIONARIO_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS = auto()
    VIA_DICCIONARIO_SIN_RESULTADOS_DESC = auto()
    DIRECTO_DESCRIPCION_CON_RESULTADOS = auto()
    DIRECTO_DESCRIPCION_VACIA = auto() # Usado también si la búsqueda es vacía y devuelve todo
    ERROR_CARGA_DICCIONARIO = auto()
    ERROR_CARGA_DESCRIPCION = auto()
    ERROR_CONFIGURACION_COLUMNAS_DICC = auto()
    ERROR_CONFIGURACION_COLUMNAS_DESC = auto()
    ERROR_BUSQUEDA_INTERNA_MOTOR = auto()
    TERMINO_INVALIDO = auto() # Usado si la query post-negación es inválida o si la negación lleva a error

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
    def es_error_operacional(self) -> bool:
        return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR

    @property
    def es_termino_invalido(self) -> bool:
        return self == OrigenResultados.TERMINO_INVALIDO


# --- Clases de Utilidad ---
class ExtractorMagnitud:
    MAPEO_MAGNITUDES_PREDEFINIDO = (
        {}
    ) # Se puede expandir con unidades por defecto si es necesario

    def __init__(self, mapeo_magnitudes: Optional[Dict[str, List[str]]] = None):
        self.sinonimo_a_canonico_normalizado: Dict[str, str] = {}
        mapeo_a_usar = (
            mapeo_magnitudes
            if mapeo_magnitudes is not None
            else self.MAPEO_MAGNITUDES_PREDEFINIDO
        )

        for forma_canonica, lista_sinonimos in mapeo_a_usar.items():
            canonico_norm = self._normalizar_texto(forma_canonica)
            if not canonico_norm: # Si la forma canónica normalizada es vacía, se ignora
                logger.warning(
                    f"Forma canónica '{forma_canonica}' resultó vacía tras normalizar. Se ignora."
                )
                continue

            self.sinonimo_a_canonico_normalizado[canonico_norm] = canonico_norm

            for sinonimo in lista_sinonimos:
                sinonimo_norm = self._normalizar_texto(sinonimo)
                if sinonimo_norm: # Solo procesar si el sinónimo normalizado no es vacío
                    if (
                        sinonimo_norm in self.sinonimo_a_canonico_normalizado
                        and self.sinonimo_a_canonico_normalizado[sinonimo_norm]
                        != canonico_norm
                    ):
                        logger.warning(
                            f"Conflicto de mapeo: El sinónimo normalizado '{sinonimo_norm}' (de '{sinonimo}' para '{forma_canonica}') "
                            f"ya está mapeado a '{self.sinonimo_a_canonico_normalizado[sinonimo_norm]}'. "
                            f"Se sobrescribirá con el mapeo a '{canonico_norm}'. "
                            "Revise su MAPEO_MAGNITUDES para evitar ambigüedades."
                        )
                    self.sinonimo_a_canonico_normalizado[sinonimo_norm] = canonico_norm
        logger.debug(
            f"ExtractorMagnitud inicializado con mapeo: {self.sinonimo_a_canonico_normalizado}"
        )

    @staticmethod
    def _normalizar_texto(texto: str) -> str:
        """Normaliza texto específicamente para unidades/magnitudes: Mayúsculas, sin acentos, mantiene alfanuméricos, espacios y puntuación relevante para unidades."""
        if not isinstance(texto, str) or not texto:
            return ""
        try:
            texto_upper = texto.upper()
            forma_normalizada = unicodedata.normalize("NFKD", texto_upper)
            # Eliminar caracteres de combinación (acentos)
            # Mantener letras, números, espacios y puntuación que podría ser parte de una unidad (ej. 'G/S', 'M.2')
            res = "".join(c for c in forma_normalizada if not unicodedata.combining(c) and (c.isalnum() or c.isspace() or c in string.punctuation))
            return ' '.join(res.split()) # Normalizar múltiples espacios a uno solo
        except TypeError: # En caso de que `texto` no sea procesable por normalize
            return ""

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        if not texto_unidad: # Cubre None o string vacío
            return None
        normalizada = self._normalizar_texto(texto_unidad)
        if not normalizada: # Si la normalización resulta en string vacío
            return None
        return self.sinonimo_a_canonico_normalizado.get(normalizada)


class ManejadorExcel:
    @staticmethod
    def cargar_excel(
        ruta_archivo: Union[str, Path],
    ) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        ruta = Path(ruta_archivo)
        logger.info(f"Intentando cargar archivo Excel: {ruta}")
        if not ruta.exists():
            error_msg = f"¡Archivo no encontrado! Ruta: {ruta}"
            logger.error(error_msg)
            return None, error_msg
        try:
            # Determinar el motor basado en la extensión del archivo
            engine = "openpyxl" if ruta.suffix.lower() == ".xlsx" else None # xlrd para .xls si es necesario
            df = pd.read_excel(ruta, engine=engine)
            logger.info(f"Archivo '{ruta.name}' cargado ({len(df)} filas).")
            return df, None
        except Exception as e:
            error_msg = (
                f"No se pudo cargar el archivo:\n{ruta}\n\nError: {e}\n\n"
                "Posibles causas:\n"
                "- El archivo está siendo usado por otro programa.\n"
                "- Formato de archivo no soportado o corrupto.\n"
                "- Para archivos .xlsx, asegúrese de tener 'openpyxl' instalado.\n"
                "- Para archivos .xls más antiguos, podría necesitar 'xlrd'."
            )
            logger.exception(f"Error inesperado al cargar archivo Excel: {ruta}") # .exception incluye traceback
            return None, error_msg


# --- Clase MotorBusqueda ---
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
            f"MotorBusqueda inicializado. Índices búsqueda diccionario: {self.indices_columnas_busqueda_dic or 'Todas las de texto/objeto'}"
        )

        # Patrones Regex precompilados para eficiencia
        self.patron_comparacion = re.compile(
            # Ejemplo: "> 100W", "<= 50.5 KG", "= 20"
            r"^\s*([<>]=?)\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$"
        )
        self.patron_rango = re.compile(
            # Ejemplo: "10 - 20 V", "5.5 - 10.2"
            r"^\s*(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$"
        )
        self.patron_termino_negado = re.compile(
            # Ejemplo: #palabra, #"frase completa", #"123", #"término-con-guion"
            # Captura (frase entre comillas) O (palabra sin espacios o con guiones/puntos internos)
             r'#\s*(?:\"([^\"]+)\"|([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\.\-\_]+))',
            re.IGNORECASE | re.UNICODE
        )
        self.patron_num_unidad_df = re.compile(
            # Para extraer números y unidades de las celdas del DataFrame
            # Ejemplo: "10W", "20.5A", "50 V", "100-240VAC"
            r"(\d+(?:[.,]\d+)?)[\s\-]*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?"
        )
        self.extractor_magnitud = ExtractorMagnitud() # Inicializar con mapeo vacío o predefinido

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            return False, error_msg_carga # error_msg_carga ya contendrá un mensaje detallado

        # (Re)Inicializar ExtractorMagnitud con valores de la primera columna del diccionario cargado
        # Esto permite que las unidades sean dinámicas según el diccionario.
        col0_vals = []
        if df_cargado.shape[1] > 0: # Asegurar que hay al menos una columna
            col0_vals = df_cargado.iloc[:, 0].dropna().unique() # Tomar valores únicos y no nulos

        mapeo_dinamico = {}
        for val in col0_vals:
            val_str = str(val).strip()
            if val_str: # Solo añadir si el valor no es una cadena vacía
                # Cada valor de la primera columna se considera una forma canónica y un sinónimo de sí mismo.
                mapeo_dinamico[val_str] = [val_str] 
        self.extractor_magnitud = ExtractorMagnitud(mapeo_magnitudes=mapeo_dinamico)
        logger.info(f"Extractor de magnitudes actualizado con {len(mapeo_dinamico)} mapeos desde el diccionario.")

        valido, msg_val_cols = self._validar_columnas_df(
            df=df_cargado,
            indices_cfg=self.indices_columnas_busqueda_dic,
            nombre_df_log="diccionario",
        )
        if not valido:
            logger.warning(
                f"Validación de columnas del diccionario '{ruta.name}' fallida. Causa: {msg_val_cols}"
            )
            self.datos_diccionario = None # Asegurar que no se usa un DF inválido
            self.archivo_diccionario_actual = None
            return (
                False,
                msg_val_cols or "Validación de columnas del diccionario fallida (mensaje genérico).",
            )

        self.datos_diccionario = df_cargado
        self.archivo_diccionario_actual = ruta
        logger.info(f"Diccionario '{ruta.name}' cargado y validado exitosamente.")
        return True, None

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_descripcion = None
            self.archivo_descripcion_actual = None
            return False, error_msg_carga # error_msg_carga ya contendrá un mensaje detallado

        # Podríamos añadir validación de columnas para descripciones si fuera necesario en el futuro
        # Por ahora, se asume que si carga, es usable.
        self.datos_descripcion = df_cargado
        self.archivo_descripcion_actual = ruta
        logger.info(f"Archivo de descripciones '{ruta.name}' cargado exitosamente.")
        return True, None

    def _validar_columnas_df(
        self, df: pd.DataFrame, indices_cfg: List[int], nombre_df_log: str
    ) -> Tuple[bool, Optional[str]]:
        # df ya no puede ser None aquí si se llama después de la carga exitosa
        num_cols_df = len(df.columns)

        if not indices_cfg or indices_cfg == [-1]: # Configuración para buscar en todas las columnas (o todas las de texto)
            if num_cols_df == 0:
                msg = f"El archivo del {nombre_df_log} está vacío o no contiene columnas (configuración: buscar en todas)."
                logger.error(msg)
                return False, msg
            return True, None # Es válido si hay columnas y la config es "todas"

        # Validar que los índices configurados sean enteros no negativos
        if not all(isinstance(idx, int) and idx >= 0 for idx in indices_cfg):
            msg = f"Configuración de índices para {nombre_df_log} inválida: {indices_cfg}. Deben ser enteros no negativos."
            logger.error(msg)
            return False, msg

        max_indice_requerido = max(indices_cfg) # Ya sabemos que indices_cfg no está vacío y son >= 0

        if num_cols_df == 0: # Caso improbable si ya pasó el chequeo de "todas", pero por seguridad.
            msg = f"El {nombre_df_log} no tiene columnas, pero se especificaron índices."
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
        self, df: pd.DataFrame, indices_cfg: List[int], nombre_df_log: str
    ) -> Tuple[Optional[List[str]], Optional[str]]:
        # df ya no puede ser None aquí
        columnas_disponibles = df.columns
        num_cols_df = len(columnas_disponibles)

        if not indices_cfg or indices_cfg == [-1]: # Si es [-1] o lista vacía, buscar en todas las de texto/objeto
            cols_texto_obj = [
                col
                for col in df.columns
                if pd.api.types.is_string_dtype(df[col])
                or pd.api.types.is_object_dtype(df[col]) # object dtype a menudo contiene texto
            ]
            if cols_texto_obj:
                logger.info(
                    f"Buscando en columnas de texto/object (detectadas) del {nombre_df_log}: {cols_texto_obj}"
                )
                return cols_texto_obj, None
            elif num_cols_df > 0: # Si no hay de texto/objeto pero hay columnas, usar todas como fallback
                logger.warning(
                    f"No se encontraron columnas de texto/object en {nombre_df_log} (config: 'todas'). "
                    f"Se usarán todas las {num_cols_df} columnas como fallback."
                )
                return list(df.columns), None
            else: # No hay columnas en el DF
                msg = f"El DataFrame del {nombre_df_log} no tiene columnas (config: 'todas')."
                logger.error(msg)
                return None, msg

        # Si se especificaron índices explícitos
        nombres_columnas_seleccionadas = []
        indices_validos_usados = []
        for indice in indices_cfg: # Ya validados como int >= 0 y dentro del rango por _validar_columnas_df
            nombres_columnas_seleccionadas.append(columnas_disponibles[indice])
            indices_validos_usados.append(indice)
        
        # Este chequeo es redundante si _validar_columnas_df se llamó antes, pero por si acaso.
        if not nombres_columnas_seleccionadas:
            msg = f"No se pudieron seleccionar columnas en {nombre_df_log} con los índices configurados: {indices_cfg}"
            logger.error(msg)
            return None, msg

        logger.debug(
            f"Se buscará en columnas del {nombre_df_log}: {nombres_columnas_seleccionadas} (índices: {indices_validos_usados})"
        )
        return nombres_columnas_seleccionadas, None
    
    def _normalizar_para_busqueda(self, texto: str) -> str:
        """Normaliza texto para búsqueda general: sin acentos, mayúsculas, y strip."""
        if not isinstance(texto, str) or not texto:
            return ""
        try:
            # Descomponer caracteres acentuados en su forma base + diacrítico
            texto_norm = unicodedata.normalize('NFKD', texto)
            # Eliminar los diacríticos (caracteres de combinación)
            texto_sin_acentos = "".join([c for c in texto_norm if not unicodedata.combining(c)])
            # Convertir a mayúsculas y quitar espacios al inicio/final
            return texto_sin_acentos.upper().strip()
        except TypeError: # Por si acaso `texto` no es procesable por normalize
            return ""

    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, columnas_a_buscar: List[str], texto_complejo: str) -> Tuple[pd.DataFrame, str]:
        """
        Procesa la consulta para extraer términos de negación, filtra el DataFrame
        original para excluir filas que coincidan con estos términos negados,
        y devuelve el DataFrame filtrado junto con la parte de la consulta
        que corresponde a términos positivos.
        """
        texto_limpio = texto_complejo.strip()
        # Si el texto de búsqueda está vacío, no hay nada que negar ni positivo que extraer.
        if not texto_limpio:
            return df_original.copy() if df_original is not None else pd.DataFrame(), ""

        # Si el DataFrame original es None o está vacío, no se puede aplicar negación.
        # Se devuelve un DataFrame vacío (con las columnas originales si es posible) y el texto limpio como positivo.
        if df_original is None or df_original.empty:
            empty_df_cols = df_original.columns if df_original is not None else []
            return pd.DataFrame(columns=empty_df_cols), texto_limpio

        terminos_negados_final = [] # Lista para almacenar los términos negados normalizados
        partes_positivas = [] # Lista para construir la cadena de términos positivos
        last_end = 0 # Puntero para reconstruir la cadena de términos positivos

        # Iterar sobre todas las coincidencias del patrón de negación en el texto limpio
        for match in self.patron_termino_negado.finditer(texto_limpio):
            partes_positivas.append(texto_limpio[last_end:match.start()]) # Añadir el segmento antes de la negación
            last_end = match.end() # Actualizar el puntero

            termino_negado_raw = match.group(1) or match.group(2) # Grupo 1 para frases, Grupo 2 para palabras
            if termino_negado_raw:
                # Normalizar el término negado para la búsqueda
                termino_norm = self._normalizar_para_busqueda(termino_negado_raw.strip('"')) # Quitar comillas si es frase
                if termino_norm and termino_norm not in terminos_negados_final: # Evitar duplicados y vacíos
                    terminos_negados_final.append(termino_norm)
        
        partes_positivas.append(texto_limpio[last_end:]) # Añadir el segmento restante después de la última negación
        terminos_positivos_str = "".join(partes_positivas).strip() # Unir todas las partes positivas
        terminos_positivos_str = ' '.join(terminos_positivos_str.split()) # Normalizar espacios múltiples a uno solo

        df_filtrado = df_original.copy() # Trabajar con una copia del DataFrame

        # Si no se encontraron términos negados, devolver el DataFrame original y la cadena de positivos.
        if not terminos_negados_final:
            return df_filtrado, terminos_positivos_str

        # Si hay términos negados, proceder a filtrar el DataFrame
        mascara_filas_a_excluir_total = pd.Series(False, index=df_filtrado.index) # Máscara para acumular filas a excluir

        for termino_negado_norm in terminos_negados_final:
            if not termino_negado_norm: # Seguridad, aunque el `if termino_norm` anterior debería cubrirlo
                continue
            
            logger.debug(f"Aplicando negación para el término: '{termino_negado_norm}'")
            mascara_coincidencia_termino_actual = pd.Series(False, index=df_filtrado.index) # Máscara para este término negado

            for col_nombre in columnas_a_buscar:
                if col_nombre not in df_filtrado.columns:
                    logger.warning(f"Columna de negación '{col_nombre}' no encontrada en el DataFrame. Se omite para este término.")
                    continue
                
                try:
                    # Normalizar el contenido de la columna del DataFrame para la comparación
                    serie_normalizada_df = df_filtrado[col_nombre].astype(str).map(
                        self._normalizar_para_busqueda
                    )
                    # Patrón regex para buscar el término negado como palabra completa (whole word)
                    # re.escape maneja caracteres especiales en el término para que no interfieran con el regex
                    patron_regex_termino = r"\b" + re.escape(termino_negado_norm) + r"\b"
                    # Acumular (OR lógico) las coincidencias dentro de las columnas para este término negado
                    mascara_coincidencia_termino_actual |= serie_normalizada_df.str.contains(
                        patron_regex_termino, regex=True, na=False # na=False trata NaN como no coincidencia
                    )
                except Exception as e:
                    logger.error(f"Error procesando columna '{col_nombre}' para negación de '{termino_negado_norm}': {e}")
                    # Continuar con la siguiente columna o término en caso de error
            
            # Acumular (OR lógico) las filas que coinciden con CUALQUIERA de los términos negados
            mascara_filas_a_excluir_total |= mascara_coincidencia_termino_actual
            
        # Excluir las filas marcadas por la máscara de negación
        df_resultado_final = df_filtrado[~mascara_filas_a_excluir_total]
        logger.info(f"Proceso de negación finalizado. Filas antes: {len(df_filtrado)}, Filas después: {len(df_resultado_final)}. Términos negados procesados: {terminos_negados_final}")
        
        return df_resultado_final, terminos_positivos_str

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        """Descompone la cadena de búsqueda por operadores OR ('|' o '/')."""
        texto_limpio = texto_complejo.strip()
        if not texto_limpio:
            return "OR", [] # Si no hay texto, no hay segmentos

        # Priorizar '|' como separador OR, luego '/'
        if "|" in texto_limpio:
            segmentos = [
                s.strip() for s in re.split(r"\s*\|\s*", texto_limpio) if s.strip() # Dividir y quitar vacíos
            ]
            return "OR", segmentos
        elif "/" in texto_limpio:
            segmentos = [
                s.strip() for s in re.split(r"\s*/\s*", texto_limpio) if s.strip()
            ]
            return "OR", segmentos
        else:
            # Si no hay OR, toda la cadena es un único segmento que se tratará con AND
            return "AND", [texto_limpio] 

    def _descomponer_nivel2_and(
        self, termino_segmento_n1: str
    ) -> Tuple[str, List[str]]:
        """Descompone un segmento (proveniente de OR o la query completa) por el operador AND ('+')."""
        termino_limpio = termino_segmento_n1.strip()
        if not termino_limpio:
            return "AND", []

        op_principal_interno = "AND" # Por defecto, si solo hay un término o se usa '+'
        
        terminos_brutos_finales = []
        if "+" in termino_limpio:
            # Para splitear por '+' pero preservar términos numéricos/comparación como unidades.
            # Se usa una máquina de estados simplificada para identificar los separadores '+' que actúan como AND.
            # Esto es para evitar splitear incorrectamente, por ejemplo, un rango como "10+A - 20+A" si '+' fuera parte de la unidad.
            # Sin embargo, la estrategia actual con patrones de comparación y rango debería manejar esto antes.
            # Un split más simple, asumiendo que '+' como AND estará espaciado o que los términos complejos ya fueron analizados.
            
            # Máquina de estados para un análisis cuidadoso de '+'
            # Estado 0: general, 1: en operador comp, 2: en número, 3: en unidad post-número
            estado = 0  
            termino_actual_maquina = []
            pos = 0
            while pos < len(termino_limpio):
                char = termino_limpio[pos]
                
                if estado == 0: # Estado General
                    if char == '+': # Separador AND detectado
                        sub_termino = "".join(termino_actual_maquina).strip()
                        if sub_termino: terminos_brutos_finales.append(sub_termino)
                        termino_actual_maquina = [] # Reiniciar para el siguiente término
                    elif char in "<>=": estado = 1; termino_actual_maquina.append(char)
                    elif char.isdigit(): estado = 2; termino_actual_maquina.append(char)
                    else: termino_actual_maquina.append(char) # Caracter general
                
                elif estado == 1: # En operador de comparación < > = <= >=
                    termino_actual_maquina.append(char)
                    if char.isdigit() or char in ".,": estado = 2 # Transición a número
                    elif char.isspace(): # Espacio después de operador
                        # Si el operador está completo (ej: <=) y sigue un espacio, es válido.
                        # Si no viene un número después, el operador podría ser parte de un término string.
                        # Una validación simple: si lo que sigue no es un dígito, volver a general.
                        current_op_candidate = "".join(termino_actual_maquina).strip()
                        if not any(c.isdigit() for c in termino_limpio[pos+1:]): 
                             if current_op_candidate not in ['<','>','<=','>=','=']:
                                estado = 0 # No es un operador válido seguido de no-dígito
                    elif not (char.isalnum() or char in "."): # Si no es alfanumérico ni punto (fin de posible operador)
                        estado = 0 

                elif estado == 2: # En número
                    termino_actual_maquina.append(char)
                    if not (char.isdigit() or char in ".,"): # Si deja de ser dígito o separador decimal/miles
                        if char.isalpha() or char in ['µ', 'Ω']: estado = 3 # Podría ser una unidad
                        else: estado = 0 # Fin del número, carácter general
                
                elif estado == 3: # En unidad post-número
                    termino_actual_maquina.append(char)
                    if not (char.isalnum() or char in ['µ', 'Ω', '.', '/', '-']): # Fin de la unidad
                        estado = 0 
                pos += 1

            # Añadir el último término acumulado
            sub_termino_final = "".join(termino_actual_maquina).strip()
            if sub_termino_final: terminos_brutos_finales.append(sub_termino_final)
            
            # Caso especial: si la query es solo '+' o está vacía después del split
            if not terminos_brutos_finales and termino_limpio.strip() == '+':
                 return op_principal_interno, [] 
        else:
            terminos_brutos_finales = [termino_limpio] # Un solo término si no hay '+'

        return op_principal_interno, [t for t in terminos_brutos_finales if t] # Filtrar términos vacíos


    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        """Analiza una lista de términos brutos (strings) y los clasifica (numérico, rango, string)."""
        palabras_analizadas = []
        for term_orig_bruto in terminos_brutos:
            term_orig = str(term_orig_bruto).strip() # Asegurar string y sin espacios extra
            if not term_orig: # Saltar términos vacíos
                continue

            item_analizado: Dict[str, Any] = {"original": term_orig}
            
            # Intentar hacer match con patrón de comparación (>, <, >=, <=, =)
            match_comp = self.patron_comparacion.match(term_orig)
            # Intentar hacer match con patrón de rango (num - num)
            match_range = self.patron_rango.match(term_orig)

            if match_comp:
                op, v_str, unidad_str = match_comp.groups() # El cuarto grupo (.*) no se usa
                v_num = self._parse_numero(v_str)
                if v_num is not None:
                    op_map = {">": "gt", "<": "lt", ">=": "ge", "<=": "le", "=": "eq"}
                    unidad_canon_comp = None
                    if unidad_str and unidad_str.strip(): # Procesar unidad solo si existe y no es vacía
                        unidad_canon_comp = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str.strip())
                        if unidad_canon_comp is None:
                             logger.debug(f"Unidad de búsqueda '{unidad_str.strip()}' en '{term_orig}' no reconocida por el mapeo. Comparación numérica sin filtro de unidad específico.")
                    item_analizado.update({"tipo": op_map.get(op, "str"), "valor": v_num, "unidad_busqueda": unidad_canon_comp})
                else: # Falló el parseo del número, tratar como string normalizado
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(term_orig)})
            elif match_range:
                v1_str, v2_str, unidad_rango_str = match_range.groups()
                v1, v2 = self._parse_numero(v1_str), self._parse_numero(v2_str)
                if v1 is not None and v2 is not None:
                    unidad_canon_range = None
                    if unidad_rango_str and unidad_rango_str.strip(): # Procesar unidad solo si existe y no es vacía
                        unidad_canon_range = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_rango_str.strip())
                        if unidad_canon_range is None:
                            logger.debug(f"Unidad de rango '{unidad_rango_str.strip()}' en '{term_orig}' no reconocida. Rango sin filtro de unidad específico.")
                    item_analizado.update({"tipo": "range", "valor": sorted([v1, v2]), "unidad_busqueda": unidad_canon_range}) # Guardar rango como [min, max]
                else: # Falló parseo de números del rango, tratar como string normalizado
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(term_orig)})
            else: # Es un término de string normal
                item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(term_orig)})
            
            palabras_analizadas.append(item_analizado)
        logger.debug(f"Términos analizados (post-negación): {palabras_analizadas}")
        return palabras_analizadas

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        """Convierte un string (o número) a float. Maneja comas como separadores decimales."""
        if isinstance(num_str, (int, float)): # Si ya es numérico
            return float(num_str)
        if not isinstance(num_str, str): # Si no es string ni numérico
            return None
        try:
            return float(num_str.replace(",", ".")) # Reemplazar coma por punto para conversión a float
        except ValueError:
            return None # No se pudo convertir a float

    def _generar_mascara_para_un_termino(
        self,
        df: pd.DataFrame,
        cols_a_buscar: List[str],
        termino_analizado: Dict[str, Any],
    ) -> pd.Series:
        """Genera una máscara booleana para un único término analizado sobre el DataFrame."""
        tipo_sub = termino_analizado["tipo"]
        valor_sub = termino_analizado["valor"] # Ya está parseado (float, list de floats, o string normalizado)
        unidad_sub_requerida_canon = termino_analizado.get("unidad_busqueda") # Puede ser None
        
        mascara_total_subtermino = pd.Series(False, index=df.index) # Iniciar con todo False (para OR entre columnas)

        for col_nombre in cols_a_buscar:
            if col_nombre not in df.columns:
                logger.warning(f"Columna '{col_nombre}' no encontrada en DataFrame al generar máscara. Se omite.")
                continue
            
            col_series = df[col_nombre]
            
            if tipo_sub in ["gt", "lt", "ge", "le", "range", "eq"]: # Búsqueda numérica o de rango
                mascara_col_actual_para_numerico = pd.Series(False, index=df.index)
                
                # Iterar sobre cada celda de la columna para extraer números y unidades
                for idx, valor_celda_raw in col_series.items():
                    if pd.isna(valor_celda_raw) or str(valor_celda_raw).strip() == "":
                        continue # Saltar celdas vacías o nulas
                    
                    # Buscar todos los patrones número-unidad en la celda
                    for match_celda in self.patron_num_unidad_df.finditer(str(valor_celda_raw)):
                        try:
                            num_celda_val = self._parse_numero(match_celda.group(1))
                            if num_celda_val is None: continue # Si no se pudo parsear el número de la celda

                            unidad_celda_raw = match_celda.group(2) # Puede ser None
                            unidad_celda_canon: Optional[str] = None
                            if unidad_celda_raw and unidad_celda_raw.strip():
                                unidad_celda_canon = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_celda_raw.strip())

                            # Lógica de coincidencia de unidades:
                            # 1. Si la búsqueda no especifica unidad, cualquier unidad en la celda es válida (o sin unidad).
                            # 2. Si la búsqueda especifica unidad, la unidad de la celda (canonizada) debe coincidir.
                            # 3. Fallback: Si la unidad de la celda no se mapeó pero su forma normalizada simple coincide con la de búsqueda.
                            unidad_coincide = False
                            if unidad_sub_requerida_canon is None: 
                                unidad_coincide = True
                            elif unidad_celda_canon is not None and unidad_celda_canon == unidad_sub_requerida_canon:
                                unidad_coincide = True
                            elif unidad_celda_raw and unidad_sub_requerida_canon and \
                                 self.extractor_magnitud._normalizar_texto(unidad_celda_raw.strip()) == unidad_sub_requerida_canon:
                                # Este fallback es si la unidad de la celda no está en el mapeo pero coincide textualmente (normalizada)
                                unidad_coincide = True
                            
                            if not unidad_coincide: continue # Si las unidades no coinciden, probar el siguiente número/unidad en la celda

                            # Aplicar la condición numérica
                            cond_ok = False
                            if tipo_sub == "eq" and num_celda_val == valor_sub: cond_ok = True
                            elif tipo_sub == "gt" and num_celda_val > valor_sub: cond_ok = True
                            elif tipo_sub == "lt" and num_celda_val < valor_sub: cond_ok = True
                            elif tipo_sub == "ge" and num_celda_val >= valor_sub: cond_ok = True
                            elif tipo_sub == "le" and num_celda_val <= valor_sub: cond_ok = True
                            elif tipo_sub == "range" and valor_sub[0] <= num_celda_val <= valor_sub[1]: cond_ok = True # valor_sub es [min, max]
                            
                            if cond_ok:
                                mascara_col_actual_para_numerico.at[idx] = True
                                break # Un match numérico en la celda es suficiente para esta celda y columna
                        except ValueError: # Error al parsear número de la celda
                            logger.debug(f"Error parseando valor numérico en celda '{valor_celda_raw}' de columna '{col_nombre}'.")
                            continue # Probar el siguiente match en la celda
                    # Si ya se encontró una coincidencia en esta celda para esta columna, no seguir con otras celdas de la misma columna.
                    # Esto es incorrecto, debe seguir iterando las celdas de la columna. El break anterior era para los matches DENTRO de una celda.
                mascara_total_subtermino |= mascara_col_actual_para_numerico # Acumular (OR) resultados de esta columna
            
            elif tipo_sub == "str": # Búsqueda de texto
                try:
                    # valor_sub ya está normalizado por _analizar_terminos
                    valor_normalizado_busqueda = str(valor_sub) 
                    
                    serie_normalizada_df = col_series.astype(str).map(
                        self._normalizar_para_busqueda # Normalizar columna del DF para la comparación
                    )
                    # Usar regex para buscar palabra completa normalizada (whole word match)
                    patron_regex = r"\b" + re.escape(valor_normalizado_busqueda) + r"\b"
                    mascara_col_actual = serie_normalizada_df.str.contains(
                        patron_regex, regex=True, na=False # na=False trata NaN como no coincidencia
                    )
                    mascara_total_subtermino |= mascara_col_actual # Acumular (OR) resultados de esta columna
                except Exception as e_conv_str:
                    logger.warning(f"No se pudo convertir/buscar string en columna '{col_nombre}' para término '{valor_sub}': {e_conv_str}")
        
        return mascara_total_subtermino

    def _aplicar_mascara_combinada_para_segmento_and(
        self,
        df: pd.DataFrame,
        cols_a_buscar: List[str],
        terminos_analizados_segmento: List[Dict[str, Any]],
    ) -> pd.Series:
        """Combina máscaras de términos individuales dentro de un segmento AND."""
        if df is None or df.empty or not cols_a_buscar:
            return pd.Series(False, index=df.index if df is not None else None) # No hay dónde buscar o está vacío

        if not terminos_analizados_segmento: # Si un segmento AND está vacío (ej. query "palabra1 + "), no debería coincidir con nada.
            return pd.Series(False, index=df.index) 

        mascara_final_segmento_and = pd.Series(True, index=df.index) # Empezar con todo True para la intersección AND

        for termino_individual_analizado in terminos_analizados_segmento:
            mascara_este_termino = self._generar_mascara_para_un_termino(
                df, cols_a_buscar, termino_individual_analizado
            )
            mascara_final_segmento_and &= mascara_este_termino # Intersección (AND lógico)
            if not mascara_final_segmento_and.any(): # Optimización: si en algún punto no hay coincidencias, el AND total será False
                break

        return mascara_final_segmento_and

    def _combinar_mascaras_de_segmentos_or(
        self, lista_mascaras_segmentos: List[pd.Series], df_index_ref: Optional[pd.Index] = None
    ) -> pd.Series:
        """Combina máscaras de diferentes segmentos OR."""
        if not lista_mascaras_segmentos:
            if df_index_ref is not None: # Si se provee un índice de referencia y no hay máscaras
                return pd.Series(False, index=df_index_ref) # Ningún segmento OR coincidió
            else: # No se puede crear una Series sin índice
                logger.warning("Llamada a _combinar_mascaras_de_segmentos_or con lista vacía y sin df_index_ref.")
                return pd.Series(dtype=bool) # Devolver Series vacía de tipo bool

        # Se asume que todas las máscaras en lista_mascaras_segmentos tienen el mismo índice (el del DF original).
        # Si el índice de referencia no se pasa, tomar el de la primera máscara.
        idx_a_usar = df_index_ref if df_index_ref is not None else lista_mascaras_segmentos[0].index
        mascara_final_or = pd.Series(False, index=idx_a_usar) # Empezar con todo False para la unión OR

        for mascara_segmento in lista_mascaras_segmentos:
            if not mascara_segmento.index.equals(idx_a_usar):
                logger.warning("Índices de máscaras OR no coinciden. Puede llevar a resultados inesperados.")
                # Podría intentar realinear, pero es mejor asegurar que los índices son consistentes desde el origen.
                # Por ahora, se procede asumiendo que el usuario o la lógica previa lo maneja.
            mascara_final_or |= mascara_segmento # Unión (OR lógico)
        return mascara_final_or

    def _procesar_busqueda_en_df_objetivo(
        self,
        df_objetivo: pd.DataFrame,
        cols_objetivo: List[str],
        termino_busqueda_original: str,
    ) -> Tuple[pd.DataFrame, Optional[str]]:
        """Procesa la búsqueda completa (negación, OR, AND) en un DataFrame objetivo."""

        # 1. Aplicar negaciones y extraer la cadena de consulta para términos positivos
        df_despues_de_negaciones, terminos_positivos_query = self._aplicar_negaciones_y_extraer_positivos(
            df_objetivo,
            cols_objetivo, 
            termino_busqueda_original
        )
        
        # Si después de la negación el DF está vacío Y NO hay términos positivos, es un resultado válido (negación vació todo).
        if df_despues_de_negaciones.empty and not terminos_positivos_query.strip():
            logger.info("Búsqueda por negación resultó en un DataFrame vacío y no hay términos positivos restantes.")
            return df_despues_de_negaciones.copy(), None 
        # Si no hay términos positivos restantes (query era solo negaciones o la parte positiva se anuló),
        # la búsqueda termina con el resultado de las negaciones.
        if not terminos_positivos_query.strip():
            logger.info("La consulta solo contenía términos de negación válidos o la parte positiva quedó vacía. Devolviendo resultado de negaciones.")
            return df_despues_de_negaciones.copy(), None

        # 2. Procesar la parte positiva de la query (OR y AND)
        op_nivel1, segmentos_nivel1 = self._descomponer_nivel1_or(
            terminos_positivos_query # Usar la query ya procesada por la negación
        )

        if not segmentos_nivel1: 
            # Si la query positiva es inválida o vacía después del parseo OR (ej. " | | ")
            # y el df_despues_de_negaciones NO está vacío, significa que la parte positiva no filtró más.
            # No se debe devolver todo df_despues_de_negaciones si se esperaban más filtros positivos.
            # Si la query original era vacía y terminó aquí, es un caso especial.
            if not termino_busqueda_original.strip() and not terminos_positivos_query.strip():
                # Caso de búsqueda completamente vacía, que ya debería ser manejado por el método `buscar`.
                # Aquí, significa que la negación no hizo nada y la query positiva fue vacía.
                 return df_despues_de_negaciones.copy(), None 
            
            # Si había una query positiva pero se parseo a nada, devolver vacío.
            return ( 
                pd.DataFrame(columns=df_despues_de_negaciones.columns), # Devolver DF vacío con mismas columnas
                "Término de búsqueda positivo inválido o vacío tras parseo OR.",
            )

        lista_mascaras_para_or = []
        for seg_n1 in segmentos_nivel1: # Iterar sobre cada segmento OR
            op_nivel2, terminos_brutos_n2 = self._descomponer_nivel2_and(seg_n1) # Descomponer por AND
            terminos_atomicos_analizados = self._analizar_terminos(terminos_brutos_n2) # Analizar cada término AND

            if not terminos_atomicos_analizados: # Si un segmento AND no tiene términos válidos
                logger.debug(f"Segmento OR '{seg_n1}' no produjo términos analizables (AND). No contribuirá a la búsqueda OR.")
                # Un segmento OR que no tiene términos válidos (ej. query "palabra1 | + | palabra2") no debe añadir True a ninguna fila.
                # Se añade una máscara de todo False para este segmento OR.
                mascara_segmento_n1 = pd.Series(False, index=df_despues_de_negaciones.index)
            else:
                # Aplicar lógica AND a los términos de este segmento sobre el DF ya filtrado por negaciones
                mascara_segmento_n1 = self._aplicar_mascara_combinada_para_segmento_and(
                    df_despues_de_negaciones, 
                    cols_objetivo, 
                    terminos_atomicos_analizados
                )
            lista_mascaras_para_or.append(mascara_segmento_n1)

        if not lista_mascaras_para_or: # No debería ocurrir si segmentos_nivel1 no estaba vacío y se manejó arriba.
            logger.error("Error inesperado: lista_mascaras_para_or vacía a pesar de tener segmentos_nivel1.")
            return (
                pd.DataFrame(columns=df_despues_de_negaciones.columns),
                "Error interno: no se generaron máscaras para segmentos OR.",
            )

        # Combinar las máscaras de los segmentos OR
        # Si df_despues_de_negaciones está vacío, el índice estará vacío, y la máscara final también.
        mascara_final_df_objetivo = self._combinar_mascaras_de_segmentos_or(
            lista_mascaras_para_or, df_despues_de_negaciones.index
        )
        
        # Aplicar la máscara final al DataFrame que ya pasó el filtro de negación
        return df_despues_de_negaciones[mascara_final_df_objetivo].copy(), None

    def buscar(
        self,
        termino_busqueda_original: str,
        buscar_via_diccionario_flag: bool,
    ) -> Tuple[
        Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[str]
    ]:
        logger.info(
            f"Motor.buscar: termino='{termino_busqueda_original}', via_dicc={buscar_via_diccionario_flag}"
        )

        # DataFrame vacío de referencia para descripciones, con columnas si están cargadas
        df_vacio_desc_con_cols = pd.DataFrame(
            columns=(
                self.datos_descripcion.columns
                if self.datos_descripcion is not None
                else []
            )
        )
        fcds_obtenidos: Optional[pd.DataFrame] = None # Para almacenar resultados del diccionario

        # CASO 1: BÚSQUEDA VACÍA (sin texto en la entrada de búsqueda)
        if not termino_busqueda_original.strip():
            if self.datos_descripcion is not None:
                # Una búsqueda vacía devuelve todas las descripciones. La negación no aplica.
                logger.info("Búsqueda vacía: devolviendo todas las descripciones.")
                return (
                    self.datos_descripcion.copy(),
                    OrigenResultados.DIRECTO_DESCRIPCION_VACIA, 
                    None, # No hay FCDs
                    None, # No hay mensaje de error
                )
            else: # No hay descripciones cargadas
                logger.warning("Búsqueda vacía pero descripciones no cargadas.")
                return (
                    df_vacio_desc_con_cols, # Devolver DF vacío con cols (si se conocen) o sin ellas
                    OrigenResultados.ERROR_CARGA_DESCRIPCION,
                    None,
                    "Descripciones no cargadas para búsqueda vacía.",
                )

        # CASO 2: BÚSQUEDA VÍA DICCIONARIO
        if buscar_via_diccionario_flag:
            if self.datos_diccionario is None:
                logger.error("Intento de búsqueda vía diccionario, pero diccionario no cargado.")
                return (None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, "Diccionario no cargado.")
            
            cols_dic, err_cols_dic = self._obtener_nombres_columnas_busqueda_df(
                self.datos_diccionario, self.indices_columnas_busqueda_dic, "diccionario (búsqueda)")
            if not cols_dic: # Si no se pudieron obtener columnas válidas para buscar
                logger.error(f"Error obteniendo columnas para búsqueda en diccionario: {err_cols_dic}")
                return (None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, err_cols_dic)

            try:
                # Procesar búsqueda (incluyendo negaciones, OR, AND) en el DICCIONARIO
                fcds_obtenidos, error_procesamiento_dic = self._procesar_busqueda_en_df_objetivo(
                    self.datos_diccionario, cols_dic, termino_busqueda_original
                )
                if error_procesamiento_dic: 
                    logger.warning(f"Error/Término inválido al procesar búsqueda en diccionario: {error_procesamiento_dic}")
                    return (None, OrigenResultados.TERMINO_INVALIDO, None, error_procesamiento_dic)
                logger.info(f"FCDs encontrados tras búsqueda en diccionario: {len(fcds_obtenidos) if fcds_obtenidos is not None else 0}")

            except Exception as e_motor_dic:
                logger.exception("Error crítico en motor al buscar en diccionario.")
                return (None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, f"Error interno buscando en diccionario: {e_motor_dic}")

            # Continuar con la lógica de usar FCDs para buscar en DESCRIPCIONES
            if self.datos_descripcion is None:
                logger.error("Diccionario procesado, pero descripciones no cargadas para paso final.")
                return (None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos, "Descripciones no cargadas.")

            if fcds_obtenidos is None or fcds_obtenidos.empty:
                # No se encontraron FCDs, o la búsqueda en diccionario (con negaciones) resultó vacía.
                # El DataFrame de descripciones devuelto debe estar vacío.
                logger.info("No se encontraron FCDs en el diccionario, o la búsqueda resultó vacía.")
                return (df_vacio_desc_con_cols, OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, fcds_obtenidos, None)

            # Extraer términos clave de las columnas especificadas de los FCDs encontrados
            # Usar las mismas `cols_dic` que se usaron para la búsqueda inicial en el diccionario.
            terminos_extraidos_de_fcds = self._extraer_terminos_diccionario(fcds_obtenidos, cols_dic)
            if not terminos_extraidos_de_fcds:
                logger.info("Se encontraron FCDs, pero no se extrajeron términos válidos de ellos para buscar en descripciones.")
                return (df_vacio_desc_con_cols, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, fcds_obtenidos, None)

            try:
                # Construir una query OR con los términos extraídos. Encerrar en comillas para tratar frases.
                termino_or_para_desc = " | ".join(f'"{t}"' for t in terminos_extraidos_de_fcds) 
                logger.debug(f"Query para descripciones basada en FCDs: {termino_or_para_desc}")

                # Para la búsqueda en descripciones a partir de FCDs, buscar en todas las columnas de texto/objeto de descripciones.
                cols_desc_para_fcds, err_cols_desc_fcd = self._obtener_nombres_columnas_busqueda_df(
                    self.datos_descripcion, [], "descripciones (vía FCDs)") # [] o [-1] para todas texto/obj
                if not cols_desc_para_fcds:
                    logger.error(f"Error obteniendo columnas para búsqueda en descripciones (vía FCDs): {err_cols_desc_fcd}")
                    return (None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos, err_cols_desc_fcd)

                # La búsqueda en descripción a partir de FCDs NO debe aplicar la negación original.
                # _procesar_busqueda_en_df_objetivo aplicará negación si termino_or_para_desc la tuviera,
                # pero como los términos extraídos son palabras/frases, no deberían tener '#'.
                resultados_desc_via_dic, error_proc_desc_fcd = self._procesar_busqueda_en_df_objetivo(
                    self.datos_descripcion, cols_desc_para_fcds, termino_or_para_desc
                )
                if error_proc_desc_fcd: # Error durante el parseo o lógica interna de la búsqueda en descripciones
                    logger.warning(f"Error/Término inválido al buscar términos de FCDs en descripciones: {error_proc_desc_fcd}")
                    return (None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos, error_proc_desc_fcd)
                
                if resultados_desc_via_dic is None or resultados_desc_via_dic.empty:
                    logger.info("Búsqueda vía diccionario: FCDs encontrados, pero no produjeron resultados en descripciones.")
                    return (df_vacio_desc_con_cols, OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, fcds_obtenidos, None)
                else:
                    logger.info(f"Búsqueda vía diccionario exitosa. {len(resultados_desc_via_dic)} resultados en descripciones.")
                    return (resultados_desc_via_dic, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos, None)

            except Exception as e_motor_desc_via_dic:
                logger.exception("Error crítico en motor al buscar términos de FCDs en descripciones.")
                return (None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos, f"Error interno procesando FCDs en descripciones: {e_motor_desc_via_dic}")

        # CASO 3: BÚSQUEDA DIRECTA EN DESCRIPCIONES
        else: # buscar_via_diccionario_flag es False
            if self.datos_descripcion is None:
                logger.error("Intento de búsqueda directa en descripciones, pero no están cargadas.")
                return (None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, "Descripciones no cargadas.")

            # Para búsqueda directa, usar todas las columnas de texto/objeto de las descripciones.
            cols_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(
                self.datos_descripcion, [], "descripciones (búsqueda directa)") 
            if not cols_desc_directo:
                logger.error(f"Error obteniendo columnas para búsqueda directa en descripciones: {err_cols_desc_directo}")
                return (None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, err_cols_desc_directo)

            try:
                # Procesar búsqueda (incluyendo negaciones, OR, AND) directamente en DESCRIPCIONES
                resultados_directos, error_proc_desc_directo = self._procesar_busqueda_en_df_objetivo(
                    self.datos_descripcion, cols_desc_directo, termino_busqueda_original
                )
                if error_proc_desc_directo:
                     logger.warning(f"Error/Término inválido al procesar búsqueda directa en descripciones: {error_proc_desc_directo}")
                     return (None, OrigenResultados.TERMINO_INVALIDO, None, error_proc_desc_directo)
                logger.info(f"Resultados directos en descripción: {len(resultados_directos) if resultados_directos is not None else 0}")

            except Exception as e_motor_desc_directa:
                logger.exception("Error crítico en motor al buscar directamente en descripciones.")
                return (None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, f"Error interno buscando en descripciones: {e_motor_desc_directa}")

            if resultados_directos is None or resultados_directos.empty:
                logger.info("Búsqueda directa en descripciones no produjo resultados.")
                return (df_vacio_desc_con_cols, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None)
            else:
                logger.info(f"Búsqueda directa en descripciones exitosa. {len(resultados_directos)} resultados.")
                return (resultados_directos, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None)


    def _extraer_terminos_diccionario(
        self, df_coincidencias: pd.DataFrame, cols_nombres_dic: List[str]
    ) -> Set[str]:
        """Extrae y normaliza términos significativos de las columnas especificadas del DataFrame de coincidencias del diccionario."""
        terminos_clave: Set[str] = set()
        if df_coincidencias is None or df_coincidencias.empty or not cols_nombres_dic:
            logger.debug("No hay coincidencias en diccionario o no hay columnas especificadas para extraer términos.")
            return terminos_clave

        # Filtrar cols_nombres_dic para usar solo las que realmente existen en df_coincidencias
        columnas_validas_en_df = [c for c in cols_nombres_dic if c in df_coincidencias.columns]
        if not columnas_validas_en_df:
            logger.warning("Ninguna de las columnas del diccionario especificadas para extracción existe en el DataFrame de coincidencias.")
            return terminos_clave
        
        logger.debug(f"Extrayendo términos de las columnas del diccionario: {columnas_validas_en_df}")

        for col_nombre in columnas_validas_en_df:
            try:
                # Para cada celda en la columna de coincidencias
                for texto_celda in df_coincidencias[col_nombre].dropna().astype(str):
                    # Normalizar el texto de la celda ANTES de splitear, usando la normalización general de búsqueda
                    texto_celda_norm = self._normalizar_para_busqueda(texto_celda)
                    
                    # Dividir por espacios. Se podrían aplicar filtros más sofisticados aquí (stopwords, etc.) si es necesario.
                    # Por ahora, se toman palabras alfanuméricas de más de 2 caracteres.
                    palabras_celda = [
                        palabra for palabra in texto_celda_norm.split() 
                        if len(palabra) > 2 and palabra.isalnum() # Criterio simple para "palabra significativa"
                    ]
                    terminos_clave.update(palabras_celda)
            except Exception as e:
                logger.warning(f"Error extrayendo términos de la columna '{col_nombre}' del diccionario: {e}")

        logger.info(f"Se extrajeron {len(terminos_clave)} términos clave únicos del diccionario para búsqueda en descripciones.")
        if terminos_clave: logger.debug(f"Términos clave extraídos (muestra): {list(terminos_clave)[:15]}...")
        return terminos_clave


# --- Interfaz Gráfica (Sin cambios según la solicitud) ---
class InterfazGrafica(tk.Tk):
    CONFIG_FILE = "config_buscador_v0_7_4_mapeo_refactor.json" # Nombre de archivo de config actualizado

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado")
        self.geometry("1250x800")

        self.config = self._cargar_configuracion()
        indices_cfg = self.config.get("indices_columnas_busqueda_dic", [])
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg)

        self.resultados_actuales: Optional[pd.DataFrame] = None
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change)
        self.ultimo_termino_buscado: Optional[str] = None
        self.reglas_guardadas: List[Dict[str, Any]] = []

        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None

        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO
        self.color_fila_par = "white"
        self.color_fila_impar = "#f0f0f0"
        self.op_buttons: Dict[str, ttk.Button] = {}

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
        logger.info(f"Interfaz Gráfica (v{self.CONFIG_FILE.split('_')[3]}) inicializada.")


    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str):
        self._actualizar_estado_botones_operadores()

    def _cargar_configuracion(self) -> Dict:
        config = {}
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE}")
            except Exception as e:
                logger.error(f"Error al cargar config desde '{self.CONFIG_FILE}': {e}")
        else:
            logger.info(
                f"Archivo de configuración '{self.CONFIG_FILE}' no encontrado. Se usará config por defecto y se creará al cerrar."
            )

        # Asegurar que las rutas se manejan como Path y luego string si es necesario
        last_dic_path_str = config.get("last_dic_path")
        config["last_dic_path"] = (
            str(Path(last_dic_path_str)) if last_dic_path_str else None # Guardar como string
        )
        last_desc_path_str = config.get("last_desc_path")
        config["last_desc_path"] = (
            str(Path(last_desc_path_str)) if last_desc_path_str else None # Guardar como string
        )
        config.setdefault("indices_columnas_busqueda_dic", []) # Valor por defecto si no existe
        return config

    def _guardar_configuracion(self):
        self.config["last_dic_path"] = (
            str(self.motor.archivo_diccionario_actual) # Convertir Path a string para JSON
            if self.motor.archivo_diccionario_actual
            else None
        )
        self.config["last_desc_path"] = (
            str(self.motor.archivo_descripcion_actual) # Convertir Path a string para JSON
            if self.motor.archivo_descripcion_actual
            else None
        )
        # Guardar la configuración actual de índices del motor
        self.config["indices_columnas_busqueda_dic"] = (
            self.motor.indices_columnas_busqueda_dic
        )
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4) # indent=4 para legibilidad
            logger.info(f"Configuración guardada en: {self.CONFIG_FILE}")
        except Exception as e:
            logger.error(f"Error al guardar config en '{self.CONFIG_FILE}': {e}")
            messagebox.showerror( # Informar al usuario si falla el guardado
                "Error Configuración", f"No se pudo guardar la configuración:\n{e}"
            )

    def _configurar_estilo_ttk(self):
        style = ttk.Style(self)
        # Intenta usar un tema moderno si está disponible, con fallbacks
        themes = style.theme_names() # ('clam', 'alt', 'default', 'classic', 'vista', 'xpnative')
        os_name = platform.system()
        
        # Preferencias de temas por OS para una mejor apariencia nativa
        prefs = {
            "Windows": ["vista", "xpnative", "clam"],
            "Darwin": ["aqua", "clam"], # 'aqua' es el tema nativo de macOS
            "Linux": ["clam", "alt", "default"],
        }
        theme_to_use = next( # Encuentra el primer tema preferido que esté disponible
            (t for t in prefs.get(os_name, ["clam", "default"]) if t in themes), None
        )
        
        if not theme_to_use: # Si ninguna preferencia está disponible, usar el tema actual o un fallback
            theme_to_use = style.theme_use() if style.theme_use() else ("default" if "default" in themes else (themes[0] if themes else None))
        
        if theme_to_use:
            logger.info(f"Aplicando tema TTK: {theme_to_use}")
            try:
                style.theme_use(theme_to_use)
                # Configuración específica para botones de operadores si es necesario
                style.configure("Operator.TButton", padding=(2, 1), font=("TkDefaultFont", 9))
            except tk.TclError as e: # Por si el tema elegido causa problemas
                logger.warning(f"No se pudo aplicar el tema TTK '{theme_to_use}': {e}. Se usará el tema por defecto del sistema.")
                # Podríamos intentar con 'default' como fallback final aquí si falla el preferido
        else:
            logger.warning("No se encontró ningún tema TTK disponible. La apariencia puede ser básica.")

    def _crear_widgets(self):
        # Marco principal para los controles de carga y búsqueda
        self.marco_controles = ttk.LabelFrame(self, text="Controles")

        # Botones y etiquetas para cargar archivos
        self.btn_cargar_diccionario = ttk.Button(
            self.marco_controles,
            text="Cargar Diccionario",
            command=self._cargar_diccionario,
        )
        self.lbl_dic_cargado = ttk.Label(
            self.marco_controles,
            text="Dic: Ninguno",
            width=20, # Ancho fijo para consistencia
            anchor=tk.W, # Alinear texto a la izquierda
            relief=tk.SUNKEN, # Borde para dar apariencia de "display"
            borderwidth=1,
        )
        self.btn_cargar_descripciones = ttk.Button(
            self.marco_controles,
            text="Cargar Descripciones",
            command=self._cargar_excel_descripcion,
        )
        self.lbl_desc_cargado = ttk.Label(
            self.marco_controles,
            text="Desc: Ninguno",
            width=20,
            anchor=tk.W,
            relief=tk.SUNKEN,
            borderwidth=1,
        )

        # Frame para los botones de operadores de búsqueda
        self.frame_ops = ttk.Frame(self.marco_controles)
        op_buttons_defs = [ # Texto del botón, valor a insertar
            ("+", "+"), ("|", "|"), ("#", "#"), 
            (">", ">"), ("<", "<"), ("≥", ">="), ("≤", "<="), 
            ("=", "="), # Añadido botón para igualdad explícita
            ("-", "-") # Para rangos
        ]
        for i, (text, op_val) in enumerate(op_buttons_defs):
            btn = ttk.Button(
                self.frame_ops,
                text=text,
                command=lambda op=op_val: self._insertar_operador_validado(op),
                style="Operator.TButton", # Estilo específico para estos botones
                width=3, # Ancho pequeño
            )
            btn.grid(row=0, column=i, padx=1, pady=1, sticky="nsew")
            self.op_buttons[op_val] = btn # Guardar referencia al botón

        # Entrada de búsqueda y botones de acción
        self.entrada_busqueda = ttk.Entry(
            self.marco_controles, width=60, textvariable=self.texto_busqueda_var # Ancho mayor
        )
        self.btn_buscar = ttk.Button(
            self.marco_controles, text="Buscar", command=self._ejecutar_busqueda
        )
        self.btn_salvar_regla = ttk.Button( # Funcionalidad para guardar reglas/resultados (si se implementa)
            self.marco_controles, text="Salvar Regla", command=self._salvar_regla_actual, state="disabled"
        )
        self.btn_ayuda = ttk.Button( # Botón de ayuda
            self.marco_controles, text="?", command=self._mostrar_ayuda, width=3
        )
        self.btn_exportar = ttk.Button( # Botón para exportar resultados
            self.marco_controles, text="Exportar", command=self._exportar_resultados, state="disabled"
        )

        # Etiquetas para las tablas de Treeview
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario:")
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados / Descripciones:")

        # Treeview y scrollbars para el diccionario
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(
            self.frame_tabla_diccionario, show="headings", height=8 # Altura fija en filas
        )
        self.scrolly_diccionario = ttk.Scrollbar(
            self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview
        )
        self.scrollx_diccionario = ttk.Scrollbar(
            self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview
        )
        self.tabla_diccionario.configure(
            yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set
        )

        # Treeview y scrollbars para los resultados
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings") # Altura se ajustará por peso
        self.scrolly_resultados = ttk.Scrollbar(
            self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview
        )
        self.scrollx_resultados = ttk.Scrollbar(
            self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview
        )
        self.tabla_resultados.configure(
            yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set
        )

        # Barra de estado en la parte inferior
        self.barra_estado = ttk.Label(self, text="Listo.", relief=tk.SUNKEN, anchor=tk.W, borderwidth=1)
        
        self._actualizar_etiquetas_archivos() # Inicializar etiquetas de archivos cargados

    def _configurar_grid(self):
        # Configuración de expansión de filas y columnas principales de la ventana
        self.grid_rowconfigure(2, weight=1) # Frame tabla diccionario
        self.grid_rowconfigure(4, weight=3) # Frame tabla resultados (más peso para que sea más grande)
        self.grid_columnconfigure(0, weight=1) # Columna única principal se expande

        # Posicionamiento del marco de controles
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        # Configurar expansión de columnas dentro del marco de controles para las etiquetas de archivo
        self.marco_controles.grid_columnconfigure(1, weight=1) 
        self.marco_controles.grid_columnconfigure(3, weight=1)

        # Controles de carga de archivos
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5,0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2,10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5,0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2,5), pady=5, sticky="ew")

        # Frame de operadores
        self.frame_ops.grid(row=1, column=0, columnspan=5, padx=5, pady=(5,0), sticky="ew") # columnspan ajustado
        num_ops_buttons = len(self.op_buttons)
        for i in range(num_ops_buttons):
            self.frame_ops.grid_columnconfigure(i, weight=1) # Distribuir espacio entre botones de operadores

        # Controles de búsqueda
        self.entrada_busqueda.grid(row=2, column=0, columnspan=2, padx=5, pady=(0,5), sticky="ew")
        self.btn_buscar.grid(row=2, column=2, padx=(2,0), pady=(0,5), sticky="w")
        self.btn_salvar_regla.grid(row=2, column=3, padx=(2,0), pady=(0,5), sticky="w") # Posición ajustada
        self.btn_ayuda.grid(row=2, column=4, padx=(2,0), pady=(0,5), sticky="w") # Posición ajustada
        self.btn_exportar.grid(row=2, column=5, padx=(10,5), pady=(0,5), sticky="e") # Nueva columna para exportar

        # Tabla Diccionario
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10,0)) # Encima de su frame
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1)
        self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew")
        self.scrolly_diccionario.grid(row=0, column=1, sticky="ns")
        self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")

        # Tabla Resultados
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0,0)) # Encima de su frame
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.scrolly_resultados.grid(row=0, column=1, sticky="ns")
        self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # Barra de estado
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5,0))

    def _configurar_eventos(self):
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())
        self.protocol("WM_DELETE_WINDOW", self.on_closing) # Manejar cierre de ventana

    def _actualizar_estado(self, mensaje: str):
        self.barra_estado.config(text=mensaje)
        logger.info(f"Estado UI: {mensaje}")
        self.update_idletasks() # Forzar actualización de la UI

    def _mostrar_ayuda(self):
        ayuda = """Sintaxis de Búsqueda:
-------------------------------------
- Texto simple: Busca la palabra o frase (sensible a mayúsculas/minúsculas y acentos según normalización interna). Ej: `router cisco`
- Operadores Lógicos:
  * `término1 + término2`: Busca filas con AMBOS (AND). Ej: `tarjeta + 16 puertos`
  * `término1 | término2` (o ` / `): Busca filas con AL MENOS UNO (OR). Ej: `modulo | SFP`
- Comparaciones numéricas (unidad opcional, si se usa debe coincidir con mapeo de 1ª col. diccionario):
  * `> num[UNIDAD]`: Mayor. Ej: ` > 1000` o ` > 1000W`
  * `< num[UNIDAD]`: Menor. Ej: ` < 50` o ` < 50V`
  * `>= num[UNIDAD]` o ` ≥ num[UNIDAD]`: Mayor o igual. Ej: ` >= 48A`
  * `<= num[UNIDAD]` o ` ≤ num[UNIDAD]`: Menor o igual. Ej: ` <= 10.5W`
  * `= num[UNIDAD]`: Igual a. Ej: ` = 24V`
- Rangos numéricos (unidad opcional, ambos extremos incluidos):
  * `num1 - num2[UNIDAD]`: Entre num1 y num2. Ej: `10 - 20` o `50 - 100V`
- Negación (excluir):
  * `#palabra`: Excluye filas que contengan `palabra`. Ej: `switch + #gestionable`
  * `# "frase completa"`: Excluye filas con la frase. Ej: `fuente + #"bajo rendimiento"`

Modo de Búsqueda:
1. La búsqueda (incluyendo negaciones y operadores) se aplica primero al Diccionario (si está cargado).
2. Si hay coincidencias en Diccionario (FCDs):
   - Se extraen términos clave de los FCDs.
   - Estos términos se buscan (con OR) en las Descripciones. La negación original NO se propaga aquí.
3. Si NO hay FCDs, o si los FCDs no llevan a resultados en Descripciones:
   - Se preguntará si desea buscar el término original directamente en las Descripciones.
   - Esta búsqueda directa SÍ aplicará la negación y operadores originales.
4. Búsqueda vacía (sin texto): Muestra todas las descripciones.
5. Espaciado: Se recomienda espaciar los operadores (`+`, `|`, `/`, `>`, `<`, `>=`, `<=`, `=`, `-`) para claridad, aunque el parseador intenta ser flexible. La negación `#` debe ir pegada a la palabra o frase entre comillas.
   Ej: `palabra1 + palabra2`, `> 100 W`, `10 - 20 KG`, `#palabraExcluida`, `#"frase a excluir"`
"""
        messagebox.showinfo("Ayuda - Sintaxis de Búsqueda", ayuda)

    def _configurar_tags_treeview(self):
        # Configurar colores alternos para filas en ambas tablas
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure("par", background=self.color_fila_par)
            tabla.tag_configure("impar", background=self.color_fila_impar)

    def _configurar_orden_tabla(self, tabla: ttk.Treeview):
        # Asignar comando de ordenación a las cabeceras de columna
        cols = tabla["columns"]
        if cols: # Solo si hay columnas definidas
            for col in cols:
                tabla.heading(
                    col,
                    text=str(col), # Texto de la cabecera
                    anchor=tk.W,   # Alinear a la izquierda
                    command=lambda c=col, t=tabla: self._ordenar_columna(t, c, False), # False para orden ascendente inicial
                )

    def _ordenar_columna(self, tabla: ttk.Treeview, col: str, reverse: bool):
        # Determinar qué DataFrame se está ordenando
        df_para_ordenar = None
        if tabla == self.tabla_diccionario:
            # Ordenar directamente self.motor.datos_diccionario si se desea que la vista previa se ordene permanentemente
            # O una copia si solo es para la vista actual. Asumimos que se ordena la fuente.
            df_para_ordenar = self.motor.datos_diccionario 
        elif tabla == self.tabla_resultados:
            df_para_ordenar = self.resultados_actuales

        if df_para_ordenar is None or df_para_ordenar.empty or col not in df_para_ordenar.columns:
            logger.debug(f"No se puede ordenar la tabla por columna '{col}'. DataFrame no disponible, vacío o columna inexistente.")
            # Re-asignar el comando para permitir cambiar la dirección en el próximo clic
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, not reverse))
            return

        logger.info(f"Ordenando tabla por columna '{col}', descendente={reverse}")
        try:
            # Intentar ordenación numérica si es posible, sino textual
            # Guardar una copia antes de modificar el df original si es necesario.
            # Aquí, se modifica el df referenciado.
            
            # Convertir a numérico si es posible para una mejor ordenación numérica
            # Si no, se ordenará como string.
            # El manejo de NaNs es importante para la ordenación.
            df_col_numeric = pd.to_numeric(df_para_ordenar[col], errors='coerce')
            
            if not df_col_numeric.isna().all(): # Si hay al menos un valor numérico
                df_ordenado = df_para_ordenar.sort_values(by=col, ascending=not reverse, na_position='last', key=lambda x: pd.to_numeric(x, errors='coerce'))
            else: # Tratar como texto si no se puede convertir a numérico
                df_ordenado = df_para_ordenar.sort_values(by=col, ascending=not reverse, na_position='last', key=lambda x: x.astype(str).str.lower())

            # Actualizar el DataFrame en el motor o en los resultados actuales
            if tabla == self.tabla_diccionario:
                self.motor.datos_diccionario = df_ordenado
                self._actualizar_tabla(tabla, df_ordenado, limite_filas=100) # Límite para diccionario
            elif tabla == self.tabla_resultados:
                self.resultados_actuales = df_ordenado
                self._actualizar_tabla(tabla, df_ordenado) # Sin límite para resultados

            # Actualizar el comando de la cabecera para invertir el orden en el siguiente clic
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, not reverse))
            self._actualizar_estado(f"Tabla ordenada por '{col}' ({'Ascendente' if not reverse else 'Descendente'}).")

        except Exception as e:
            logging.exception(f"Error al intentar ordenar por columna '{col}'")
            messagebox.showerror("Error al Ordenar", f"No se pudo ordenar por '{col}':\n{e}")
            # Restaurar el comando de ordenación original en caso de error
            tabla.heading(col, command=lambda c=col, t=tabla: self._ordenar_columna(t, c, False))


    def _actualizar_tabla(
        self,
        tabla: ttk.Treeview,
        datos: Optional[pd.DataFrame],
        limite_filas: Optional[int] = None, # Límite de filas a mostrar (útil para previews grandes)
        columnas_a_mostrar: Optional[List[str]] = None, # Para mostrar solo un subconjunto de columnas
    ):
        is_diccionario = tabla == self.tabla_diccionario
        tabla_nombre_log = "Diccionario" if is_diccionario else "Resultados"
        logger.debug(f"Actualizando tabla {tabla_nombre_log}.")

        # Limpiar tabla anterior
        try:
            for i in tabla.get_children():
                tabla.delete(i)
        except tk.TclError as e: # Puede ocurrir si la tabla se destruye/recrea
            logger.warning(f"Error Tcl al limpiar tabla {tabla_nombre_log}: {e}. Puede ser ignorable si la tabla está siendo refrescada.")
        tabla["columns"] = () # Resetear columnas

        if datos is None or datos.empty:
            logger.debug(f"No hay datos para mostrar en la tabla {tabla_nombre_log}.")
            self._configurar_orden_tabla(tabla) # Reconfigurar cabeceras vacías para ordenación futura
            return

        df_para_mostrar = datos
        
        # Determinar qué columnas mostrar
        cols_df_originales = list(df_para_mostrar.columns)
        cols_finales_a_usar = cols_df_originales
        if columnas_a_mostrar: # Si se especifica un subconjunto
            cols_finales_a_usar = [c for c in columnas_a_mostrar if c in cols_df_originales]
            if not cols_finales_a_usar: # Si las columnas especificadas no existen, mostrar todas como fallback
                logger.warning(f"Columnas especificadas para tabla {tabla_nombre_log} no encontradas. Mostrando todas.")
                cols_finales_a_usar = cols_df_originales
        
        if not cols_finales_a_usar:
            logger.warning(f"DataFrame para tabla {tabla_nombre_log} no tiene columnas para mostrar.")
            self._configurar_orden_tabla(tabla)
            return

        tabla["columns"] = tuple(cols_finales_a_usar)

        # Configurar cabeceras y anchos de columna
        for col in cols_finales_a_usar:
            tabla.heading(col, text=str(col), anchor=tk.W)
            try:
                # Intentar calcular un ancho basado en el contenido y la cabecera
                col_as_str = df_para_mostrar[col].astype(str)
                ancho_contenido = col_as_str.str.len().max() if not col_as_str.empty else 0
                ancho_cabecera = len(str(col))
                # Ajustar factores para el ancho (empírico)
                ancho = max(70, min(int(max(ancho_cabecera * 8, ancho_contenido * 6.5) + 25), 400)) 
                tabla.column(col, anchor=tk.W, width=ancho, minwidth=70) # minwidth para evitar columnas muy estrechas
            except Exception: # Fallback a un ancho fijo si hay problemas
                tabla.column(col, anchor=tk.W, width=100, minwidth=50)

        # Limitar filas si se especifica (para rendimiento con DFs grandes en vista previa)
        df_iterar = df_para_mostrar[cols_finales_a_usar] # Usar solo las columnas seleccionadas
        if limite_filas is not None and len(df_iterar) > limite_filas:
            df_iterar = df_iterar.head(limite_filas)
            logger.debug(f"Mostrando primeras {limite_filas} filas en tabla {tabla_nombre_log}.")

        # Insertar filas
        for i, (_, row) in enumerate(df_iterar.iterrows()):
            vals = [str(v) if pd.notna(v) else "" for v in row.values] # Convertir todo a string para Treeview
            tag = "par" if i % 2 == 0 else "impar" # Aplicar tag para color de fila alterno
            try:
                tabla.insert("", "end", values=vals, tags=(tag,))
            except tk.TclError as e_insert: # Manejar errores si algún valor no es compatible con Tcl/Tk
                logger.warning(f"Error Tcl al insertar fila {i} en tabla {tabla_nombre_log}: {e_insert}. Intentando con valores ASCII.")
                try: # Fallback a valores ASCII-ficados
                    vals_ascii = [v.encode("ascii", "ignore").decode("ascii", "ignore") for v in vals]
                    tabla.insert("", "end", values=vals_ascii, tags=(tag,))
                except Exception as e_inner: # Si el fallback también falla
                    logger.error(f"Fallo el fallback ASCII para fila {i} en tabla {tabla_nombre_log}: {e_inner}. Fila omitida.")
        
        self._configurar_orden_tabla(tabla) # Reaplicar configuración de ordenación a las nuevas columnas

    def _actualizar_etiquetas_archivos(self):
        dic_path = self.motor.archivo_diccionario_actual
        desc_path = self.motor.archivo_descripcion_actual
        
        dic_name = dic_path.name if dic_path else "Ninguno"
        desc_name = desc_path.name if desc_path else "Ninguno"

        max_len_label = 25 # Longitud máxima para el nombre del archivo en la etiqueta
        
        # Acortar nombres de archivo si son muy largos para la etiqueta
        dic_display = f"Dic: {dic_name}" if len(dic_name) <= max_len_label else f"Dic: ...{dic_name[-(max_len_label-4):]}"
        desc_display = f"Desc: {desc_name}" if len(desc_name) <= max_len_label else f"Desc: ...{desc_name[-(max_len_label-4):]}"

        self.lbl_dic_cargado.config(text=dic_display, foreground="green" if dic_path else "red")
        self.lbl_desc_cargado.config(text=desc_display, foreground="green" if desc_path else "red")

    def _actualizar_botones_estado_general(self):
        dic_cargado = self.motor.datos_diccionario is not None
        desc_cargado = self.motor.datos_descripcion is not None

        # Habilitar/deshabilitar botones de operadores basados en si hay datos cargados
        if dic_cargado or desc_cargado:
            self._actualizar_estado_botones_operadores() # Esto considerará el texto en la entrada
        else:
            self._deshabilitar_botones_operadores()

        # Botón de búsqueda principal
        self.btn_buscar["state"] = "normal" if dic_cargado and desc_cargado else "disabled"

        # Lógica para habilitar "Salvar Regla"
        # (Asume que una regla se puede salvar si hay un término buscado y algún tipo de resultado)
        puede_salvar_algo = False
        if self.ultimo_termino_buscado and self.origen_principal_resultados != OrigenResultados.NINGUNO:
            # Si fue vía diccionario y hubo FCDs O resultados finales de descripción
            if self.origen_principal_resultados.es_via_diccionario and \
               ((self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty) or \
                (self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and \
                 self.origen_principal_resultados == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC)):
                puede_salvar_algo = True
            # Si fue directo a descripción y hubo resultados (o era búsqueda vacía que muestra todo)
            elif (self.origen_principal_resultados.es_directo_descripcion or \
                  self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA) and \
                 self.desc_finales_de_ultima_busqueda is not None : # No necesariamente not empty para el caso de "todo"
                puede_salvar_algo = True
        
        self.btn_salvar_regla["state"] = "normal" if puede_salvar_algo else "disabled"
        
        # Botón de exportar (habilitado si hay reglas guardadas o resultados actuales)
        export_enabled = (self.resultados_actuales is not None and not self.resultados_actuales.empty)
        self.btn_exportar["state"] = "normal" if export_enabled else "disabled"


    def _cargar_diccionario(self):
        # Usar la ruta del último diccionario cargado como directorio inicial
        last_dir_str = self.config.get("last_dic_path")
        initial_dir = str(Path(last_dir_str).parent) if last_dir_str and Path(last_dir_str).exists() else os.getcwd()

        ruta_seleccionada = filedialog.askopenfilename(
            title="Seleccionar Archivo Diccionario",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            initialdir=initial_dir,
        )
        if not ruta_seleccionada: # Si el usuario cancela
            logger.info("Carga de diccionario cancelada por el usuario.")
            return

        nombre_archivo = Path(ruta_seleccionada).name
        self._actualizar_estado(f"Cargando diccionario: {nombre_archivo}...")
        
        # Limpiar vistas y datos anteriores relacionados con el diccionario
        self._actualizar_tabla(self.tabla_diccionario, None)
        # Si se cambia el diccionario, los resultados anteriores pueden no ser válidos
        self._actualizar_tabla(self.tabla_resultados, None) 
        self.resultados_actuales = None
        self.fcds_de_ultima_busqueda = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        exito_carga, msg_error_carga = self.motor.cargar_excel_diccionario(ruta_seleccionada)
        
        if exito_carga and self.motor.datos_diccionario is not None:
            self.config["last_dic_path"] = ruta_seleccionada # Guardar ruta para la próxima vez
            self._guardar_configuracion() # Guardar toda la config
            
            df_dic = self.motor.datos_diccionario
            num_filas = len(df_dic)
            
            # Obtener nombres de columnas de búsqueda para mostrar en la etiqueta de la tabla
            cols_busqueda_nombres, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                df_dic, self.motor.indices_columnas_busqueda_dic, "diccionario (preview)"
            )
            
            indices_str = "Todas Texto/Obj" # Por defecto si no hay config específica o es -1
            if self.motor.indices_columnas_busqueda_dic and self.motor.indices_columnas_busqueda_dic != [-1]:
                indices_str = ", ".join(map(str, self.motor.indices_columnas_busqueda_dic))

            lbl_text_dic_preview = f"Vista Previa Diccionario ({num_filas} filas)"
            if cols_busqueda_nombres: # Si se pudieron determinar columnas específicas
                lbl_text_dic_preview = f"Diccionario (Buscar en: {', '.join(cols_busqueda_nombres)} - Índices: {indices_str})"
            self.lbl_tabla_diccionario.config(text=lbl_text_dic_preview)

            self._actualizar_tabla(self.tabla_diccionario, df_dic, limite_filas=100, columnas_a_mostrar=cols_busqueda_nombres)
            self.title(f"Buscador - Dic: {nombre_archivo}") # Actualizar título de la ventana
            self._actualizar_estado(f"Diccionario '{nombre_archivo}' ({num_filas} filas) cargado.")
        else:
            self._actualizar_estado(f"Error al cargar diccionario: {msg_error_carga or 'Desconocido'}")
            if msg_error_carga: messagebox.showerror("Error Carga Diccionario", msg_error_carga)
            # Restaurar título por defecto si falla la carga
            current_desc_title = Path(self.motor.archivo_descripcion_actual).name if self.motor.archivo_descripcion_actual else "N/A"
            self.title(f"Buscador - Dic: N/A | Desc: {current_desc_title}")

        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()


    def _cargar_excel_descripcion(self):
        last_dir_str = self.config.get("last_desc_path")
        initial_dir = str(Path(last_dir_str).parent) if last_dir_str and Path(last_dir_str).exists() else os.getcwd()

        ruta_seleccionada = filedialog.askopenfilename(
            title="Seleccionar Archivo de Descripciones",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            initialdir=initial_dir,
        )
        if not ruta_seleccionada:
            logger.info("Carga de descripciones cancelada por el usuario.")
            return

        nombre_archivo = Path(ruta_seleccionada).name
        self._actualizar_estado(f"Cargando descripciones: {nombre_archivo}...")
        
        # Limpiar vistas y datos anteriores de resultados/descripciones
        self.resultados_actuales = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO
        self._actualizar_tabla(self.tabla_resultados, None) # Limpiar la tabla de resultados

        exito_carga, msg_error_carga = self.motor.cargar_excel_descripcion(ruta_seleccionada)
        
        if exito_carga and self.motor.datos_descripcion is not None:
            self.config["last_desc_path"] = ruta_seleccionada
            self._guardar_configuracion()
            
            df_desc = self.motor.datos_descripcion
            num_filas = len(df_desc)
            self._actualizar_estado(f"Descripciones '{nombre_archivo}' ({num_filas} filas) cargadas. Mostrando vista previa...")
            # Mostrar todas las descripciones cargadas en la tabla de resultados como preview inicial
            self._actualizar_tabla(self.tabla_resultados, df_desc) 
            
            dic_n_title = Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A"
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: {nombre_archivo}")
        else:
            self._actualizar_estado(f"Error al cargar descripciones: {msg_error_carga or 'Desconocido'}")
            if msg_error_carga: messagebox.showerror("Error Carga Descripciones", msg_error_carga)
            dic_n_title = Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A"
            self.title(f"Buscador - Dic: {dic_n_title} | Desc: N/A")

        self._actualizar_etiquetas_archivos()
        self._actualizar_botones_estado_general()

    def _ejecutar_busqueda(self):
        # Asegurarse de que ambos archivos necesarios están cargados
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", "Por favor, cargue tanto el archivo de Diccionario como el de Descripciones antes de buscar.")
            return

        termino_busqueda_actual_ui = self.texto_busqueda_var.get() # Obtener término de la UI
        self.ultimo_termino_buscado = termino_busqueda_actual_ui # Guardar para posible "Salvar Regla"

        # Resetear resultados anteriores
        self.resultados_actuales = None
        self._actualizar_tabla(self.tabla_resultados, None) # Limpiar tabla de resultados
        self.fcds_de_ultima_busqueda = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO

        self._actualizar_estado(f"Buscando '{termino_busqueda_actual_ui}'...")

        # Ejecutar la búsqueda a través del motor (asumiendo búsqueda vía diccionario por defecto)
        res_df, origen_res, fcds_res, err_msg_motor = self.motor.buscar(
            termino_busqueda_original=termino_busqueda_actual_ui,
            buscar_via_diccionario_flag=True, # Iniciar siempre vía diccionario si está disponible
        )

        self.fcds_de_ultima_busqueda = fcds_res # Guardar FCDs, incluso si son None o vacíos
        self.origen_principal_resultados = origen_res # Guardar el origen

        df_desc_cols_ref = self.motor.datos_descripcion.columns if self.motor.datos_descripcion is not None else []

        # --- Manejo de resultados y errores ---
        if err_msg_motor and origen_res.es_error_operacional: # Errores críticos del motor
            messagebox.showerror("Error de Búsqueda (Motor)", f"Error interno del motor: {err_msg_motor}")
            self._actualizar_estado(f"Error en motor: {err_msg_motor}")
            self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref) # Mostrar tabla vacía
        elif origen_res.es_error_carga or origen_res.es_error_configuracion or origen_res.es_termino_invalido:
            msg_shown = err_msg_motor or f"Error impidió la operación: {origen_res.name}"
            messagebox.showerror("Error de Búsqueda", msg_shown)
            self._actualizar_estado(msg_shown)
            self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref)
        
        # --- Flujo normal de resultados ---
        elif origen_res == OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC:
            self.resultados_actuales = res_df
            num_fcds = len(fcds_res) if fcds_res is not None else 0
            num_res_df = len(res_df) if res_df is not None else 0
            self._actualizar_estado(f"'{termino_busqueda_actual_ui}': {num_fcds} en Dic., {num_res_df} en Desc.")
        
        elif origen_res in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS] or \
             (fcds_res is not None and fcds_res.empty and origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC):
            # Diccionario no produjo resultados finales en descripciones (o no hubo FCDs / términos válidos)
            self.resultados_actuales = res_df # res_df será un DF vacío con columnas de descripciones
            
            num_fcds_info = len(fcds_res) if fcds_res is not None else 0
            msg_info_fcd = f"{num_fcds_info} coincidencias en Diccionario"
            if fcds_res is not None and fcds_res.empty: msg_info_fcd = "Ninguna coincidencia en Diccionario"

            if origen_res == OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS:
                msg_desc_issue = "pero no se extrajeron términos válidos para buscar en descripciones."
                self._actualizar_estado(f"'{termino_busqueda_actual_ui}': {msg_info_fcd}, sin términos válidos para Desc.")
            else: # VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                msg_desc_issue = "lo que no produjo resultados en Descripciones."
                self._actualizar_estado(f"'{termino_busqueda_actual_ui}': {msg_info_fcd}, 0 en Desc.")

            # Preguntar al usuario si desea buscar directamente en descripciones
            if messagebox.askyesno("Búsqueda Alternativa",
                f"{msg_info_fcd} para '{termino_busqueda_actual_ui}', {msg_desc_issue}\n\n"
                f"¿Desea buscar '{termino_busqueda_actual_ui}' directamente en las Descripciones?"):
                
                self._actualizar_estado(f"Buscando directamente '{termino_busqueda_actual_ui}' en descripciones...")
                res_df_directo, origen_directo, _, err_msg_motor_directo = self.motor.buscar(
                    termino_busqueda_original=termino_busqueda_actual_ui, buscar_via_diccionario_flag=False
                )
                # Actualizar estado con los resultados de la búsqueda directa
                self.origen_principal_resultados = origen_directo
                self.fcds_de_ultima_busqueda = None # No hay FCDs en búsqueda directa

                if err_msg_motor_directo and origen_directo.es_error_operacional:
                    messagebox.showerror("Error Búsqueda Directa", f"Error interno: {err_msg_motor_directo}")
                    self._actualizar_estado(f"Error búsqueda directa: {err_msg_motor_directo}")
                    self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref)
                elif origen_directo.es_error_carga or origen_directo.es_error_configuracion or origen_directo.es_termino_invalido:
                    msg_shown_directo = err_msg_motor_directo or f"Error en búsqueda directa: {origen_directo.name}"
                    messagebox.showerror("Error Búsqueda Directa", msg_shown_directo)
                    self._actualizar_estado(msg_shown_directo)
                    self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref)
                else: # Búsqueda directa tuvo éxito (con o sin resultados)
                    self.resultados_actuales = res_df_directo
                    num_rdd = len(self.resultados_actuales) if self.resultados_actuales is not None else 0
                    self._actualizar_estado(f"Búsqueda directa '{termino_busqueda_actual_ui}': {num_rdd} resultados.")
                    if num_rdd == 0 and origen_directo == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and termino_busqueda_actual_ui.strip():
                        messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_busqueda_actual_ui}' en búsqueda directa.")
            # Si el usuario dice "No" a la búsqueda alternativa, self.resultados_actuales ya es el DF vacío de la vía diccionario.
        
        elif origen_res == OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS: # Búsqueda directa (inicial) tuvo resultados
            self.resultados_actuales = res_df
            self._actualizar_estado(f"Búsqueda directa '{termino_busqueda_actual_ui}': {len(res_df) if res_df is not None else 0} resultados.")
        
        elif origen_res == OrigenResultados.DIRECTO_DESCRIPCION_VACIA: # Búsqueda directa (inicial) o búsqueda vacía
            self.resultados_actuales = res_df # Puede ser todas las descripciones o un DF vacío
            num_res = len(res_df) if res_df is not None else 0
            if not termino_busqueda_actual_ui.strip(): # Si la búsqueda original fue vacía
                 self._actualizar_estado(f"Mostrando todas las descripciones ({num_res} filas).")
            else: # Búsqueda directa no vacía pero sin resultados
                 self._actualizar_estado(f"Búsqueda directa '{termino_busqueda_actual_ui}': 0 resultados.")
                 messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_busqueda_actual_ui}' en búsqueda directa.")

        # Asegurar que self.resultados_actuales siempre sea un DataFrame para la tabla
        if self.resultados_actuales is None:
            self.resultados_actuales = pd.DataFrame(columns=df_desc_cols_ref)
        
        self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy() # Guardar copia para "Salvar Regla"

        self._actualizar_tabla(self.tabla_resultados, self.resultados_actuales)
        self._actualizar_botones_estado_general() # Actualizar estado de botones (ej. exportar)

        # Intentar enfocar en la vista previa del diccionario si hay FCDs
        if self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty:
             self._buscar_y_enfocar_en_preview()
        elif self.motor.datos_diccionario is not None and not self.motor.datos_diccionario.empty \
             and self.origen_principal_resultados != OrigenResultados.NINGUNO \
             and not self.origen_principal_resultados.es_via_diccionario: # Si la búsqueda fue directa, pero hay dicc. cargado
             self._buscar_y_enfocar_en_preview() # Intentar enfocar la query original en el dicc.


    def _buscar_y_enfocar_en_preview(self):
        df_completo_dic = self.motor.datos_diccionario
        if df_completo_dic is None or df_completo_dic.empty:
            return

        termino_buscar_raw = self.texto_busqueda_var.get()
        if not termino_buscar_raw.strip(): # No enfocar si la búsqueda es vacía
            return

        # Intentar obtener el primer término "positivo" de la búsqueda para enfocar
        # (simplificado: tomar el primer segmento OR, luego el primer término AND)
        # La negación no se usa para enfocar, se busca el término positivo.
        
        _, terminos_positivos_query = self.motor._aplicar_negaciones_y_extraer_positivos(
            pd.DataFrame(), [], termino_buscar_raw # DF y cols no importan, solo queremos los términos positivos
        )
        if not terminos_positivos_query.strip():
            # Si solo había negaciones o la query positiva es vacía, no hay nada que enfocar.
            # O podríamos intentar enfocar el *primer término negado* (sin el '#') si es el único.
            # Por simplicidad, si no hay positivos, no enfocamos.
            return

        op_n1, segmentos_n1 = self.motor._descomponer_nivel1_or(terminos_positivos_query)
        if not segmentos_n1: return

        primer_segmento_n1 = segmentos_n1[0]
        op_n2, terminos_brutos_n2 = self.motor._descomponer_nivel2_and(primer_segmento_n1)
        if not terminos_brutos_n2: return

        termino_enfocar_bruto_final = terminos_brutos_n2[0] # Tomar el primer término "atómico"
        
        # Normalizar el término a enfocar de la misma manera que se normaliza para la búsqueda textual
        termino_enfocar_normalizado = self.motor._normalizar_para_busqueda(termino_enfocar_bruto_final)
        if not termino_enfocar_normalizado: return # No enfocar si el término normalizado es vacío

        items_preview_ids = self.tabla_diccionario.get_children("") # Obtener todos los items (filas) del Treeview
        if not items_preview_ids: return

        logger.info(f"Intentando enfocar '{termino_enfocar_normalizado}' (original: '{termino_enfocar_bruto_final}') en vista previa del diccionario...")
        found_item_id = None

        # Columnas visibles en la preview del diccionario (según cómo se actualizó la tabla)
        columnas_preview_visibles_actuales = self.tabla_diccionario["columns"]
        
        # Iterar sobre los items del Treeview (que son un subconjunto limitado del diccionario)
        for item_id in items_preview_ids:
            try:
                # Obtener los valores mostrados en la fila del Treeview
                # Estos valores ya son strings.
                valores_fila_preview_mostrados_tupla = self.tabla_diccionario.item(item_id, "values")
                # Asegurar que es una lista o tupla de strings
                valores_fila_preview_mostrados = list(valores_fila_preview_mostrados_tupla) if isinstance(valores_fila_preview_mostrados_tupla, tuple) else []


                # Normalizar los valores de la fila del Treeview y verificar coincidencia
                if any(
                    termino_enfocar_normalizado in self.motor._normalizar_para_busqueda(str(val_prev))
                    for val_prev in valores_fila_preview_mostrados if val_prev is not None
                ):
                    found_item_id = item_id
                    break # Encontrado, salir del bucle de items
            except Exception as e_focus_tree: # Error al procesar un item del treeview
                logger.warning(f"Error procesando item {item_id} en preview (búsqueda treeview): {e_focus_tree}")
                continue

        if found_item_id:
            logger.info(f"Término '{termino_enfocar_normalizado}' enfocado en preview (item ID: {found_item_id}).")
            try:
                self.tabla_diccionario.selection_set(found_item_id) # Seleccionar
                self.tabla_diccionario.see(found_item_id) # Asegurar que es visible (scroll)
                self.tabla_diccionario.focus(found_item_id) # Poner foco
            except Exception as e_tk_focus: # Error en operaciones de Tkinter
                logger.error(f"Error al intentar enfocar item {found_item_id} en Treeview de preview: {e_tk_focus}")
        else:
            logger.info(f"Término '{termino_enfocar_normalizado}' no encontrado/enfocado en la vista previa actual del diccionario.")


    def _salvar_regla_actual(self):
        # Esta función es un placeholder o requiere una definición más clara de "regla"
        # Por ahora, asumimos que se refiere a salvar los resultados actuales de la búsqueda.
        origen_nombre_actual = self.origen_principal_resultados.name
        logger.info(f"Intentando salvar regla/resultados. Origen: {origen_nombre_actual}, Último término: '{self.ultimo_termino_buscado}'")

        if not self.ultimo_termino_buscado and not (self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and self.desc_finales_de_ultima_busqueda is not None) :
            messagebox.showerror("Error al Salvar", "No hay término de búsqueda o resultados válidos para salvar.")
            return

        # Implementación básica: Guardar el DataFrame de resultados actual
        # Se podría expandir para guardar la query, FCDs, etc. en un formato estructurado.
        
        df_a_salvar = None
        tipo_datos_salvados = "DESCONOCIDO"

        if self.origen_principal_resultados.es_via_diccionario:
            # Aquí se podría usar el diálogo _mostrar_dialogo_seleccion_salvado_via_diccionario
            # Por simplicidad, si hay desc_finales, se salvan esos. Si no, los FCDs.
            if self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty:
                df_a_salvar = self.desc_finales_de_ultima_busqueda
                tipo_datos_salvados = "RESULTADOS_DESCRIPCION_VIA_DICCIONARIO"
            elif self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty:
                df_a_salvar = self.fcds_de_ultima_busqueda
                tipo_datos_salvados = "COINCIDENCIAS_DICCIONARIO"
        elif self.origen_principal_resultados.es_directo_descripcion or \
             self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            if self.desc_finales_de_ultima_busqueda is not None: # Puede estar vacío si la búsqueda no tuvo resultados
                df_a_salvar = self.desc_finales_de_ultima_busqueda
                tipo_datos_salvados = "RESULTADOS_DESCRIPCION_DIRECTA"
                if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and not self.ultimo_termino_buscado.strip():
                    tipo_datos_salvados = "TODAS_LAS_DESCRIPCIONES"
        
        if df_a_salvar is not None:
            regla = {
                "termino_busqueda_original": self.ultimo_termino_buscado or "N/A (Búsqueda vacía)",
                "fuente_datos_salvados": origen_nombre_actual,
                "tipo_datos_guardados": tipo_datos_salvados,
                "timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                "datos_snapshot_num_filas": len(df_a_salvar),
                # "datos_snapshot": df_a_salvar.to_dict(orient="records") # Esto puede ser muy grande
            }
            self.reglas_guardadas.append(regla) # Guardar metadatos de la regla/búsqueda
            self._actualizar_estado(f"Regla/Búsqueda '{self.ultimo_termino_buscado}' (tipo: {tipo_datos_salvados}, {len(df_a_salvar)} filas) registrada. Total: {len(self.reglas_guardadas)}.")
            logger.info(f"Regla guardada: {regla}")
            messagebox.showinfo("Regla Salvada", f"Metadatos de la búsqueda para '{self.ultimo_termino_buscado}' guardados.\nTipo: {tipo_datos_salvados}\nFilas: {len(df_a_salvar)}")

        else:
            messagebox.showwarning("Nada que Salvar", "No hay datos de resultados claros para la última búsqueda.")
        
        self._actualizar_botones_estado_general()


    def _mostrar_dialogo_seleccion_salvado_via_diccionario(self) -> Dict[str, bool]:
        # (Esta función es llamada por _salvar_regla_actual si se quiere una selección más granular)
        # (Por ahora, la lógica de _salvar_regla_actual es más simple y no la usa directamente)
        # (Se mantiene por si se quiere reactivar esa granularidad)
        decision = {"confirmed": False, "save_fcd": False, "save_rfd": False}
        # ... (implementación original del diálogo) ...
        # Este diálogo necesitaría ser invocado desde _salvar_regla_actual si se quiere esta opción.
        # Por ahora, la he simplificado en _salvar_regla_actual.
        logger.warning("_mostrar_dialogo_seleccion_salvado_via_diccionario no está actualmente en uso directo por _salvar_regla_actual.")
        return decision # Retornar default


    def _exportar_resultados(self):
        if self.resultados_actuales is None or self.resultados_actuales.empty:
            messagebox.showinfo("Exportar Resultados", "No hay resultados actuales para exportar.")
            return
        
        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
            title="Guardar resultados como...",
            initialfile=f"resultados_busqueda_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}" #Sugerir nombre con timestamp
        )
        if not ruta_guardado: # Si el usuario cancela
            return
        
        try:
            if ruta_guardado.endswith(".xlsx"):
                self.resultados_actuales.to_excel(ruta_guardado, index=False)
            elif ruta_guardado.endswith(".csv"):
                self.resultados_actuales.to_csv(ruta_guardado, index=False, encoding='utf-8-sig') # utf-8-sig para mejor compatibilidad CSV con Excel
            else: # Si la extensión no es reconocida o no se especificó correctamente
                messagebox.showerror("Error Formato", "Formato de archivo no soportado para la exportación. Use .xlsx o .csv.")
                return
            
            messagebox.showinfo("Exportar Resultados", f"Resultados exportados correctamente a:\n{ruta_guardado}")
            self._actualizar_estado(f"Resultados exportados a {Path(ruta_guardado).name}")
        except Exception as e:
            logger.exception(f"Error al exportar resultados a '{ruta_guardado}'")
            messagebox.showerror("Error al Exportar", f"No se pudo exportar los resultados:\n{e}")


    def _actualizar_estado_botones_operadores(self):
        # Si no hay datos cargados, todos los botones de operadores deben estar deshabilitados
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None:
            self._deshabilitar_botones_operadores()
            return

        # Habilitar todos por defecto si hay datos, luego deshabilitar según contexto
        for btn in self.op_buttons.values():
            btn["state"] = "normal" 

        texto_actual = self.texto_busqueda_var.get()
        cursor_pos = self.entrada_busqueda.index(tk.INSERT) # Posición actual del cursor
        
        # Lógica simplificada: la mayoría de los operadores se pueden poner si hay texto antes o es el inicio.
        # Se podría refinar más para ser sensible al contexto exacto del último carácter.
        ultimo_char_relevante = texto_actual[:cursor_pos].strip()[-1:] if texto_actual[:cursor_pos].strip() else ""

        # Operadores lógicos como '+' y '|' usualmente necesitan un término antes.
        if not ultimo_char_relevante or ultimo_char_relevante in ["+", "|", "/", "#", "<", ">", "=", " "]:
            if "+" in self.op_buttons: self.op_buttons["+"]["state"] = "disabled"
            if "|" in self.op_buttons: self.op_buttons["|"]["state"] = "disabled"
        
        # El operador de negación '#' usualmente va al inicio de un término.
        if ultimo_char_relevante and ultimo_char_relevante not in ["+", "|", "/", " "]:
             if "#" in self.op_buttons: self.op_buttons["#"]["state"] = "disabled"

        # Operadores de comparación y rango usualmente necesitan un espacio o inicio de término antes.
        # Y no deberían seguir inmediatamente a otro operador de comparación.
        if ultimo_char_relevante in [">", "<", "="]:
            for op_key in [">", "<", ">=", "<=", "=", "-"]:
                if op_key in self.op_buttons: self.op_buttons[op_key]["state"] = "disabled"
        # Si el último carácter es un dígito, no se debería poder poner otro operador de comparación directamente,
        # sino una unidad o un operador lógico. Esto es más complejo de manejar perfectamente aquí.


    def _insertar_operador_validado(self, operador: str):
        # Esta función simplemente inserta el operador. La validación de si es semánticamente correcto
        # en esa posición es compleja y se maneja parcialmente en _actualizar_estado_botones_operadores.
        # El parseador del motor será el validador final.

        # Añadir espacios alrededor de la mayoría de los operadores para legibilidad y ayudar al parseo.
        # La negación '#' es una excepción, usualmente va pegada o seguida de espacio luego.
        texto_a_insertar = operador
        if operador in ["+", "|", "/", ">", "<", ">=", "<=", "=", "-"]: # Todos estos se benefician de espacios
            texto_a_insertar = f" {operador} "
        elif operador == "#": # La negación puede ir seguida de espacio o directamente la palabra/frase
            texto_a_insertar = f"{operador} " # El usuario añadirá comillas

        self.entrada_busqueda.insert(tk.INSERT, texto_a_insertar)
        self.entrada_busqueda.focus_set() # Devolver foco a la entrada


    def _deshabilitar_botones_operadores(self):
        for btn in self.op_buttons.values():
            btn["state"] = "disabled"

    def on_closing(self):
        logger.info("Cerrando la aplicación...")
        self._guardar_configuracion() # Guardar configuración al cerrar
        self.destroy()


if __name__ == "__main__":
    log_file_name = "buscador_app_refactorizado.log" # Nombre de log actualizado
    logging.basicConfig(
        level=logging.INFO, 
        format="%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s",
        handlers=[
            logging.FileHandler(log_file_name, encoding="utf-8", mode="a"), # 'a' para append
            logging.StreamHandler(), # También a consola
        ],
    )
    logger.info("=============================================")
    logger.info(f"=== Iniciando Aplicación Buscador ({Path(__file__).name}) ===")

    # Chequeo de dependencias
    missing_deps = []
    try:
        import pandas
        logger.info(f"Pandas versión: {pandas.__version__}")
    except ImportError:
        missing_deps.append("pandas")
    try:
        import openpyxl
        logger.info(f"openpyxl versión: {openpyxl.__version__}")
    except ImportError: # openpyxl es opcional si solo se usan .xls, pero bueno tenerlo
        missing_deps.append("openpyxl (para archivos .xlsx)")

    if "pandas" in missing_deps: # Pandas es crítico
        error_msg_dep = (f"Falta la librería crítica: pandas.\n"
                         f"Otras librerías opcionales que podrían faltar: {', '.join(d for d in missing_deps if d != 'pandas')}\n"
                         f"Instale con: pip install pandas openpyxl")
        logger.critical(error_msg_dep)
        try: # Intentar mostrar mensaje en UI si Tkinter está disponible
            root_temp = tk.Tk()
            root_temp.withdraw() # Ocultar ventana principal temporal
            messagebox.showerror("Dependencias Faltantes", error_msg_dep)
            root_temp.destroy()
        except tk.TclError: # Si Tkinter mismo falla
            print(f"ERROR CRÍTICO (Tkinter no disponible o falló): {error_msg_dep}")
        except Exception as e_tk_init: # Cualquier otro error al mostrar el msgbox
            print(f"ERROR CRÍTICO (Error al intentar mostrar mensaje de dependencias): {e_tk_init}\nMENSAJE ORIGINAL: {error_msg_dep}")
        exit(1) # Salir si pandas falta

    try:
        app = InterfazGrafica()
        app.mainloop()
    except Exception as main_error:
        logger.critical("¡Error fatal no capturado en la aplicación principal!", exc_info=True)
        # Intentar mostrar un mensaje de error final al usuario
        try:
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror("Error Fatal", f"Ocurrió un error crítico:\n{main_error}\n\nConsulte el archivo de log '{log_file_name}' para más detalles.")
            root_err.destroy()
        except Exception as fallback_error: # Si incluso el mensaje de error falla
            logger.error(f"No se pudo mostrar el mensaje de error fatal vía Tkinter: {fallback_error}")
            print(f"ERROR FATAL EN LA APLICACIÓN: {main_error}. Consulte el archivo '{log_file_name}'.")
    finally:
        logger.info(f"=== Finalizando Aplicación Buscador ({Path(__file__).name}) ===")
