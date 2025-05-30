# -*- coding: utf-8 -*-
# buscador_app/core/motor_busqueda.py

import re
import unicodedata
import logging
from pathlib import Path
from typing import Optional, List, Tuple, Set, Dict, Any, Union
import pandas as pd
import numpy as np

from ..enums import OrigenResultados 
from ..utils import ExtractorMagnitud, ManejadorExcel 

logger = logging.getLogger(__name__)

class MotorBusqueda:
    def __init__(self, indices_diccionario_cfg: Optional[List[int]] = None):
        self.datos_diccionario: Optional[pd.DataFrame] = None
        self.datos_descripcion: Optional[pd.DataFrame] = None
        self.archivo_diccionario_actual: Optional[Path] = None
        self.archivo_descripcion_actual: Optional[Path] = None
        self.indices_columnas_busqueda_dic_preview: List[int] = indices_diccionario_cfg if isinstance(indices_diccionario_cfg, list) else []
        
        logger.info(f"MotorBusqueda inicializado. Índices preview dicc: {self.indices_columnas_busqueda_dic_preview or 'Todas texto/objeto'}")
        
        # Patrones Regex (compilados para eficiencia)
        self.patron_comparacion = re.compile(r"^\s*([<>]=?)\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$")
        self.patron_rango = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?\s*$")
        self.patron_termino_negado = re.compile(r'#\s*(?:\"([^\"]+)\"|([a-zA-ZáéíóúÁÉÍÓÚñÑ0-9\.\-\_]+))', re.IGNORECASE | re.UNICODE)
        self.patron_num_unidad_df = re.compile(r"(\d+(?:[.,]\d+)?)[\s\-]*([a-zA-ZáéíóúÁÉÍÓÚñÑµΩ\.\/\-\_]+)?") # Para extraer num-unidad de celdas DF
        
        self.extractor_magnitud = ExtractorMagnitud() # Instancia del extractor de magnitudes

    def cargar_excel_diccionario(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_diccionario = None
            self.archivo_diccionario_actual = None
            self.extractor_magnitud = ExtractorMagnitud() # Resetear
            return False, error_msg_carga

        mapeo_dinamico_para_extractor: Dict[str, List[str]] = {}
        
        if df_cargado.shape[1] > 0: # Si el DataFrame tiene al menos una columna
            columna_canonica_nombre = df_cargado.columns[0] # Primera columna para formas canónicas
            inicio_col_sinonimos = 3  # Asumiendo que las columnas 0,1,2 son forma canónica, info1, info2
            max_cols_a_chequear_para_sinonimos = df_cargado.shape[1] # Hasta la última columna

            for _, fila in df_cargado.iterrows(): # Iterar sobre cada fila del diccionario
                forma_canonica_raw = fila.get(columna_canonica_nombre)
                if pd.isna(forma_canonica_raw) or str(forma_canonica_raw).strip() == "":
                    continue # Saltar filas sin forma canónica

                forma_canonica_str = str(forma_canonica_raw).strip()
                sinonimos_para_esta_canonica: List[str] = [forma_canonica_str] # Incluir la propia forma canónica

                # Recorrer columnas de sinónimos (desde la 4ta en adelante, índice 3)
                for i in range(inicio_col_sinonimos, max_cols_a_chequear_para_sinonimos):
                    if i < len(df_cargado.columns): # Asegurar que el índice de columna es válido
                        nombre_col_sinonimo_actual = df_cargado.columns[i]
                        sinonimo_celda_raw = fila.get(nombre_col_sinonimo_actual)
                        if pd.notna(sinonimo_celda_raw) and str(sinonimo_celda_raw).strip() != "":
                            sinonimos_para_esta_canonica.append(str(sinonimo_celda_raw).strip())
                
                mapeo_dinamico_para_extractor[forma_canonica_str] = list(set(sinonimos_para_esta_canonica)) # Usar set para eliminar duplicados

            if mapeo_dinamico_para_extractor:
                self.extractor_magnitud = ExtractorMagnitud(mapeo_magnitudes=mapeo_dinamico_para_extractor)
                logger.info(f"Extractor de magnitudes actualizado desde '{ruta.name}' usando formas canónicas y sinónimos.")
            else:
                logger.warning(f"No se extrajeron mapeos de unidad válidos desde '{ruta.name}'. ExtractorMagnitud usará su predefinido (si existe) o estará vacío.")
                self.extractor_magnitud = ExtractorMagnitud() # Re-inicializar a vacío o predefinido
        else:
            logger.warning(f"El archivo de diccionario '{ruta.name}' no tiene columnas. No se pudo actualizar el extractor de magnitudes.")
            self.extractor_magnitud = ExtractorMagnitud() # Re-inicializar

        self.datos_diccionario = df_cargado
        self.archivo_diccionario_actual = ruta

        if logger.isEnabledFor(logging.DEBUG) and self.datos_diccionario is not None:
            logger.debug(f"Archivo de diccionario '{ruta.name}' cargado (primeras 3 filas):\n{self.datos_diccionario.head(3).to_string()}")
        
        return True, None

    def cargar_excel_descripcion(self, ruta_str: str) -> Tuple[bool, Optional[str]]:
        ruta = Path(ruta_str)
        df_cargado, error_msg_carga = ManejadorExcel.cargar_excel(ruta)

        if df_cargado is None:
            self.datos_descripcion = None
            self.archivo_descripcion_actual = None
            return False, error_msg_carga
            
        self.datos_descripcion = df_cargado
        self.archivo_descripcion_actual = ruta
        logger.info(f"Archivo de descripciones '{ruta.name}' cargado.")
        return True, None

    def _obtener_nombres_columnas_busqueda_df(self, df: pd.DataFrame, indices_cfg: List[int], tipo_busqueda: str) -> Tuple[Optional[List[str]], Optional[str]]:
        if df is None or df.empty:
            return None, f"DF para '{tipo_busqueda}' vacío."

        columnas_disponibles = list(df.columns)
        num_cols_df = len(columnas_disponibles)

        if num_cols_df == 0:
            return None, f"DF '{tipo_busqueda}' sin columnas."

        usar_columnas_por_defecto = not indices_cfg or indices_cfg == [-1] # -1 como indicador de usar todas texto/objeto

        if usar_columnas_por_defecto:
            cols_texto_obj = [col for col in columnas_disponibles if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col])]
            if cols_texto_obj:
                logger.debug(f"Para '{tipo_busqueda}', usando columnas de tipo texto/objeto (defecto): {cols_texto_obj}")
                return cols_texto_obj, None
            else:
                # Si no hay de texto/objeto, usar todas como fallback
                logger.warning(f"Para '{tipo_busqueda}', no hay cols texto/objeto. Usando todas las {num_cols_df} columnas: {columnas_disponibles}")
                return columnas_disponibles, None

        nombres_columnas_seleccionadas: List[str] = []
        indices_invalidos: List[str] = []

        for i in indices_cfg:
            if not (isinstance(i, int) and 0 <= i < num_cols_df):
                indices_invalidos.append(str(i))
            else:
                nombres_columnas_seleccionadas.append(columnas_disponibles[i])
        
        if indices_invalidos:
            return None, f"Índice(s) {', '.join(indices_invalidos)} inválido(s) para '{tipo_busqueda}'. Columnas: {num_cols_df} (0 a {num_cols_df-1})."
        
        if not nombres_columnas_seleccionadas: # Si la lista de índices era válida pero no resultó en ninguna columna (ej. lista vacía después de filtrar inválidos)
            return None, f"Config de índices {indices_cfg} no resultó en columnas válidas para '{tipo_busqueda}'."

        logger.debug(f"Para '{tipo_busqueda}', usando columnas por índices {indices_cfg}: {nombres_columnas_seleccionadas}")
        return nombres_columnas_seleccionadas, None

    def _normalizar_para_busqueda(self, texto: str) -> str:
        if not isinstance(texto, str) or not texto:
            return ""
        try:
            # Convertir a mayúsculas
            texto_upper = texto.upper()
            # Normalizar para descomponer acentos (NFKD)
            texto_norm_nfkd = unicodedata.normalize('NFKD', texto_upper)
            # Eliminar caracteres diacríticos (acentos)
            texto_sin_acentos = "".join([c for c in texto_norm_nfkd if not unicodedata.combining(c)])
            # Conservar solo alfanuméricos, espacios y algunos caracteres especiales (. - / _)
            texto_limpio_final = re.sub(r'[^\w\s\.\-\/\_]', '', texto_sin_acentos) # \w incluye números y _
            # Normalizar espacios (múltiples a uno) y quitar espacios al inicio/final
            return ' '.join(texto_limpio_final.split()).strip()
        except Exception as e:
            logger.error(f"Error al normalizar el texto '{texto[:50]}...': {e}")
            return str(texto).upper().strip() # Fallback simple

    def _aplicar_negaciones_y_extraer_positivos(self, df_original: pd.DataFrame, cols: List[str], texto: str) -> Tuple[pd.DataFrame, str, List[str]]:
        texto_limpio_entrada = texto.strip()
        terminos_negados_encontrados: List[str] = []
        df_a_procesar = df_original.copy() if df_original is not None else pd.DataFrame()

        if not texto_limpio_entrada:
            return df_a_procesar, "", terminos_negados_encontrados

        partes_positivas: List[str] = []
        ultimo_indice_fin_negado = 0

        for match_negado in self.patron_termino_negado.finditer(texto_limpio_entrada):
            partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:match_negado.start()])
            ultimo_indice_fin_negado = match_negado.end()
            termino_negado_raw = match_negado.group(1) or match_negado.group(2) # group(1) para comillas, group(2) sin comillas
            if termino_negado_raw:
                termino_negado_normalizado = self._normalizar_para_busqueda(termino_negado_raw.strip('"')) # Quitar comillas si las hay
                if termino_negado_normalizado and termino_negado_normalizado not in terminos_negados_encontrados:
                    terminos_negados_encontrados.append(termino_negado_normalizado)
        
        partes_positivas.append(texto_limpio_entrada[ultimo_indice_fin_negado:])
        terminos_positivos_final_str = ' '.join("".join(partes_positivas).split()).strip()

        if df_a_procesar.empty or not terminos_negados_encontrados or not cols:
            logger.debug(f"Parseo negación: Query='{texto_limpio_entrada}', Positivos='{terminos_positivos_final_str}', Negados={terminos_negados_encontrados}. No se aplicó filtro al DF.")
            return df_a_procesar, terminos_positivos_final_str, terminos_negados_encontrados

        mascara_exclusion_total = pd.Series(False, index=df_a_procesar.index)
        for termino_negado_actual in terminos_negados_encontrados:
            if not termino_negado_actual: # Skip si el término negado es vacío después de normalizar
                continue
            
            mascara_para_este_termino_negado = pd.Series(False, index=df_a_procesar.index)
            # Usar word boundaries (\b) para buscar la palabra/frase exacta negada
            patron_regex_negado = r"\b" + re.escape(termino_negado_actual) + r"\b"
            
            for nombre_columna in cols:
                if nombre_columna not in df_a_procesar.columns:
                    continue
                try:
                    # Normalizar la columna del DataFrame para la comparación
                    serie_columna_normalizada = df_a_procesar[nombre_columna].astype(str).map(self._normalizar_para_busqueda)
                    mascara_para_este_termino_negado |= serie_columna_normalizada.str.contains(patron_regex_negado, regex=True, na=False)
                except Exception as e_neg_col:
                    logger.error(f"Error aplicando negación en col '{nombre_columna}', term '{termino_negado_actual}': {e_neg_col}")

            mascara_exclusion_total |= mascara_para_este_termino_negado
        
        df_resultado_filtrado = df_a_procesar[~mascara_exclusion_total]
        logger.info(f"Filtrado por negación (Query='{texto_limpio_entrada}'): {len(df_a_procesar)} -> {len(df_resultado_filtrado)} filas. Negados: {terminos_negados_encontrados}. Positivos: '{terminos_positivos_final_str}'")
        return df_resultado_filtrado, terminos_positivos_final_str, terminos_negados_encontrados

    def _descomponer_nivel1_or(self, texto_complejo: str) -> Tuple[str, List[str]]:
        texto_limpio = texto_complejo.strip()
        if not texto_limpio:
            return "OR", [] # Si está vacío, no hay segmentos OR

        # Si hay un '+' de alto nivel (no entre paréntesis), se trata como un único bloque AND
        # Esto es para manejar casos como "A + B | C + D" donde el OR es el principal.
        # Pero si es solo "A + B", el "+" es AND.
        if '+' in texto_complejo and not (texto_limpio.startswith("(") and texto_limpio.endswith(")")):
             logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Detectado '+' de alto nivel, tratando como AND. Segmento=['{texto_complejo}']")
             return "AND", [texto_limpio] # Se devuelve como un único segmento que se tratará como AND en N2

        # Solo "|" es un separador OR de alto nivel. Ya no se usa '/' para OR.
        separadores_or = [
            (r"\s*\|\s*", "|") # Para "palabra1 | palabra2"
        ]

        for sep_regex, sep_char_literal in separadores_or:
            # Solo considera el separador OR si no hay un '+' de alto nivel que lo "anule"
            if '+' not in texto_complejo and sep_char_literal in texto_limpio:
                # Usar re.split para manejar el separador como regex
                segmentos_potenciales = [s.strip() for s in re.split(sep_regex, texto_complejo) if s.strip()]
                # Comprobar si la división realmente ocurrió o si el separador estaba al inicio/final
                if len(segmentos_potenciales) > 1 or (len(segmentos_potenciales) == 1 and texto_limpio != segmentos_potenciales[0]):
                    logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Op=OR, Segs={segmentos_potenciales}")
                    return "OR", segmentos_potenciales
        
        # Si no hay OR explícito de alto nivel, se asume AND para este segmento (que puede ser único)
        logger.debug(f"Descomp. N1 (OR) para '{texto_complejo}': Op=AND (no OR explícito de alto nivel), Seg=['{texto_limpio}']")
        return "AND", [texto_limpio]

    def _descomponer_nivel2_and(self, termino_segmento_n1: str) -> Tuple[str, List[str]]:
        termino_limpio = termino_segmento_n1.strip()
        if not termino_limpio:
            return "AND", []

        # Separar por " + " (espacio, más, espacio) para términos AND
        partes_crudas = re.split(r'\s+\+\s+', termino_limpio) # El '+' es literal, \s+ es uno o más espacios
        partes_limpias_finales = [p.strip() for p in partes_crudas if p.strip()]

        logger.debug(f"Descomp. N2 (AND) para '{termino_segmento_n1}': Partes={partes_limpias_finales}")
        return "AND", partes_limpias_finales # El operador siempre es AND en este nivel

    def _analizar_terminos(self, terminos_brutos: List[str]) -> List[Dict[str, Any]]:
        terminos_analizados: List[Dict[str, Any]] = []

        for termino_original_bruto in terminos_brutos:
            termino_original_procesado = str(termino_original_bruto).strip()
            es_frase_exacta = False
            termino_final_para_analisis = termino_original_procesado

            # Detectar si es una frase exacta (entre comillas)
            if len(termino_final_para_analisis) >= 2 and \
               termino_final_para_analisis.startswith('"') and \
               termino_final_para_analisis.endswith('"'):
                termino_final_para_analisis = termino_final_para_analisis[1:-1] # Quitar comillas
                es_frase_exacta = True
            
            if not termino_final_para_analisis: # Si el término queda vacío (ej. solo comillas "")
                continue

            item_analizado: Dict[str, Any] = {"original": termino_final_para_analisis} # Guardar el término sin comillas si era frase

            # Primero, intentar parsear como comparación o rango, solo si NO es frase exacta
            match_comparacion = self.patron_comparacion.match(termino_final_para_analisis)
            match_rango = self.patron_rango.match(termino_final_para_analisis)

            if match_comparacion and not es_frase_exacta:
                operador_str, valor_str, unidad_str_raw = match_comparacion.groups()
                valor_numerico = self._parse_numero(valor_str)
                if valor_numerico is not None:
                    mapa_operadores = {">": "gt", "<": "lt", ">=": "ge", "<=": "le", "=": "eq"}
                    unidad_canonica: Optional[str] = None
                    if unidad_str_raw and unidad_str_raw.strip():
                        unidad_canonica = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_raw.strip())
                    
                    item_analizado.update({
                        "tipo": mapa_operadores.get(operador_str), 
                        "valor": valor_numerico, 
                        "unidad_busqueda": unidad_canonica # Puede ser None si no hay unidad o no se normaliza
                    })
                else: # No se pudo parsear como número, tratar como string
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            
            elif match_rango and not es_frase_exacta:
                valor1_str, valor2_str, unidad_str_r_raw = match_rango.groups()
                valor1_num = self._parse_numero(valor1_str)
                valor2_num = self._parse_numero(valor2_str)
                if valor1_num is not None and valor2_num is not None:
                    unidad_canonica_r: Optional[str] = None
                    if unidad_str_r_raw and unidad_str_r_raw.strip():
                        unidad_canonica_r = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_str_r_raw.strip())
                    
                    item_analizado.update({
                        "tipo": "range", 
                        "valor": sorted([valor1_num, valor2_num]), # Guardar como [min, max]
                        "unidad_busqueda": unidad_canonica_r
                    })
                else: # No se pudo parsear como rango numérico, tratar como string
                    item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            
            else: # Si no es comparación, ni rango, o si es una frase exacta, tratar como string
                  # Si era frase exacta, ya se quitaron las comillas de `termino_final_para_analisis`
                  # La normalización se aplica para la búsqueda, pero el "original" se mantiene para otros usos.
                item_analizado.update({"tipo": "str", "valor": self._normalizar_para_busqueda(termino_final_para_analisis)})
            
            terminos_analizados.append(item_analizado)

        logger.debug(f"Términos (post-AND) analizados para búsqueda detallada: {terminos_analizados}")
        return terminos_analizados

    def _parse_numero(self, num_str: Any) -> Optional[float]:
        if isinstance(num_str, (int, float)):
            return float(num_str)
        
        if not isinstance(num_str, str):
            logger.debug(f"Parseo num: Entrada '{num_str}' no es string.")
            return None
            
        s_limpio = num_str.strip()
        if not s_limpio:
            logger.debug(f"Parseo num: Entrada '{num_str}' vacía tras limpiar.")
            return None
        
        # Loguear el intento
        logger.debug(f"Parseo num: Intentando convertir '{s_limpio}' (originado de '{num_str}')")

        try:
            # Caso 1: Sin comas ni puntos (ej. "123", "-45")
            if ',' not in s_limpio and '.' not in s_limpio:
                logger.debug(f"  '{s_limpio}': Sin separadores. Intento de float directo.")
                return float(s_limpio)
            
            # Normalizar separador decimal a punto (.)
            s_con_puntos = s_limpio.replace(',', '.')
            partes = s_con_puntos.split('.')
            
            # Caso 2: Un solo "punto" después de normalizar comas (ej. "123.45", "1.234" si la coma era decimal)
            if len(partes) == 1: # Esto sucedería si era "123" o si la coma era el único separador y se convirtió a punto
                logger.debug(f"  '{s_limpio}' -> '{s_con_puntos}': Sin puntos post-normalización. Intento float directo.")
                return float(s_con_puntos) # Ej: "1234" o "-500"

            # Caso 3: Múltiples "puntos" (ej. "1.234.567,89" -> "1.234.567.89")
            # O un solo punto decimal (ej. "1234.56")
            ultima_parte = partes[-1]
            partes_principales_str = "".join(partes[:-1]) # Todo menos la última parte, sin puntos.

            # Si la última parte es numérica
            if ultima_parte.isdigit():
                # Subcaso 3a: Última parte tiene 3 o más dígitos (probable separador de miles, ej. 1.234.567)
                # O si es el único punto y es un entero grande "1000." (aunque esto es menos común)
                if len(ultima_parte) >= 3: 
                    # Considerar que los puntos eran separadores de miles
                    numero_reconstruido_str = "".join(partes) # Unir todo sin puntos
                    logger.debug(f"  '{s_limpio}' -> '{s_con_puntos}' -> partes={partes}. Última parte '{ultima_parte}' (>=3 dig) -> miles. Reconstruido: '{numero_reconstruido_str}'")
                    return float(numero_reconstruido_str)
                
                # Subcaso 3b: Última parte tiene 1 o 2 dígitos (probable decimal, ej. xxx.1 o xxx.12)
                elif len(ultima_parte) == 1 or len(ultima_parte) == 2 :
                    # Considerar el último punto como decimal
                    numero_reconstruido_str = f"{partes_principales_str}.{ultima_parte}"
                    logger.debug(f"  '{s_limpio}' -> '{s_con_puntos}' -> partes={partes}. Última parte '{ultima_parte}' (1-2 dig) -> decimal. Reconstruido: '{numero_reconstruido_str}'")
                    return float(numero_reconstruido_str)
                
                # Subcaso 3c: Última parte es vacía (ej. "123." o "1.234.")
                else: # len(ultima_parte) == 0, es decir, s_con_puntos terminaba en "."
                    if not ultima_parte: # Es decir, terminaba en "."
                        # Interpretar "123." como 123.0, o "1.234." como 1234.0
                        if partes_principales_str.isdigit() or (partes_principales_str.startswith('-') and partes_principales_str[1:].isdigit()):
                             logger.debug(f"  '{s_limpio}' -> '{s_con_puntos}' -> partes={partes}. Última parte vacía. Reconstruido: '{partes_principales_str}'")
                             return float(partes_principales_str)
                        else:
                            logger.warning(f"  Formato no reconocido tras quitar punto/coma final para '{num_str}'. Parte principal: '{partes_principales_str}'")
                            return None
                    else: # Este else no debería alcanzarse si ultima_parte.isdigit() y no es len >=3, 1, 2, o 0.
                        logger.warning(f"  Formato de última parte '{ultima_parte}' no reconocido para '{num_str}'.")
                        return None
            else: 
                # Última parte no es puramente numérica (podría ser "kg", "V", etc. o un error de formato)
                logger.warning(f"  Última parte '{ultima_parte}' de '{num_str}' (procesado como '{s_con_puntos}') no es puramente numérica.")
                # Fallback: si solo hay un punto, intentar convertir directamente (ej. "12.3A" -> error, pero "12.3" -> ok)
                if s_con_puntos.count('.') <= 1:
                    try:
                        val_fallback = float(s_con_puntos)
                        logger.debug(f"  Fallback a conversión directa para '{s_con_puntos}' -> {val_fallback}")
                        return val_fallback
                    except ValueError:
                        logger.warning(f"  Fallback falló para '{s_con_puntos}'.")
                        return None
                return None # Si hay múltiples puntos y la última parte no es número, es ambiguo

        except ValueError:
            logger.warning(f"  ValueError final al convertir '{s_limpio}' (originado de '{num_str}') a float.")
            return None
        except Exception as e_parse: # Captura otras excepciones inesperadas
            logger.error(f"  Excepción inesperada '{type(e_parse).__name__}' en _parse_numero para '{s_limpio}': {e_parse}")
            return None

    def _generar_mascara_para_un_termino(self, df: pd.DataFrame, cols: List[str], term_an: Dict[str, Any], filtro_numerico_original: Optional[Dict[str, Any]] = None) -> pd.Series:
        tipo_termino = term_an["tipo"]
        valor_termino = term_an["valor"]
        unidad_requerida_canonica_query = term_an.get("unidad_busqueda")

        # Si se pasa un filtro numérico original (ej. de la query principal cuando este término es un sinónimo de FCD),
        # se usa ese filtro en lugar del que podría tener el término sinónimo.
        valor_a_comparar_final = valor_termino
        unidad_final_para_comparar_canonica = unidad_requerida_canonica_query
        operador_final_para_comparar = tipo_termino
        
        if filtro_numerico_original:
            logger.debug(f"  Aplicando filtro numérico original: {filtro_numerico_original} sobre término actual (sinónimo): {term_an}")
            valor_a_comparar_final = filtro_numerico_original["valor"]
            unidad_final_para_comparar_canonica = filtro_numerico_original.get("unidad_busqueda")
            operador_final_para_comparar = filtro_numerico_original["tipo"]

        mascara_total_termino = pd.Series(False, index=df.index)

        for nombre_columna in cols:
            if nombre_columna not in df.columns:
                continue
            
            columna_serie = df[nombre_columna]
            
            # Búsqueda numérica (gt, lt, ge, le, range, eq con unidad)
            if operador_final_para_comparar in ["gt", "lt", "ge", "le", "range", "eq"]: # "eq" aquí es para numérico+unidad
                mascara_columna_actual_numerica = pd.Series(False, index=df.index)
                
                for indice_fila, valor_celda_raw in columna_serie.items():
                    if pd.isna(valor_celda_raw) or str(valor_celda_raw).strip() == "":
                        continue
                    
                    texto_celda_str = str(valor_celda_raw)
                    # Iterar sobre todos los posibles "numero unidad" en la celda
                    for match_num_unidad_celda in self.patron_num_unidad_df.finditer(texto_celda_str):
                        try:
                            # Validar delimitadores del match para evitar sub-matches incorrectos
                            match_text_completo = match_num_unidad_celda.group(0)
                            inicio_match_en_celda = match_num_unidad_celda.start()
                            fin_match_en_celda = match_num_unidad_celda.end()

                            char_antes_valido = (inicio_match_en_celda == 0) or \
                                                (not texto_celda_str[inicio_match_en_celda - 1].isalnum())
                            char_despues_valido = (fin_match_en_celda == len(texto_celda_str)) or \
                                                  (not texto_celda_str[fin_match_en_celda].isalnum())

                            if not (char_antes_valido and char_despues_valido):
                                logger.debug(f"    Match '{match_text_completo}' descartado por delimitadores en celda: '{texto_celda_str}'")
                                continue # Ir al siguiente match en la misma celda

                            num_celda_str = match_num_unidad_celda.group(1) # El número
                            num_celda_val = self._parse_numero(num_celda_str)
                            unidad_celda_raw = match_num_unidad_celda.group(2) # La unidad (opcional)

                            if num_celda_val is None:
                                continue # No se pudo parsear el número, probar el siguiente match en la celda
                            
                            unidad_celda_canonica = self.extractor_magnitud.obtener_magnitud_normalizada(unidad_celda_raw.strip()) \
                                                    if unidad_celda_raw and unidad_celda_raw.strip() else None
                            
                            # Comprobar coincidencia de unidades:
                            # 1. Si la query no especificó unidad, cualquier unidad de celda es válida (o ninguna).
                            # 2. Si la query especificó unidad, la unidad de la celda (normalizada) debe coincidir.
                            # 3. O si la unidad de la celda (raw normalizada) coincide directamente con la unidad canónica de la query (útil si la normalización de ExtractorMagnitud cambia mucho)
                            unidad_coincide = (unidad_final_para_comparar_canonica is None) or \
                                              (unidad_celda_canonica is not None and unidad_celda_canonica == unidad_final_para_comparar_canonica) or \
                                              (unidad_celda_raw and unidad_final_para_comparar_canonica and self.extractor_magnitud._normalizar_texto(unidad_celda_raw.strip()) == unidad_final_para_comparar_canonica)

                            if not unidad_coincide:
                                continue # Unidad no coincide, probar siguiente match en la celda

                            # Comparación numérica
                            condicion_numerica_cumplida = False
                            if operador_final_para_comparar == "eq" and np.isclose(num_celda_val, valor_a_comparar_final): condicion_numerica_cumplida = True
                            elif operador_final_para_comparar == "gt" and num_celda_val > valor_a_comparar_final and not np.isclose(num_celda_val, valor_a_comparar_final): condicion_numerica_cumplida = True
                            elif operador_final_para_comparar == "lt" and num_celda_val < valor_a_comparar_final and not np.isclose(num_celda_val, valor_a_comparar_final): condicion_numerica_cumplida = True
                            elif operador_final_para_comparar == "ge" and (num_celda_val >= valor_a_comparar_final or np.isclose(num_celda_val, valor_a_comparar_final)): condicion_numerica_cumplida = True
                            elif operador_final_para_comparar == "le" and (num_celda_val <= valor_a_comparar_final or np.isclose(num_celda_val, valor_a_comparar_final)): condicion_numerica_cumplida = True
                            elif operador_final_para_comparar == "range" and \
                                 ((valor_a_comparar_final[0] <= num_celda_val or np.isclose(num_celda_val, valor_a_comparar_final[0])) and \
                                  (num_celda_val <= valor_a_comparar_final[1] or np.isclose(num_celda_val, valor_a_comparar_final[1]))):
                                condicion_numerica_cumplida = True
                            
                            if condicion_numerica_cumplida:
                                # Si estamos aplicando un filtro numérico original (es decir, term_an es un sinónimo de FCD),
                                # también debemos asegurarnos de que el texto del sinónimo esté presente en la celda.
                                if filtro_numerico_original:
                                    texto_sinonimo_normalizado_de_fcd = self._normalizar_para_busqueda(term_an["original"]) # "original" es el texto del sinónimo
                                    patron_regex_sinonimo = r"\b" + re.escape(texto_sinonimo_normalizado_de_fcd) + r"\b"
                                    if re.search(patron_regex_sinonimo, self._normalizar_para_busqueda(texto_celda_str)):
                                        mascara_columna_actual_numerica.at[indice_fila] = True
                                        break # Match encontrado para esta celda, salir del bucle de matches de num_unidad_df
                                    # else: el sinónimo no está, aunque el número/unidad coincida, no es un match para este sinónimo específico
                                else: # No hay filtro numérico original, solo el término actual
                                    mascara_columna_actual_numerica.at[indice_fila] = True
                                    break # Match encontrado para esta celda
                        
                        except ValueError: # Error en _parse_numero para la celda
                            continue # Probar el siguiente match en la celda
                    
                    if mascara_columna_actual_numerica.at[indice_fila]: # Si se marcó True para esta fila, no seguir en otras columnas
                        break # Salir del bucle de celdas de `columna_serie`
                
                mascara_total_termino |= mascara_columna_actual_numerica
            
            # Búsqueda de texto (string) - solo si no es un término numérico o si no estamos aplicando un filtro_numerico_original
            # (porque si hay filtro_numerico_original, la parte textual ya se verificó arriba)
            if tipo_termino == "str" and not filtro_numerico_original : # `valor_termino` aquí ya está normalizado por `_analizar_terminos`
                try:
                    valor_normalizado_busqueda = str(valor_termino) # Asegurar que sea string
                    if not valor_normalizado_busqueda: # Si el término de búsqueda es vacío
                        continue

                    serie_normalizada_df_columna = columna_serie.astype(str).map(self._normalizar_para_busqueda)
                    # Usar word boundaries (\b) para buscar la palabra/frase exacta
                    patron_regex = r"\b" + re.escape(valor_normalizado_busqueda) + r"\b"
                    mascara_columna_actual_str = serie_normalizada_df_columna.str.contains(patron_regex, regex=True, na=False)
                    mascara_total_termino |= mascara_columna_actual_str
                except Exception as e:
                    logger.warning(f"Error búsqueda STR en columna '{nombre_columna}' para término '{valor_termino}': {e}")
        
        return mascara_total_termino

    def _aplicar_mascara_combinada_para_segmento_and(self, df: pd.DataFrame, cols: List[str], term_an_seg: List[Dict[str, Any]], filtro_numerico_original_para_desc: Optional[Dict] = None) -> pd.Series:
        if df is None or df.empty or not cols:
            return pd.Series(False, index=df.index if df is not None else None) # Devolver serie vacía con índice correcto si es posible
        
        if not term_an_seg: # No hay términos para aplicar AND
            return pd.Series(False, index=df.index) # Devolver todos False

        mascara_final = pd.Series(True, index=df.index) # Empezar con todos True para la operación AND

        for term_ind_an in term_an_seg:
            # Manejar sub-queries OR dentro de un segmento AND, ej: "A + (B|C) + D"
            # Si un término 'str' contiene '|' y está entre paréntesis, se trata como sub-query OR.
            if term_ind_an["tipo"] == "str" and \
               ("|" in term_ind_an["original"]) and \
               term_ind_an["original"].startswith("(") and term_ind_an["original"].endswith(")"):
                logger.debug(f"Segmento AND contiene sub-query OR: '{term_ind_an['original']}'. Se procesará por separado.")
                
                sub_mascara_or_series, err_sub_or = self._procesar_busqueda_en_df_objetivo(
                    df, cols, term_ind_an["original"], 
                    None, # No pasar negativos adicionales aquí, se manejan globalmente
                    return_mask_only=True,
                    filtro_numerico_original_desc=None # El filtro numérico original no aplica a esta sub-query textual
                )
                if err_sub_or or sub_mascara_or_series is None:
                    logger.warning(f"Sub-query OR '{term_ind_an['original']}' falló o no devolvió máscara: {err_sub_or}")
                    return pd.Series(False, index=df.index) # Falla el AND completo si la sub-query falla
                
                # Asegurar que la máscara de la sub-query tenga el mismo índice que el DF principal
                mascara_este_term = sub_mascara_or_series.reindex(df.index, fill_value=False)
            else:
                mascara_este_term = self._generar_mascara_para_un_termino(
                    df, cols, term_ind_an, 
                    filtro_numerico_original=filtro_numerico_original_para_desc
                )
            
            mascara_final &= mascara_este_term # Aplicar AND
            if not mascara_final.any(): # Optimización: si ya no hay True, no seguir
                break
        
        return mascara_final

    def _combinar_mascaras_de_segmentos_or(self, lista_mascaras: List[pd.Series], df_idx_ref: Optional[pd.Index] = None) -> pd.Series:
        if not lista_mascaras:
            return pd.Series(False, index=df_idx_ref) if df_idx_ref is not None else pd.Series(dtype=bool)

        # Determinar el índice de referencia para la máscara final
        idx_usar = df_idx_ref
        if idx_usar is None or idx_usar.empty: # Si no se proveyó o está vacío
            if lista_mascaras and not lista_mascaras[0].empty: # Usar el índice de la primera máscara si existe
                idx_usar = lista_mascaras[0].index
        
        if idx_usar is None or idx_usar.empty: # Si aún no hay índice, devolver máscara vacía
             return pd.Series(dtype=bool)
        
        mascara_final = pd.Series(False, index=idx_usar) # Empezar con todos False para la operación OR

        for masc_seg in lista_mascaras:
            if masc_seg.empty:
                continue
            
            mascara_alineada = masc_seg
            # Asegurar que la máscara actual esté alineada con el índice de referencia
            if not masc_seg.index.equals(idx_usar):
                try:
                    mascara_alineada = masc_seg.reindex(idx_usar, fill_value=False)
                except Exception as e_reidx:
                    logger.error(f"Fallo reindex máscara OR: {e_reidx}. Máscara ignorada.")
                    continue # Saltar esta máscara si falla el reindex
            
            mascara_final |= mascara_alineada # Aplicar OR
        
        return mascara_final

    def _procesar_busqueda_en_df_objetivo(self, 
                                        df_obj: pd.DataFrame, 
                                        cols_obj: List[str], 
                                        termino_busqueda_original_para_este_df: str, 
                                        terminos_negativos_adicionales: Optional[List[str]] = None,
                                        return_mask_only: bool = False,
                                        filtro_numerico_original_desc: Optional[Dict] = None
                                        ) -> Union[Tuple[pd.DataFrame, Optional[str]], Tuple[Optional[pd.Series], Optional[str]]]:
        
        logger.debug(f"Proc. búsqueda DF: Query='{termino_busqueda_original_para_este_df}' en {len(cols_obj)} cols de DF ({len(df_obj if df_obj is not None else [])} filas). Neg. Adic: {terminos_negativos_adicionales}, ReturnMask: {return_mask_only}, FiltroNumDesc: {filtro_numerico_original_desc is not None}")

        if df_obj is None: df_obj = pd.DataFrame() # Asegurar que df_obj no sea None

        # 1. Aplicar negaciones de la query actual y extraer positivos
        df_despues_negaciones_query, terminos_positivos_de_query, _ = self._aplicar_negaciones_y_extraer_positivos(
            df_obj, cols_obj, termino_busqueda_original_para_este_df
        )
        df_actual_procesando = df_despues_negaciones_query

        # 2. Aplicar negaciones adicionales (ej. globales pasadas explícitamente)
        if terminos_negativos_adicionales and not df_actual_procesando.empty:
            query_solo_negativos_adicionales = " ".join([f"#{neg}" for neg in terminos_negativos_adicionales if neg])
            if query_solo_negativos_adicionales: # Solo si hay algo que negar
                logger.debug(f"Aplicando neg. ADICIONALES: '{query_solo_negativos_adicionales}' a {len(df_actual_procesando)} filas.")
                df_actual_procesando, _, _ = self._aplicar_negaciones_y_extraer_positivos(
                    df_actual_procesando, cols_obj, query_solo_negativos_adicionales
                )
                logger.info(f"Filtrado por neg. ADICIONALES: {len(df_despues_negaciones_query)} -> {len(df_actual_procesando)} filas.")

        terminos_positivos_final_para_parseo = terminos_positivos_de_query
        idx_ref_base = df_obj.index if df_obj is not None and not df_obj.empty else None

        # 3. Si el DF ya está vacío después de negaciones Y NO hay términos positivos, devolver vacío
        if df_actual_procesando.empty and not terminos_positivos_final_para_parseo.strip():
            logger.debug("DF vacío post-negaciones y sin términos positivos. Devolviendo DF/Máscara vacía.")
            if return_mask_only:
                return pd.Series(False, index=idx_ref_base), None
            else:
                return df_actual_procesando.copy(), None # Devolver el DF vacío pero con columnas

        # 4. Si NO hay términos positivos (pero el DF podría no estar vacío, ej. query puramente negativa)
        if not terminos_positivos_final_para_parseo.strip():
            logger.debug(f"Sin términos positivos ('{terminos_positivos_final_para_parseo}'). Devolviendo DF/Máscara post-negaciones ({len(df_actual_procesando)} filas).")
            if return_mask_only:
                # Crear máscara que refleje df_actual_procesando sobre el índice original df_obj.index
                mask = pd.Series(False, index=idx_ref_base) # Iniciar con todos False en el índice original
                if not df_actual_procesando.empty and idx_ref_base is not None and mask.index.equals(df_actual_procesando.index): # Si los índices coinciden
                    mask.loc[df_actual_procesando.index] = True # Marcar como True los que quedaron
                elif not df_actual_procesando.empty: # Si los índices no coinciden (raro) o no había idx_ref_base
                    mask = pd.Series(True, index=df_actual_procesando.index) # Crear máscara basada en el DF actual
                return mask, None
            else:
                return df_actual_procesando.copy(), None

        # 5. Descomponer términos positivos en OR (Nivel 1) y luego AND (Nivel 2)
        operador_nivel1, segmentos_nivel1_or = self._descomponer_nivel1_or(terminos_positivos_final_para_parseo)

        if not segmentos_nivel1_or: # Si la descomposición OR no da segmentos válidos
            # Esto puede pasar si la query positiva es inválida después del parseo OR (ej. solo operadores)
            if termino_busqueda_original_para_este_df.strip() or terminos_positivos_final_para_parseo.strip(): # Si había algo en la query
                msg_error_segmentos = f"Término positivo '{terminos_positivos_final_para_parseo}' (de query original '{termino_busqueda_original_para_este_df}') es inválido o no generó segmentos de búsqueda."
                logger.warning(msg_error_segmentos)
                if return_mask_only:
                    return pd.Series(False, index=df_actual_procesando.index if not df_actual_procesando.empty else idx_ref_base), msg_error_segmentos
                else: # Devolver DataFrame vacío con columnas originales
                    return pd.DataFrame(columns=df_actual_procesando.columns), msg_error_segmentos
            else: # Si la query original y la positiva eran vacías (ya manejado arriba, pero por si acaso)
                logger.debug("Query original y positiva post-negación vacías. Devolviendo DF/Máscara post-negaciones.")
                # (Lógica duplicada de arriba, podría refactorizarse)
                if return_mask_only:
                    mask = pd.Series(False, index=idx_ref_base)
                    if not df_actual_procesando.empty and idx_ref_base is not None and mask.index.equals(df_actual_procesando.index): mask.loc[df_actual_procesando.index] = True
                    elif not df_actual_procesando.empty: mask = pd.Series(True, index=df_actual_procesando.index)
                    return mask, None
                else: return df_actual_procesando.copy(), None

        # Procesar cada segmento OR
        lista_mascaras_para_or: List[pd.Series] = []
        for segmento_or_actual in segmentos_nivel1_or:
            _operador_nivel2, terminos_brutos_nivel2_and = self._descomponer_nivel2_and(segmento_or_actual)
            terminos_atomicos_analizados_and = self._analizar_terminos(terminos_brutos_nivel2_and)
            
            mascara_para_segmento_or_actual: pd.Series
            if not terminos_atomicos_analizados_and: # Si un segmento AND no tiene términos atómicos
                if operador_nivel1 == "AND": # Si el operador principal era AND, esto es un fallo
                    msg_error_and = f"Segmento AND '{segmento_or_actual}' no produjo términos atómicos válidos. La búsqueda AND completa falla."
                    logger.warning(msg_error_and)
                    if return_mask_only:
                        return pd.Series(False, index=df_actual_procesando.index if not df_actual_procesando.empty else idx_ref_base), msg_error_and
                    else:
                        return pd.DataFrame(columns=df_actual_procesando.columns), msg_error_and
                # Si el operador principal es OR, un segmento vacío simplemente se ignora
                logger.debug(f"Segmento OR '{segmento_or_actual}' no produjo términos atómicos válidos. Se ignorará para la operación OR.")
                mascara_para_segmento_or_actual = pd.Series(False, index=df_actual_procesando.index if not df_actual_procesando.empty else idx_ref_base)
            else:
                # Aplicar la lógica AND para los términos atómicos de este segmento
                mascara_para_segmento_or_actual = self._aplicar_mascara_combinada_para_segmento_and(
                    df_actual_procesando, cols_obj, terminos_atomicos_analizados_and,
                    filtro_numerico_original_para_desc=filtro_numerico_original_desc # Pasar el filtro numérico original
                )
            lista_mascaras_para_or.append(mascara_para_segmento_or_actual)
        
        # Si después de procesar todos los segmentos OR, no hay máscaras (y el DF no estaba vacío)
        if not lista_mascaras_para_or and not df_actual_procesando.empty : # Es un error si había segmentos y DF
            logger.error("Error interno: no se generaron máscaras OR a pesar de tener segmentos N1 y un DataFrame no vacío.")
            if return_mask_only:
                return pd.Series(False, index=idx_ref_base), "Error interno: no se generaron máscaras OR."
            else:
                return pd.DataFrame(columns=df_obj.columns if df_obj is not None else []), "Error interno: no se generaron máscaras OR."
        elif not lista_mascaras_para_or and df_actual_procesando.empty: # Si el DF ya estaba vacío, no hay nada que hacer
            if return_mask_only: return pd.Series(dtype=bool), None # Máscara vacía
            else: return df_actual_procesando.copy(), None # DF vacío

        # Combinar las máscaras de los segmentos OR
        mascara_final_df_objetivo = self._combinar_mascaras_de_segmentos_or(
            lista_mascaras_para_or, 
            df_actual_procesando.index if not df_actual_procesando.empty else idx_ref_base
        )

        # 6. Devolver resultado final
        if return_mask_only:
            logger.debug(f"Devolviendo solo máscara para '{termino_busqueda_original_para_este_df}': {mascara_final_df_objetivo.sum()} coincidencias.")
            return mascara_final_df_objetivo, None
        else:
            df_resultado_final: pd.DataFrame
            # Si la máscara final está vacía (y no tiene el índice del df_actual_procesando o el df está vacío)
            if mascara_final_df_objetivo.empty and (df_actual_procesando.empty or not df_actual_procesando.index.equals(mascara_final_df_objetivo.index)):
                 df_resultado_final = pd.DataFrame(columns=df_obj.columns if df_obj is not None else [])
            elif not mascara_final_df_objetivo.any(): # Si la máscara no tiene ningún True
                 df_resultado_final = pd.DataFrame(columns=df_obj.columns if df_obj is not None else []) # DataFrame vacío con las columnas correctas
            else: # Aplicar la máscara al DataFrame que ya fue filtrado por negaciones
                df_resultado_final = df_actual_procesando[mascara_final_df_objetivo].copy()
            
            logger.debug(f"Resultado _procesar_busqueda_en_df_objetivo para '{termino_busqueda_original_para_este_df}': {len(df_resultado_final)} filas.")
            return df_resultado_final, None

    def _extraer_terminos_de_fila_completa(self, fila_df: pd.Series) -> Set[str]:
        terminos_extraidos_de_fila: Set[str] = set()
        if fila_df is None or fila_df.empty:
            return terminos_extraidos_de_fila

        for valor_celda in fila_df.values: # Iterar sobre todos los valores de la fila
            if pd.notna(valor_celda):
                texto_celda_str = str(valor_celda).strip()
                if texto_celda_str: # Si la celda no está vacía
                    # Normalizar el texto de la celda
                    texto_celda_norm = self._normalizar_para_busqueda(texto_celda_str)
                    # Dividir en palabras y tomar las significativas (ej. >1 caracter, no solo números)
                    palabras_significativas_celda = [
                        palabra for palabra in texto_celda_norm.split() 
                        if len(palabra) > 1 and not palabra.isdigit() # Evitar números solos y palabras de 1 letra
                    ]
                    if palabras_significativas_celda:
                        terminos_extraidos_de_fila.update(palabras_significativas_celda)
                    # Si no hay palabras significativas pero el texto normalizado es usable (y no es numérico)
                    elif texto_celda_norm and len(texto_celda_norm) > 1 and not texto_celda_norm.isdigit() and self._parse_numero(texto_celda_norm) is None:
                        terminos_extraidos_de_fila.add(texto_celda_norm) # Añadir el texto completo de la celda como un término

        return terminos_extraidos_de_fila

    def buscar(self, termino_busqueda_original: str, buscar_via_diccionario_flag: bool) -> Tuple[Optional[pd.DataFrame], OrigenResultados, Optional[pd.DataFrame], Optional[List[int]], Optional[str]]:
        logger.info(f"Motor.buscar INICIO: termino='{termino_busqueda_original}', via_dicc={buscar_via_diccionario_flag}")
        
        columnas_descripcion_ref = self.datos_descripcion.columns if self.datos_descripcion is not None else []
        df_vacio_para_descripciones = pd.DataFrame(columns=columnas_descripcion_ref)
        fcds_obtenidos_final_para_ui: Optional[pd.DataFrame] = None
        indices_fcds_a_resaltar_en_preview: Optional[List[int]] = None

        # --- Manejo de query vacía ---
        if not termino_busqueda_original.strip():
            if self.datos_descripcion is not None:
                return self.datos_descripcion.copy(), OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
            else:
                return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, "Descripciones no cargadas."

        # --- Parseo global de negaciones y positivos ---
        _df_dummy, terminos_positivos_globales, terminos_negativos_globales = self._aplicar_negaciones_y_extraer_positivos(pd.DataFrame(), [], termino_busqueda_original)
        logger.info(f"Parseo global: Positivos='{terminos_positivos_globales}', Negativos Globales={terminos_negativos_globales}")
        
        # --- Detección de filtro numérico/unidad en la query original (si existe y es el primer término) ---
        filtro_numerico_original_de_query: Optional[Dict[str, Any]] = None
        if terminos_positivos_globales.strip():
            _op_l1, segs_l1 = self._descomponer_nivel1_or(terminos_positivos_globales)
            if segs_l1: # Tomar el primer segmento OR (que podría ser toda la query si no hay OR)
                _op_l2, segs_l2 = self._descomponer_nivel2_and(segs_l1[0]) # Tomar el primer sub-término AND
                if segs_l2:
                    terminos_analizados_temp = self._analizar_terminos([segs_l2[0]]) # Analizar solo el primer sub-término
                    if terminos_analizados_temp and \
                       terminos_analizados_temp[0]["tipo"] in ["gt", "lt", "ge", "le", "eq", "range"] and \
                       terminos_analizados_temp[0].get("unidad_busqueda"): # Solo si tiene unidad explícita
                        filtro_numerico_original_de_query = terminos_analizados_temp[0].copy()
                        logger.info(f"Detectado filtro numérico/unidad en query original: {filtro_numerico_original_de_query}")

        # --- Flujo Principal: Búsqueda Vía Diccionario ---
        if buscar_via_diccionario_flag:
            if self.datos_diccionario is None:
                return None, OrigenResultados.ERROR_CARGA_DICCIONARIO, None, None, "Diccionario no cargado."
            
            columnas_dic_para_fcds, err_msg_cols_dic = self._obtener_nombres_columnas_busqueda_df(
                self.datos_diccionario, [], "diccionario_fcds_inicial" # Usar todas las texto/objeto por defecto
            )
            if not columnas_dic_para_fcds:
                return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, None, None, err_msg_cols_dic

            # --- Sub-flujo: Manejo de AND explícito en la query positiva global ---
            # (Ej: "terminoA + terminoB", donde cada parte se busca en FCDs y luego sus sinónimos se combinan en descripciones)
            if "+" in terminos_positivos_globales and not (terminos_positivos_globales.startswith('"') and terminos_positivos_globales.endswith('"')):
                logger.info(f"Detectada búsqueda AND en positivos globales: '{terminos_positivos_globales}'")
                partes_and = [p.strip() for p in terminos_positivos_globales.split("+") if p.strip()]
                
                df_resultado_acumulado_desc = self.datos_descripcion.copy() if self.datos_descripcion is not None else pd.DataFrame(columns=columnas_descripcion_ref)
                fcds_indices_acumulados = set() # Para la preview de FCDs
                todas_partes_and_produjeron_terminos_validos = True
                hay_error_en_busqueda_de_parte_o_desc = False
                error_msg_critico_partes: Optional[str] = None

                if self.datos_descripcion is None:
                    logger.error("Archivo de descripciones no cargado, no se puede proceder con búsqueda AND vía diccionario.")
                    return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas para búsqueda AND."
                
                columnas_desc_para_filtrado, err_cols_desc_fil = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                if not columnas_desc_para_filtrado:
                    return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_fil

                for i, parte_and_actual_str in enumerate(partes_and):
                    if not parte_and_actual_str: continue # Saltar partes vacías

                    logger.debug(f"Procesando parte AND '{parte_and_actual_str}' (parte {i+1}/{len(partes_and)}) en diccionario...")
                    fcds_para_esta_parte, error_fcd_parte = self._procesar_busqueda_en_df_objetivo(
                        self.datos_diccionario, columnas_dic_para_fcds, parte_and_actual_str, None # Sin negativos adicionales aquí
                    )

                    if error_fcd_parte:
                        todas_partes_and_produjeron_terminos_validos = False; hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_fcd_parte
                        logger.warning(f"Parte AND '{parte_and_actual_str}' falló en diccionario con error: {error_fcd_parte}"); break
                    
                    if fcds_para_esta_parte is None or fcds_para_esta_parte.empty:
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no encontró FCDs en diccionario."); break
                    
                    fcds_indices_acumulados.update(fcds_para_esta_parte.index.tolist()) # Acumular índices para la UI

                    terminos_extraidos_de_esta_parte_set: Set[str] = set()
                    for _, fila_fcd in fcds_para_esta_parte.iterrows():
                        terminos_extraidos_de_esta_parte_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))
                    
                    if not terminos_extraidos_de_esta_parte_set:
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' encontró FCDs, pero no se extrajeron términos de ellas."); break

                    terminos_or_con_comillas_actual = [f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_extraidos_de_esta_parte_set if t]
                    query_or_simple_actual = " | ".join(terminos_or_con_comillas_actual)

                    if not query_or_simple_actual: # Si no hay términos válidos para la query OR
                        todas_partes_and_produjeron_terminos_validos = False
                        logger.warning(f"Parte AND '{parte_and_actual_str}' no generó una query OR válida para descripciones."); break
                    
                    # Si el df acumulado ya está vacío, no tiene sentido seguir
                    if df_resultado_acumulado_desc.empty and i >= 0: # i>=0 es siempre true aqui, pero para claridad
                        logger.info(f"Resultados acumulados de descripción vacíos antes de aplicar filtro para '{parte_and_actual_str}'. Búsqueda AND final será vacía."); break
                    
                    logger.info(f"Aplicando filtro OR para '{parte_and_actual_str}' (Query: '{query_or_simple_actual[:100]}...') sobre {len(df_resultado_acumulado_desc)} filas de descripción.")
                    df_resultado_acumulado_desc, error_sub_busqueda_desc = self._procesar_busqueda_en_df_objetivo(
                        df_resultado_acumulado_desc, columnas_desc_para_filtrado, query_or_simple_actual, None # Negativos globales se aplican al final de todo
                    )

                    if error_sub_busqueda_desc:
                        hay_error_en_busqueda_de_parte_o_desc = True; error_msg_critico_partes = error_sub_busqueda_desc
                        logger.error(f"Error en sub-búsqueda OR para '{query_or_simple_actual}': {error_sub_busqueda_desc}"); break
                    
                    if df_resultado_acumulado_desc.empty:
                        logger.info(f"Filtro OR para '{parte_and_actual_str}' no encontró coincidencias en resultados acumulados. Búsqueda AND final será vacía."); break
                
                # Fin del bucle de partes AND
                if fcds_indices_acumulados and self.datos_diccionario is not None:
                    fcds_obtenidos_final_para_ui = self.datos_diccionario.loc[list(fcds_indices_acumulados)].drop_duplicates().copy()
                    indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()
                else:
                    fcds_obtenidos_final_para_ui = pd.DataFrame(columns=self.datos_diccionario.columns if self.datos_diccionario is not None else [])
                    indices_fcds_a_resaltar_en_preview = []
                
                if hay_error_en_busqueda_de_parte_o_desc:
                    return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_msg_critico_partes

                if not todas_partes_and_produjeron_terminos_validos or df_resultado_acumulado_desc.empty:
                    origen_fallo_and = OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS if not todas_partes_and_produjeron_terminos_validos else OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                    logger.info(f"Búsqueda AND '{terminos_positivos_globales}' no produjo resultados finales en descripciones (Origen: {origen_fallo_and.name}).")
                    return df_vacio_para_descripciones, origen_fallo_and, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                
                # Aplicar negaciones globales al resultado final del AND de ORs
                resultados_desc_final_filtrado_and = df_resultado_acumulado_desc
                if not resultados_desc_final_filtrado_and.empty and terminos_negativos_globales:
                    logger.info(f"Aplicando negativos globales {terminos_negativos_globales} a {len(resultados_desc_final_filtrado_and)} filas (resultado del AND de ORs)")
                    query_solo_negados_globales = " ".join([f"#{neg}" for neg in terminos_negativos_globales])
                    df_temp_neg, _, _ = self._aplicar_negaciones_y_extraer_positivos(resultados_desc_final_filtrado_and, columnas_desc_para_filtrado, query_solo_negados_globales)
                    resultados_desc_final_filtrado_and = df_temp_neg
                
                logger.info(f"Búsqueda AND '{terminos_positivos_globales}' vía diccionario produjo {len(resultados_desc_final_filtrado_and)} resultados en descripciones.")
                return resultados_desc_final_filtrado_and, OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None

            # --- Sub-flujo: Manejo de query simple (sin AND explícito de alto nivel) o puramente negativa ---
            else:
                origen_propuesto_flujo_simple: OrigenResultados = OrigenResultados.NINGUNO
                fcds_intento1: Optional[pd.DataFrame] = None

                if terminos_positivos_globales.strip(): # Si hay términos positivos
                    logger.info(f"BUSCAR EN DICC (FCDs) - Intento 1 (Query Original): Query='{terminos_positivos_globales}'")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC
                    try:
                        fcds_temp, error_dic_pos = self._procesar_busqueda_en_df_objetivo(
                            self.datos_diccionario, columnas_dic_para_fcds, terminos_positivos_globales, None # Sin negativos aquí
                        )
                        if error_dic_pos:
                            return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_pos
                        fcds_intento1 = fcds_temp
                    except Exception as e_dic_pos:
                        logger.exception("Excepción búsqueda en diccionario (positivos simples).")
                        return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-positivos simples): {e_dic_pos}"
                
                elif terminos_negativos_globales: # Si NO hay positivos pero SÍ negativos globales
                    logger.info(f"BUSCAR EN DICC (FCDs) - Puramente Negativo: Negs Globales={terminos_negativos_globales}")
                    origen_propuesto_flujo_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC
                    try:
                        query_solo_negados_fcd = " ".join([f"#{neg}" for neg in terminos_negativos_globales])
                        # Aquí la query para _procesar_busqueda_en_df_objetivo es puramente de negación.
                        # El método _aplicar_negaciones_y_extraer_positivos devolverá todos los FCDs menos los negados.
                        fcds_temp, error_dic_neg = self._procesar_busqueda_en_df_objetivo(
                            self.datos_diccionario, columnas_dic_para_fcds, query_solo_negados_fcd, None
                        )
                        if error_dic_neg: # Si la propia query de negación es inválida
                            return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_dic_neg
                        fcds_intento1 = fcds_temp # FCDs que NO contienen los términos negados
                    except Exception as e_dic_neg:
                        logger.exception("Excepción búsqueda en diccionario (puramente negativo).")
                        return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (dicc-negativo): {e_dic_neg}"
                
                else: # Ni positivos ni negativos (esto no debería ocurrir si la query original no era vacía)
                    # Si llegó aquí, es probable que terminos_positivos_globales y terminos_negativos_globales sean vacíos
                    # lo cual es un caso extraño si la query original no lo era.
                    # Devolver DICCIONARIO_SIN_COINCIDENCIAS porque no hay nada que buscar.
                    logger.warning(f"Caso inesperado: Ni términos positivos ni negativos globales para query '{termino_busqueda_original}'.")
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None

                fcds_obtenidos_final_para_ui = fcds_intento1 # Resultado del Intento 1 (positivos o negativos)

                # --- Sub-flujo: Intento 2 (Búsqueda alternativa por unidad si Intento 1 falló y la query original tenía unidad) ---
                if (fcds_obtenidos_final_para_ui is None or fcds_obtenidos_final_para_ui.empty) and \
                   filtro_numerico_original_de_query and \
                   filtro_numerico_original_de_query.get("unidad_busqueda"):
                    
                    unidad_query_original_can = filtro_numerico_original_de_query["unidad_busqueda"]
                    logger.info(f"Intento 1 (numérico+unidad) falló. Iniciando Intento 2: buscando FCDs solo por unidad '{unidad_query_original_can}' en diccionario.")
                    
                    query_solo_unidad_para_fcd = f'"{unidad_query_original_can}"' # Buscar la unidad como frase exacta
                    fcds_por_unidad, err_fcd_unidad = self._procesar_busqueda_en_df_objetivo(
                        self.datos_diccionario, columnas_dic_para_fcds, query_solo_unidad_para_fcd, None
                    )

                    if err_fcd_unidad: # Error en la propia búsqueda de unidad
                        logger.warning(f"Error en búsqueda alternativa de FCDs por unidad '{query_solo_unidad_para_fcd}': {err_fcd_unidad}")
                        # No se puede continuar este flujo, devolver DICCIONARIO_SIN_COINCIDENCIAS
                        return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None 

                    if fcds_por_unidad is not None and not fcds_por_unidad.empty:
                        logger.info(f"Intento 2: Encontrados {len(fcds_por_unidad)} FCDs alternativos basados solo en la unidad '{query_solo_unidad_para_fcd}'.")
                        fcds_obtenidos_final_para_ui = fcds_por_unidad # Actualizar los FCDs para la UI
                        indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()

                        terminos_de_unidad_para_desc_set: Set[str] = set()
                        for _, fila_fcd_unidad in fcds_por_unidad.iterrows():
                            terminos_de_unidad_para_desc_set.update(self._extraer_terminos_de_fila_completa(fila_fcd_unidad))

                        if not terminos_de_unidad_para_desc_set:
                            logger.info("Intento 2: FCDs por unidad encontrados, pero no se extrajeron términos para descripciones.")
                            return df_vacio_para_descripciones, OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                        
                        query_or_de_unidades_para_desc = " | ".join([f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_de_unidad_para_desc_set if t])
                        if not query_or_de_unidades_para_desc:
                            return df_vacio_para_descripciones, OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Query OR de unidades para descripciones (alternativa) vacía."
                        
                        if self.datos_descripcion is None:
                            return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Descripciones no cargadas."
                        
                        columnas_desc_alt, err_cols_desc_alt = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds_alt")
                        if not columnas_desc_alt:
                            return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, err_cols_desc_alt

                        logger.info(f"BUSCAR EN DESC (Intento 2 - vía FCDs por unidad): Query sinónimos='{query_or_de_unidades_para_desc[:100]}...'. Aplicando filtro numérico original: {filtro_numerico_original_de_query} y Neg. Globales: {terminos_negativos_globales}")
                        
                        # Los negativos globales se aplican siempre, excepto si la query original era puramente negativa (porque ya se aplicaron para obtener FCDs)
                        neg_glob_alt = terminos_negativos_globales if origen_propuesto_flujo_simple != OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC else []
                        
                        resultados_desc_alt, error_desc_alt = self._procesar_busqueda_en_df_objetivo(
                            self.datos_descripcion, columnas_desc_alt, query_or_de_unidades_para_desc,
                            terminos_negativos_adicionales=neg_glob_alt, # Aplicar negativos globales
                            filtro_numerico_original_desc=filtro_numerico_original_de_query # Aplicar el filtro numérico original
                        )

                        if error_desc_alt: # Error en la búsqueda en descripciones
                            return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_desc_alt
                        
                        if resultados_desc_alt is None or resultados_desc_alt.empty:
                            return df_vacio_para_descripciones, OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                        else:
                            return resultados_desc_alt, OrigenResultados.VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                    else: # Intento 2: No se encontraron FCDs basados solo en la unidad.
                        logger.info(f"Intento 2: No se encontraron FCDs basados solo en la unidad '{query_solo_unidad_para_fcd}'. Revirtiendo a DICCIONARIO_SIN_COINCIDENCIAS.")
                        return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None # No hay FCDs, no hay resultados

                # --- Continuación del Flujo Simple (si Intento 1 dio FCDs o si Intento 2 no aplicó/falló y volvemos a DICC_SIN_COINCIDENCIAS) ---
                if fcds_obtenidos_final_para_ui is not None and not fcds_obtenidos_final_para_ui.empty: 
                    # Tenemos FCDs (del intento 1)
                    if indices_fcds_a_resaltar_en_preview is None: # Si no se asignaron en intento 2
                        indices_fcds_a_resaltar_en_preview = fcds_obtenidos_final_para_ui.index.tolist()

                    logger.info(f"FCDs obtenidas del diccionario (flujo estándar simple/negativo): {len(fcds_obtenidos_final_para_ui)} filas.")

                    if self.datos_descripcion is None:
                        return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Descripciones no cargadas."
                    
                    terminos_para_buscar_en_descripcion_set: Set[str] = set()
                    for _, fila_fcd in fcds_obtenidos_final_para_ui.iterrows():
                        terminos_para_buscar_en_descripcion_set.update(self._extraer_terminos_de_fila_completa(fila_fcd))

                    if not terminos_para_buscar_en_descripcion_set:
                        logger.info("FCDs encontrados (flujo estándar), pero no se extrajeron términos para descripciones.")
                        origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                        if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                            origen_final_sinterm = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                        return df_vacio_para_descripciones, origen_final_sinterm, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                    
                    logger.info(f"Términos para desc ({len(terminos_para_buscar_en_descripcion_set)} únicos, muestra): {sorted(list(terminos_para_buscar_en_descripcion_set))[:10]}...")
                    
                    terminos_or_con_comillas_desc = [f'"{t}"' if " " in t and not (t.startswith('"') and t.endswith('"')) else t for t in terminos_para_buscar_en_descripcion_set if t]
                    query_or_para_desc_simple = " | ".join(terminos_or_con_comillas_desc)

                    if not query_or_para_desc_simple: # Si no hay términos válidos para la query OR
                        origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS
                        if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                            origen_q_vacia = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                        return df_vacio_para_descripciones, origen_q_vacia, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, "Query OR para descripciones vacía."

                    columnas_desc_final_simple, err_cols_desc_final_simple = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion_fcds")
                    if not columnas_desc_final_simple:
                        return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, err_cols_desc_final_simple
                    
                    # Aplicar negativos globales si la query original no era puramente negativa
                    # (si era puramente negativa, los FCDs ya están filtrados, y se buscan esos FCDs en desc.)
                    negativos_a_aplicar_desc_simple = terminos_negativos_globales if origen_propuesto_flujo_simple != OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC else []
                    
                    logger.info(f"BUSCAR EN DESC (vía FCD estándar): Query='{query_or_para_desc_simple[:200]}...'. Neg. Adicionales a aplicar en Desc: {negativos_a_aplicar_desc_simple}")
                    try:
                        resultados_desc_final_simple, error_busqueda_desc_simple = self._procesar_busqueda_en_df_objetivo(
                            self.datos_descripcion, columnas_desc_final_simple, query_or_para_desc_simple,
                            terminos_negativos_adicionales=negativos_a_aplicar_desc_simple
                            # NO se pasa filtro_numerico_original_desc aquí, porque la query original (si era numérica) ya filtró los FCDs.
                            # La búsqueda en descripciones es por los sinónimos de esos FCDs.
                        )
                        if error_busqueda_desc_simple:
                            return df_vacio_para_descripciones, OrigenResultados.TERMINO_INVALIDO, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, error_busqueda_desc_simple
                        
                        if resultados_desc_final_simple is None or resultados_desc_final_simple.empty:
                            origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                            if origen_propuesto_flujo_simple == OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC:
                                origen_res_desc_vacio_simple = OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC
                            return df_vacio_para_descripciones, origen_res_desc_vacio_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                        else:
                            # El origen es el que se propuso al inicio del flujo simple (VIA_DICC_CON_RES o VIA_DICC_PURAMENTE_NEG_CON_RES)
                            return resultados_desc_final_simple, origen_propuesto_flujo_simple, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, None
                            
                    except Exception as e_desc_proc_simple:
                        logger.exception("Excepción búsqueda final en descripciones (flujo estándar).")
                        return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, fcds_obtenidos_final_para_ui, indices_fcds_a_resaltar_en_preview, f"Error motor (desc final estándar): {e_desc_proc_simple}"
                else: 
                    # No se encontraron FCDs en el Intento 1, y el Intento 2 (por unidad) no aplicó o también falló.
                    logger.info(f"No se encontraron FCDs en diccionario para '{termino_busqueda_original}' y el flujo alternativo no aplicó.")
                    return df_vacio_para_descripciones, OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS, None, None, None
        
        # --- Flujo Alternativo: Búsqueda Directa en Descripciones (si buscar_via_diccionario_flag es False) ---
        else: 
            if self.datos_descripcion is None:
                return None, OrigenResultados.ERROR_CARGA_DESCRIPCION, None, None, "Descripciones no cargadas."

            columnas_desc_directo, err_cols_desc_directo = self._obtener_nombres_columnas_busqueda_df(self.datos_descripcion, [], "descripcion")
            if not columnas_desc_directo:
                return None, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC, None, None, err_cols_desc_directo
            
            try:
                logger.info(f"BUSCAR EN DESC (DIRECTO): Query '{termino_busqueda_original}'")
                # La query original (con sus negaciones) se pasa directamente.
                # El filtro numérico (si existe en la query original) se aplicará directamente en descripciones.
                resultados_directos_desc, error_busqueda_desc_dir = self._procesar_busqueda_en_df_objetivo(
                    self.datos_descripcion, columnas_desc_directo, termino_busqueda_original, None,
                    filtro_numerico_original_desc=filtro_numerico_original_de_query # Aplicar el numérico de la query si existe
                )

                if error_busqueda_desc_dir:
                    return None, OrigenResultados.TERMINO_INVALIDO, None, None, error_busqueda_desc_dir
                
                if resultados_directos_desc is None or resultados_directos_desc.empty:
                    return df_vacio_para_descripciones, OrigenResultados.DIRECTO_DESCRIPCION_VACIA, None, None, None
                else:
                    return resultados_directos_desc, OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, None, None, None
            
            except Exception as e_desc_dir_proc:
                logger.exception("Excepción búsqueda directa en descripciones.")
                return None, OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR, None, None, f"Error motor (desc directa): {e_desc_dir_proc}"