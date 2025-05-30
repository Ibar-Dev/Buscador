# -*- coding: utf-8 -*-
# buscador_app/utils.py

import re
import unicodedata
import logging
from pathlib import Path
from typing import Optional, List, Dict, Tuple, Union, Any # Any para _parse_numero
import pandas as pd
# import numpy as np # No se usa directamente en este archivo, pero sí en motor_busqueda que lo usa.

logger = logging.getLogger(__name__)

class ExtractorMagnitud:
    MAPEO_MAGNITUDES_PREDEFINIDO: Dict[str, List[str]] = {} 

    def __init__(self, mapeo_magnitudes: Optional[Dict[str, List[str]]] = None):
        self.sinonimo_a_canonico_normalizado: Dict[str, str] = {} 
        mapeo_a_usar = mapeo_magnitudes if mapeo_magnitudes is not None else self.MAPEO_MAGNITUDES_PREDEFINIDO
        
        for forma_canonica_original, lista_sinonimos_originales in mapeo_a_usar.items(): 
            canonico_norm = self._normalizar_texto(forma_canonica_original) 
            if not canonico_norm: 
                logger.warning(f"Forma canónica '{forma_canonica_original}' resultó vacía tras normalizar y fue ignorada en ExtractorMagnitud.")
                continue
            
            self.sinonimo_a_canonico_normalizado[canonico_norm] = canonico_norm 
            
            for sinonimo_original in lista_sinonimos_originales: 
                sinonimo_norm = self._normalizar_texto(str(sinonimo_original)) 
                if sinonimo_norm: 
                    self.sinonimo_a_canonico_normalizado[sinonimo_norm] = canonico_norm 
        logger.debug(f"ExtractorMagnitud inicializado/actualizado con {len(self.sinonimo_a_canonico_normalizado)} mapeos normalizados.")


    @staticmethod
    def _normalizar_texto(texto: str) -> str:
        if not isinstance(texto, str) or not texto: 
            return "" 
        try:
            texto_upper = texto.upper() 
            forma_normalizada = unicodedata.normalize("NFKD", texto_upper) 
            # Permitir alfanuméricos, espacios, y . - _ /
            res = "".join(c for c in forma_normalizada if not unicodedata.combining(c) and (c.isalnum() or c.isspace() or c in ['.', '-', '_', '/']))
            return ' '.join(res.split()) # Normalizar espacios múltiples y trim
        except TypeError: 
            logger.error(f"TypeError en _normalizar_texto (ExtractorMagnitud) con entrada: {texto}")
            return ""

    def obtener_magnitud_normalizada(self, texto_unidad: str) -> Optional[str]:
        if not texto_unidad: 
            return None 
        normalizada = self._normalizar_texto(texto_unidad) 
        return self.sinonimo_a_canonico_normalizado.get(normalizada) if normalizada else None

class ManejadorExcel:
    @staticmethod
    def cargar_excel(ruta_archivo: Union[str, Path]) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        ruta = Path(ruta_archivo) 
        if not ruta.exists(): 
            mensaje_error = f"¡Archivo no encontrado! Ruta: {ruta}"
            logger.error(f"ManejadorExcel: {mensaje_error}") 
            return None, mensaje_error 
        try:
            engine: Optional[str] = None 
            if ruta.suffix.lower() == ".xlsx": 
                engine = "openpyxl" 
            
            logger.info(f"ManejadorExcel: Cargando '{ruta.name}' con engine='{engine or 'auto (pandas intentará xlrd para .xls)'}'...")
            df = pd.read_excel(ruta, engine=engine) 
            logger.info(f"ManejadorExcel: Archivo '{ruta.name}' ({len(df)} filas) cargado exitosamente.")
            return df, None 
            
        except ImportError as ie: 
            mensaje_error_usuario = (
                f"Error al cargar '{ruta.name}': Falta librería.\n"
                f"Para .xlsx: pip install openpyxl\n"
                f"Para .xls: pip install xlrd\n"
                f"Detalle: {ie}"
            )
            logger.exception(f"ManejadorExcel: Falta dependencia para leer '{ruta.name}'. Error: {ie}") 
            return None, mensaje_error_usuario
            
        except Exception as e: 
            mensaje_error_usuario = (
                f"No se pudo cargar '{ruta.name}': {e}\n"
                f"Verifique formato, permisos y si está en uso."
            )
            logger.exception(f"ManejadorExcel: Error genérico al cargar '{ruta.name}'.") 
            return None, mensaje_error_usuario