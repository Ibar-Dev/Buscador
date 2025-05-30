# -*- coding: utf-8 -*-
# buscador_app/enums.py

from enum import Enum, auto

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
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC = auto()
    VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC = auto() 
    VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC = auto()

    @property
    def es_via_diccionario(self) -> bool:
        return self in {
            OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
            OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC,
            OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS,
            OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC,
            OrigenResultados.VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC,
            OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC,
        }
    @property
    def es_directo_descripcion(self) -> bool:
        return self in {OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS, OrigenResultados.DIRECTO_DESCRIPCION_VACIA}
    @property
    def es_error_carga(self) -> bool:
        return self in {OrigenResultados.ERROR_CARGA_DICCIONARIO, OrigenResultados.ERROR_CARGA_DESCRIPCION}
    @property
    def es_error_configuracion(self) -> bool:
        return self in {OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DICC, OrigenResultados.ERROR_CONFIGURACION_COLUMNAS_DESC}
    @property
    def es_error_operacional(self) -> bool: return self == OrigenResultados.ERROR_BUSQUEDA_INTERNA_MOTOR
    @property
    def es_termino_invalido(self) -> bool: return self == OrigenResultados.TERMINO_INVALIDO