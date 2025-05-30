# -*- coding: utf-8 -*-
# buscador_app/gui/interfaz_grafica.py

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional, List, Dict, Any, Union
import platform
import json
import os
from pathlib import Path
import logging
import traceback

# Importaciones de otros módulos del paquete buscador_app
from ..core.motor_busqueda import MotorBusqueda # Importación relativa
from ..enums import OrigenResultados # Importación relativa

logger = logging.getLogger(__name__)

class InterfazGrafica(tk.Tk):
    CONFIG_FILE_NAME = "config_buscador_avanzado_ui.json" # Nombre del archivo de configuración

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado v1.10.3 Mod (Modularizado)") 
        self.geometry("1250x800") # Tamaño inicial de la ventana

        # Cargar configuración de la aplicación
        self.config: Dict[str, Any] = self._cargar_configuracion_app()
        indices_cfg_preview_dic = self.config.get("indices_columnas_busqueda_dic_preview", [])
        
        # Inicializar el motor de búsqueda
        self.motor = MotorBusqueda(indices_diccionario_cfg=indices_cfg_preview_dic)

        # Variables de estado de la UI
        self.resultados_actuales: Optional[pd.DataFrame] = None
        self.texto_busqueda_var = tk.StringVar(self)
        self.texto_busqueda_var.trace_add("write", self._on_texto_busqueda_change) # Actualizar botones al escribir
        self.ultimo_termino_buscado: Optional[str] = None
        self.reglas_guardadas: List[Dict[str, Any]] = [] # Para la funcionalidad "Salvar Regla"

        # Datos de la última búsqueda para la UI y "Salvar Regla"
        self.fcds_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.desc_finales_de_ultima_busqueda: Optional[pd.DataFrame] = None
        self.indices_fcds_resaltados: Optional[List[int]] = None
        self.origen_principal_resultados: OrigenResultados = OrigenResultados.NINGUNO

        # Colores para las tablas Treeview
        self.color_fila_par: str = "white"
        self.color_fila_impar: str = "#f0f0f0" # Un gris claro
        self.color_resaltado_dic: str = "sky blue" # Color para FCDs resaltados

        self.op_buttons: Dict[str, ttk.Button] = {} # Diccionario para botones de operadores

        # Configuración inicial de la UI
        self._configurar_estilo_ttk_app()
        self._crear_widgets_app()
        self._configurar_grid_layout_app()
        self._configurar_eventos_globales_app()
        self._configurar_tags_estilo_treeview_app()
        self._configurar_funcionalidad_orden_tabla(self.tabla_resultados)
        self._configurar_funcionalidad_orden_tabla(self.tabla_diccionario)

        self._actualizar_mensaje_barra_estado("Listo. Cargue Diccionario y Descripciones.")
        self._deshabilitar_botones_operadores() # Deshabilitar al inicio
        self._actualizar_estado_general_botones_y_controles() # Actualizar estado inicial de botones
        
        logger.info(f"Interfaz Gráfica (v1.10.3 Modularizada) inicializada.")

    def _try_except_wrapper(self, func, *args, **kwargs):
        """ Wrapper para capturar excepciones en funciones de UI y loggearlas. """
        try:
            return func(*args, **kwargs)
        except Exception as e:
            func_name = func.__name__
            error_type = type(e).__name__
            error_msg = str(e)
            tb_str = traceback.format_exc() # Obtener el traceback completo
            
            logger.critical(f"Error en {func_name}: {error_type} - {error_msg}\n{tb_str}")
            print(f"--- TRACEBACK COMPLETO (desde _try_except_wrapper para {func_name}) ---\n{tb_str}") # Imprimir en consola
            
            messagebox.showerror(
                f"Error Interno en {func_name}",
                f"Ocurrió un error inesperado:\n{error_type}: {error_msg}\n\nConsulte el log y la consola para el traceback completo."
            )
            # Si el error ocurre durante la carga de archivos, actualizar UI para reflejar el fallo
            if func_name in ["_cargar_diccionario_ui", "_cargar_excel_descripcion_ui"]:
                self._actualizar_etiquetas_archivos_cargados()
                self._actualizar_estado_general_botones_y_controles()
            return None # Devolver None en caso de error

    def _on_texto_busqueda_change(self, var_name: str, index: str, mode: str):
        """ Se llama cada vez que el texto en la entrada de búsqueda cambia. """
        self._actualizar_estado_botones_operadores()

    def _cargar_configuracion_app(self) -> Dict[str, Any]:
        """ Carga la configuración de la aplicación desde un archivo JSON. """
        config_cargada: Dict[str, Any] = {}
        ruta_archivo_config = Path(self.CONFIG_FILE_NAME)

        if ruta_archivo_config.exists():
            try:
                with ruta_archivo_config.open("r", encoding="utf-8") as f:
                    config_cargada = json.load(f)
                logger.info(f"Configuración cargada desde: {self.CONFIG_FILE_NAME}")
            except Exception as e:
                logger.error(f"Error al cargar config '{self.CONFIG_FILE_NAME}': {e}")
        else:
            logger.info(f"Archivo config '{self.CONFIG_FILE_NAME}' no encontrado. Se usará config por defecto.")
        
        # Convertir rutas de string a Path si existen
        for clave_ruta in ["last_dic_path", "last_desc_path"]:
            valor_ruta = config_cargada.get(clave_ruta)
            config_cargada[clave_ruta] = Path(valor_ruta) if valor_ruta else None
        
        # Asegurar que la clave de índices para preview exista
        config_cargada.setdefault("indices_columnas_busqueda_dic_preview", []) # Por defecto, lista vacía
        return config_cargada

    def _guardar_configuracion_app(self):
        """ Guarda la configuración actual de la aplicación en un archivo JSON. """
        # Actualizar rutas de los últimos archivos cargados
        self.config["last_dic_path"] = str(self.motor.archivo_diccionario_actual) if self.motor.archivo_diccionario_actual else None
        self.config["last_desc_path"] = str(self.motor.archivo_descripcion_actual) if self.motor.archivo_descripcion_actual else None
        # Guardar configuración de columnas de preview del diccionario
        self.config["indices_columnas_busqueda_dic_preview"] = self.motor.indices_columnas_busqueda_dic_preview

        try:
            with open(self.CONFIG_FILE_NAME, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4) # Guardar con indentación para legibilidad
            logger.info(f"Configuración guardada en: {self.CONFIG_FILE_NAME}")
        except Exception as e:
            logger.error(f"Error al guardar config '{self.CONFIG_FILE_NAME}': {e}")

    def _configurar_estilo_ttk_app(self):
        """ Configura el tema y estilo de los widgets ttk. """
        style = ttk.Style(self)
        os_name = platform.system()
        
        # Preferencias de tema por sistema operativo
        theme_preferences = {
            "Windows": ["vista", "xpnative"], # Temas comunes en Windows
            "Darwin": ["aqua"],             # Tema nativo de macOS
            "Linux": ["clam", "alt"]        # Temas comunes en Linux
        }
        
        # Seleccionar el primer tema disponible de la lista de preferencias, o el actual/default
        available_themes = style.theme_names()
        preferred_themes_for_os = theme_preferences.get(os_name, ["clam"]) # Fallback a 'clam'
        
        chosen_theme = style.theme_use() # Tema actual como fallback
        for theme_name in preferred_themes_for_os:
            if theme_name in available_themes:
                chosen_theme = theme_name
                break
        
        try:
            style.theme_use(chosen_theme)
            # Estilo personalizado para botones de operadores
            style.configure("Operator.TButton", padding=(2,1), font=("TkDefaultFont", 9))
            logger.info(f"Tema TTK aplicado: {chosen_theme}")
        except tk.TclError: # Por si el tema elegido falla
            logger.warning(f"Fallo al aplicar tema TTK '{chosen_theme}'. Usando tema por defecto.")
            # No es necesario hacer nada más, Tkinter usará un tema por defecto.

    def _crear_widgets_app(self):
        """ Crea todos los widgets de la interfaz gráfica. """
        # --- Marco de Controles Principal ---
        self.marco_controles = ttk.LabelFrame(self, text="Controles")

        # Carga de Diccionario
        self.btn_cargar_diccionario = ttk.Button(self.marco_controles, text="Cargar Diccionario", command=lambda: self._try_except_wrapper(self._cargar_diccionario_ui))
        self.lbl_dic_cargado = ttk.Label(self.marco_controles, text="Dic: Ninguno", width=25, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)

        # Carga de Descripciones
        self.btn_cargar_descripciones = ttk.Button(self.marco_controles, text="Cargar Descripciones", command=lambda: self._try_except_wrapper(self._cargar_excel_descripcion_ui))
        self.lbl_desc_cargado = ttk.Label(self.marco_controles, text="Desc: Ninguno", width=25, anchor=tk.W, relief=tk.SUNKEN, borderwidth=1)

        # Botones de Operadores de Búsqueda
        self.frame_ops = ttk.Frame(self.marco_controles)
        op_buttons_defs = [("+", "+"), ("|", "|"), ("#", "#"), ("> ", ">"), ("< ", "<"), ("≥ ", ">="), ("≤ ", "<="), ("-", "-")]
        for i, (text, op_val_clean) in enumerate(op_buttons_defs):
            btn = ttk.Button(self.frame_ops, text=text, command=lambda op=op_val_clean: self._insertar_operador_validado(op), style="Operator.TButton", width=3)
            btn.grid(row=0, column=i, padx=1, pady=1, sticky="nsew")
            self.op_buttons[op_val_clean] = btn

        # Entrada de Búsqueda y Botones de Acción
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=60, textvariable=self.texto_busqueda_var)
        self.btn_buscar = ttk.Button(self.marco_controles, text="Buscar", command=lambda: self._try_except_wrapper(self._ejecutar_busqueda_ui))
        self.btn_salvar_regla = ttk.Button(self.marco_controles, text="Salvar Regla", command=lambda: self._try_except_wrapper(self._salvar_regla_actual_ui), state="disabled")
        self.btn_ayuda = ttk.Button(self.marco_controles, text="?", command=self._mostrar_ayuda_ui, width=3)
        self.btn_exportar = ttk.Button(self.marco_controles, text="Exportar", command=lambda: self._try_except_wrapper(self._exportar_resultados_ui), state="disabled")

        # --- Tabla de Vista Previa del Diccionario ---
        self.lbl_tabla_diccionario = ttk.Label(self, text="Vista Previa Diccionario:")
        self.frame_tabla_diccionario = ttk.Frame(self)
        self.tabla_diccionario = ttk.Treeview(self.frame_tabla_diccionario, show="headings", height=8)
        self.scrolly_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="vertical", command=self.tabla_diccionario.yview)
        self.scrollx_diccionario = ttk.Scrollbar(self.frame_tabla_diccionario, orient="horizontal", command=self.tabla_diccionario.xview)
        self.tabla_diccionario.configure(yscrollcommand=self.scrolly_diccionario.set, xscrollcommand=self.scrollx_diccionario.set)

        # --- Tabla de Resultados / Descripciones ---
        self.lbl_tabla_resultados = ttk.Label(self, text="Resultados / Descripciones:")
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados, show="headings") # Altura se ajustará con el grid_rowconfigure
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)

        # --- Barra de Estado ---
        self.barra_estado = ttk.Label(self, text="Listo.", relief=tk.SUNKEN, anchor=tk.W, borderwidth=1)
        self._actualizar_etiquetas_archivos_cargados() # Inicializar etiquetas de archivos

    def _configurar_grid_layout_app(self):
        """ Configura el layout de la aplicación usando grid. """
        # Configuración de expansión de filas y columnas principales de la ventana
        self.grid_rowconfigure(2, weight=1)  # Fila para tabla diccionario (peso menor)
        self.grid_rowconfigure(4, weight=3)  # Fila para tabla resultados (más espacio, peso mayor)
        self.grid_columnconfigure(0, weight=1) # Columna principal se expande

        # Posicionamiento del marco de controles (ocupa toda la primera fila de la ventana)
        self.marco_controles.grid(row=0, column=0, sticky="new", padx=10, pady=(10, 5))
        # Configurar expansión de columnas dentro del marco de controles
        self.marco_controles.grid_columnconfigure(1, weight=1) # Para que lbl_dic_cargado se expanda
        self.marco_controles.grid_columnconfigure(3, weight=1) # Para que lbl_desc_cargado se expanda

        # Widgets dentro del marco de controles (Fila 0 de marco_controles)
        self.btn_cargar_diccionario.grid(row=0, column=0, padx=(5, 0), pady=5, sticky="w")
        self.lbl_dic_cargado.grid(row=0, column=1, padx=(2, 10), pady=5, sticky="ew")
        self.btn_cargar_descripciones.grid(row=0, column=2, padx=(5, 0), pady=5, sticky="w")
        self.lbl_desc_cargado.grid(row=0, column=3, padx=(2, 5), pady=5, sticky="ew")

        # Frame de botones de operadores (Fila 1 de marco_controles)
        self.frame_ops.grid(row=1, column=0, columnspan=6, padx=5, pady=(5, 0), sticky="ew") # columnspan para ocupar todas las columnas disponibles
        for i in range(len(self.op_buttons)): # Hacer que los botones de operadores se expandan uniformemente
            self.frame_ops.grid_columnconfigure(i, weight=1)

        # Entrada de búsqueda y botones de acción (Fila 2 de marco_controles)
        self.entrada_busqueda.grid(row=2, column=0, columnspan=2, padx=5, pady=(0, 5), sticky="ew")
        self.btn_buscar.grid(row=2, column=2, padx=(2, 0), pady=(0, 5), sticky="w")
        self.btn_salvar_regla.grid(row=2, column=3, padx=(2, 0), pady=(0, 5), sticky="w")
        self.btn_ayuda.grid(row=2, column=4, padx=(2, 0), pady=(0, 5), sticky="w")
        self.btn_exportar.grid(row=2, column=5, padx=(10, 5), pady=(0, 5), sticky="e") # Alinear a la derecha

        # --- Tablas (fuera del marco_controles, en la ventana principal) ---
        # Etiqueta y Tabla para Diccionario
        self.lbl_tabla_diccionario.grid(row=1, column=0, sticky="sw", padx=10, pady=(10, 0)) # Debajo de controles, alineado S-W
        self.frame_tabla_diccionario.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_diccionario.grid_rowconfigure(0, weight=1)
        self.frame_tabla_diccionario.grid_columnconfigure(0, weight=1)
        self.tabla_diccionario.grid(row=0, column=0, sticky="nsew")
        self.scrolly_diccionario.grid(row=0, column=1, sticky="ns")
        self.scrollx_diccionario.grid(row=1, column=0, sticky="ew")

        # Etiqueta y Tabla para Resultados
        self.lbl_tabla_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(0, 0)) # Debajo de tabla dicc, alineado S-W
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.scrolly_resultados.grid(row=0, column=1, sticky="ns")
        self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # Barra de Estado (al final de la ventana)
        self.barra_estado.grid(row=5, column=0, sticky="sew", padx=0, pady=(5,0)) # Ocupa todo el ancho en la parte inferior

    def _configurar_eventos_globales_app(self):
        """ Configura bindings de eventos globales. """
        # Bind para buscar al presionar Enter en la entrada de búsqueda
        self.entrada_busqueda.bind("<Return>", lambda event: self._try_except_wrapper(self._ejecutar_busqueda_ui))
        # Protocolo para manejar el cierre de la ventana
        self.protocol("WM_DELETE_WINDOW", self.on_closing_app)

    def _actualizar_mensaje_barra_estado(self, mensaje: str):
        """ Actualiza el texto de la barra de estado. """
        self.barra_estado.config(text=mensaje)
        logger.info(f"Mensaje UI (BarraEstado): {mensaje}")
        self.update_idletasks() # Forzar actualización inmediata de la UI

    def _mostrar_ayuda_ui(self):
        """ Muestra una ventana de ayuda con la sintaxis de búsqueda. """
        texto_ayuda = (
            "Sintaxis de Búsqueda:\n"
            "- Búsqueda simple: `router cisco` (busca ambas palabras)\n"
            "- Operador AND: `tarjeta + 16 puertos` (todos los términos deben estar)\n"
            "- Operador OR: `modulo | SFP` (alguno de los términos debe estar)\n"
            "  (Nota: `/` ya no se interpreta como OR)\n"
            "- Búsqueda numérica comparativa: `>1000W`, `<50V`, `>=48A`, `<=10.5W`\n"
            "  (Unidad opcional, pegada al número o separada por espacio)\n"
            "- Búsqueda numérica por rango: `10-20V` (ej. 10V a 20V, inclusive)\n"
            "  (Unidad opcional, pegada al segundo número o separada por espacio)\n"
            "- Frase exacta: `\"rack de 19 pulgadas\"` (busca la frase literal)\n"
            "- Negación: `#palabra_a_excluir` o `# \"frase a excluir\"`\n\n"
            "Flujo de Búsqueda (Vía Diccionario por defecto):\n"
            "1. La consulta se busca primero en el 'Diccionario'.\n"
            "2. Si hay coincidencias en el Diccionario (FCDs - Formas Canónicas del Diccionario):\n"
            "   a. Se extraen todos los términos/sinónimos de las filas FCD coincidentes.\n"
            "   b. Estos términos se usan para buscar (como un gran OR) en el archivo de 'Descripciones'.\n"
            "   c. Se aplica la condición numérica original (si la hubo) y las negaciones globales a estos resultados.\n"
            "3. Si la consulta original era una operación AND (ej. `A + B`):\n"
            "   a. Cada parte ('A', 'B') se busca en el Diccionario para obtener sus FCDs/sinónimos.\n"
            "   b. Los resultados en 'Descripciones' deben cumplir con los sinónimos de 'A' Y los sinónimos de 'B'.\n"
            "4. Si la consulta original tenía un filtro numérico con unidad (ej. `>10V`) y no da FCDs:\n"
            "   a. Se intenta una búsqueda alternativa: se buscan FCDs que solo contengan la unidad (`V`).\n"
            "   b. Los sinónimos de estos FCDs se buscan en 'Descripciones', aplicando el filtro numérico original (`>10`).\n"
            "5. Si no hay coincidencias en Diccionario (o si la búsqueda vía FCDs no da resultados en Descripciones):\n"
            "   Se ofrecerá la opción de buscar el término original directamente en 'Descripciones'.\n\n"
            "Búsqueda Directa en Descripciones:\n"
            "- La consulta (incluyendo operadores, números, frases, negaciones) se aplica directamente sobre el archivo de Descripciones."
        )
        messagebox.showinfo("Ayuda - Sintaxis y Flujo de Búsqueda", texto_ayuda)

    def _configurar_tags_estilo_treeview_app(self):
        """ Configura los tags para estilizar filas en las tablas Treeview. """
        for tabla in [self.tabla_diccionario, self.tabla_resultados]:
            tabla.tag_configure("par", background=self.color_fila_par)
            tabla.tag_configure("impar", background=self.color_fila_impar)
        # Tag específico para resaltar filas en la tabla de diccionario
        self.tabla_diccionario.tag_configure("resaltado_azul", background=self.color_resaltado_dic, foreground="black")

    def _configurar_funcionalidad_orden_tabla(self, tabla: ttk.Treeview):
        """ Habilita la ordenación de columnas en una tabla Treeview. """
        columnas = tabla["columns"]
        if columnas: # Solo si la tabla tiene columnas definidas
            for col_id in columnas:
                # El texto de la cabecera ya debería estar seteado.
                # El comando se enlaza a la función de ordenación.
                tabla.heading(col_id, text=str(col_id), anchor=tk.W, 
                              command=lambda c=col_id, t=tabla: self._try_except_wrapper(self._ordenar_columna_tabla_ui, t, c, False))

    def _ordenar_columna_tabla_ui(self, tabla: ttk.Treeview, columna_id: str, orden_reverso: bool):
        """ Ordena los datos de una tabla Treeview por la columna especificada. """
        df_para_ordenar: Optional[pd.DataFrame] = None
        indices_para_resaltar_post_orden: Optional[List[int]] = None

        # Determinar qué DataFrame usar según la tabla
        if tabla == self.tabla_diccionario and self.motor.datos_diccionario is not None:
            df_para_ordenar = self.motor.datos_diccionario.copy() # Trabajar con una copia
            indices_para_resaltar_post_orden = self.indices_fcds_resaltados # Mantener resaltados si es tabla diccionario
        elif tabla == self.tabla_resultados and self.resultados_actuales is not None:
            df_para_ordenar = self.resultados_actuales.copy()
        else: # No hay datos para ordenar
            tabla.heading(columna_id, command=lambda c=columna_id, t=tabla: self._try_except_wrapper(self._ordenar_columna_tabla_ui, t, c, not orden_reverso))
            return

        if df_para_ordenar.empty or columna_id not in df_para_ordenar.columns:
            tabla.heading(columna_id, command=lambda c=columna_id, t=tabla: self._try_except_wrapper(self._ordenar_columna_tabla_ui, t, c, not orden_reverso))
            return

        # Intentar conversión numérica para ordenar, si no, ordenar como string
        # Guardar el estado de si la columna es mayormente numérica
        es_columna_numerica = pd.to_numeric(df_para_ordenar[columna_id], errors='coerce').notna().sum() > (len(df_para_ordenar) / 2)

        if es_columna_numerica:
            # Ordenar numéricamente, tratando errores como NaN (que van al final/principio según na_position)
            df_ordenado = df_para_ordenar.sort_values(
                by=columna_id, 
                ascending=not orden_reverso, 
                na_position='last', # Poner NaNs al final
                key=lambda x: pd.to_numeric(x, errors='coerce')
            )
        else:
            # Ordenar alfabéticamente (ignorando mayúsculas/minúsculas)
            df_ordenado = df_para_ordenar.sort_values(
                by=columna_id, 
                ascending=not orden_reverso, 
                na_position='last',
                key=lambda x: x.astype(str).str.lower() # Convertir a string y luego a minúsculas
            )
        
        # Determinar qué columnas mostrar si es la tabla de diccionario (para preview)
        columnas_a_mostrar_en_dicc_ordenado = None
        if tabla == self.tabla_diccionario and self.motor.datos_diccionario is not None:
            columnas_a_mostrar_en_dicc_ordenado, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                df_ordenado, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview"
            )
            if not columnas_a_mostrar_en_dicc_ordenado: # Fallback si no se obtienen columnas específicas
                columnas_a_mostrar_en_dicc_ordenado = list(df_ordenado.columns)
        
        # Actualizar la tabla en la UI
        if tabla == self.tabla_diccionario:
            # Para el diccionario, si se ordena, se muestran todas las filas para mantener el contexto de los resaltados
            self._actualizar_tabla_treeview_ui(tabla, df_ordenado, limite_filas=None, 
                                               columnas_a_mostrar=columnas_a_mostrar_en_dicc_ordenado, 
                                               indices_a_resaltar=indices_para_resaltar_post_orden)
        elif tabla == self.tabla_resultados:
            self.resultados_actuales = df_ordenado # Actualizar el DataFrame de resultados actual
            self._actualizar_tabla_treeview_ui(tabla, self.resultados_actuales) # Límite de filas por defecto se aplicará

        # Actualizar el comando de la cabecera para invertir el orden en el próximo clic
        tabla.heading(columna_id, command=lambda c=columna_id, t=tabla: self._try_except_wrapper(self._ordenar_columna_tabla_ui, t, c, not orden_reverso))
        self._actualizar_mensaje_barra_estado(f"Tabla ordenada por '{columna_id}'.")

    def _actualizar_tabla_treeview_ui(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], 
                                      limite_filas: Optional[int] = None, 
                                      columnas_a_mostrar: Optional[List[Union[str, int]]] = None, 
                                      indices_a_resaltar: Optional[List[int]] = None):
        """ Limpia y rellena una tabla Treeview con datos de un DataFrame. """
        es_tabla_diccionario = (tabla == self.tabla_diccionario)
        nombre_tabla_log = "Diccionario" if es_tabla_diccionario else "Resultados"

        # Limpiar contenido anterior
        for item in tabla.get_children():
            tabla.delete(item)
        tabla["columns"] = () # Resetear columnas

        if datos is None or datos.empty:
            self._configurar_funcionalidad_orden_tabla(tabla) # Reconfigurar ordenación en tabla vacía
            logger.debug(f"Tabla '{nombre_tabla_log}' vaciada (sin datos para mostrar).")
            return

        columnas_originales_df = list(datos.columns)
        cols_finales_para_tabla: List[str]

        # Determinar qué columnas mostrar
        if columnas_a_mostrar:
            if all(isinstance(c, int) for c in columnas_a_mostrar): # Si son índices numéricos
                try:
                    cols_finales_para_tabla = [columnas_originales_df[i] for i in columnas_a_mostrar if 0 <= i < len(columnas_originales_df)]
                except IndexError:
                    logger.warning(f"Índices en columnas_a_mostrar fuera de rango para tabla '{nombre_tabla_log}'. Usando todas las columnas.")
                    cols_finales_para_tabla = columnas_originales_df
            elif all(isinstance(c, str) for c in columnas_a_mostrar): # Si son nombres de columnas
                cols_finales_para_tabla = [c for c in columnas_a_mostrar if c in columnas_originales_df]
            else: # Tipo inesperado
                logger.warning(f"Tipo inesperado para columnas_a_mostrar en tabla '{nombre_tabla_log}'. Usando todas las columnas.")
                cols_finales_para_tabla = columnas_originales_df
            
            if not cols_finales_para_tabla: # Si la selección resultó en ninguna columna válida
                logger.warning(f"columnas_a_mostrar no resultó en columnas válidas para tabla '{nombre_tabla_log}'. Usando todas las columnas.")
                cols_finales_para_tabla = columnas_originales_df
        else: # Si no se especifican columnas, mostrar todas
            cols_finales_para_tabla = columnas_originales_df

        if not cols_finales_para_tabla: # Si al final no hay columnas que mostrar
            self._configurar_funcionalidad_orden_tabla(tabla)
            logger.debug(f"Tabla '{nombre_tabla_log}' sin columnas usables para mostrar.")
            return
            
        tabla["columns"] = tuple(cols_finales_para_tabla)

        # Configurar cabeceras y ancho de columnas
        for col_id in cols_finales_para_tabla:
            tabla.heading(col_id, text=str(col_id), anchor=tk.W) # Comando de ordenación se añade por separado
            try:
                # Calcular ancho basado en contenido y cabecera (heurística)
                if col_id in datos.columns:
                    ancho_contenido = datos[col_id].astype(str).str.len().quantile(0.95) if not datos[col_id].empty else 0
                else: # Columna no existe en los datos (no debería pasar si cols_finales_para_tabla se generó bien)
                    ancho_contenido = 0
                
                ancho_cabecera = len(str(col_id))
                ancho_columna = max(70, min(int(max(ancho_cabecera * 7, ancho_contenido * 5.5) + 15), 350)) # Ajustar multiplicadores y límites
            except Exception as e_ancho:
                logger.warning(f"Error calculando ancho para columna '{col_id}' en tabla '{nombre_tabla_log}': {e_ancho}")
                ancho_columna = 100 # Ancho por defecto en caso de error
            tabla.column(col_id, anchor=tk.W, width=ancho_columna, minwidth=50)

        # Iterar sobre el DataFrame para añadir filas (limitado si es necesario)
        df_a_iterar = datos[cols_finales_para_tabla] # Seleccionar solo las columnas que se mostrarán
        num_filas_original_df = len(df_a_iterar)

        # Condición para mostrar todas las filas: si es la tabla de diccionario, hay índices para resaltar, y el DF no está vacío.
        mostrar_todas_filas_por_resaltado = es_tabla_diccionario and indices_a_resaltar and num_filas_original_df > 0

        if not mostrar_todas_filas_por_resaltado and limite_filas and num_filas_original_df > limite_filas:
            df_a_iterar = df_a_iterar.head(limite_filas)
        elif mostrar_todas_filas_por_resaltado: # Si se muestran todas por resaltado, loguearlo
             logger.debug(f"Mostrando todas las {num_filas_original_df} filas de '{nombre_tabla_log}' debido a la presencia de índices a resaltar.")

        # Insertar filas en la tabla
        for i, (idx_original_df, fila_datos) in enumerate(df_a_iterar.iterrows()):
            valores_fila = [str(v) if pd.notna(v) else "" for v in fila_datos.values]
            tags_fila = ["par" if i % 2 == 0 else "impar"]
            
            if es_tabla_diccionario and indices_a_resaltar and idx_original_df in indices_a_resaltar:
                tags_fila.append("resaltado_azul")
            
            try: # Usar un IID único basado en el índice original del DataFrame para la fila
                tabla.insert("", "end", values=valores_fila, tags=tuple(tags_fila), iid=f"row_{idx_original_df}")
            except Exception as e_insert: # Error si el IID ya existe o hay otro problema
                logger.warning(f"Error insertando fila con índice original {idx_original_df} en tabla '{nombre_tabla_log}': {e_insert}. Se intentará con IID alternativo.")
                try: # Fallback a IID genérico si el basado en índice falla (menos ideal para resaltado)
                     tabla.insert("", "end", values=valores_fila, tags=tuple(tags_fila), iid=f"generic_row_{i}")
                except Exception as e_insert_alt:
                     logger.error(f"Fallo al insertar fila con IID alternativo en tabla '{nombre_tabla_log}': {e_insert_alt}")


        self._configurar_funcionalidad_orden_tabla(tabla) # Reaplicar bindings de ordenación
        logger.debug(f"Tabla '{nombre_tabla_log}' actualizada con {len(tabla.get_children())} filas visibles.")

    def _actualizar_etiquetas_archivos_cargados(self):
        """ Actualiza las etiquetas que muestran los nombres de los archivos cargados. """
        max_longitud_nombre = 25 # Para evitar que las etiquetas se hagan muy largas
        
        path_diccionario = self.motor.archivo_diccionario_actual
        path_descripciones = self.motor.archivo_descripcion_actual

        nombre_dic = path_diccionario.name if path_diccionario else "Ninguno"
        nombre_desc = path_descripciones.name if path_descripciones else "Ninguno"

        # Acortar nombres si son muy largos
        texto_label_dic = f"Dic: {nombre_dic}" if len(nombre_dic) <= max_longitud_nombre else f"Dic: ...{nombre_dic[-(max_longitud_nombre-4):]}"
        texto_label_desc = f"Desc: {nombre_desc}" if len(nombre_desc) <= max_longitud_nombre else f"Desc: ...{nombre_desc[-(max_longitud_nombre-4):]}"
        
        # Actualizar texto y color de las etiquetas
        self.lbl_dic_cargado.config(text=texto_label_dic, foreground="green" if path_diccionario else "red")
        self.lbl_desc_cargado.config(text=texto_label_desc, foreground="green" if path_descripciones else "red")

    def _actualizar_estado_general_botones_y_controles(self):
        """ Habilita o deshabilita botones según el estado de la aplicación (archivos cargados, resultados, etc.). """
        diccionario_cargado = self.motor.datos_diccionario is not None
        descripciones_cargadas = self.motor.datos_descripcion is not None

        # Actualizar botones de operadores si hay al menos un archivo cargado
        if diccionario_cargado or descripciones_cargadas:
            self._actualizar_estado_botones_operadores()
        else:
            self._deshabilitar_botones_operadores()

        # Botón Buscar: habilitado solo si ambos archivos están cargados
        self.btn_buscar["state"] = "normal" if diccionario_cargado and descripciones_cargadas else "disabled"

        # Botón Salvar Regla: condiciones más complejas
        puede_salvar_regla = False
        if self.ultimo_termino_buscado and self.origen_principal_resultados != OrigenResultados.NINGUNO:
            # Si fue vía diccionario y hay FCDs O (hay resultados en desc y el origen es de éxito con desc)
            if self.origen_principal_resultados.es_via_diccionario and \
               ((self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty) or \
                (self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty and \
                 self.origen_principal_resultados in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, 
                                                    OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC,
                                                    OrigenResultados.VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC])):
                puede_salvar_regla = True
            # Si fue búsqueda directa en descripción (con o sin resultados) y hay un DataFrame de descripción final
            elif (self.origen_principal_resultados.es_directo_descripcion or \
                  self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA) and \
                 self.desc_finales_de_ultima_busqueda is not None:
                puede_salvar_regla = True
        self.btn_salvar_regla["state"] = "normal" if puede_salvar_regla else "disabled"

        # Botón Exportar: habilitado si hay resultados actuales y no están vacíos
        self.btn_exportar["state"] = "normal" if (self.resultados_actuales is not None and not self.resultados_actuales.empty) else "disabled"

    def _cargar_diccionario_ui(self):
        """ Abre un diálogo para cargar el archivo de diccionario. """
        # Usar la ruta del último archivo cargado (si existe) como directorio inicial
        path_config_dic = self.config.get("last_dic_path")
        dir_inicial = str(Path(path_config_dic).parent) if path_config_dic and Path(path_config_dic).exists() else os.getcwd()
        
        ruta_seleccionada = filedialog.askopenfilename(
            title="Seleccionar Archivo de Diccionario",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            initialdir=dir_inicial
        )
        if not ruta_seleccionada: # Usuario canceló
            return

        nombre_archivo = Path(ruta_seleccionada).name
        self._actualizar_mensaje_barra_estado(f"Cargando diccionario: {nombre_archivo}...")
        
        # Limpiar vistas previas y estados de búsqueda anteriores
        self._actualizar_tabla_treeview_ui(self.tabla_diccionario, None)
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, None)
        self.resultados_actuales = None
        self.fcds_de_ultima_busqueda = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO
        self.indices_fcds_resaltados = None
        
        # Cargar el archivo usando el motor
        carga_ok, mensaje_error = self.motor.cargar_excel_diccionario(ruta_seleccionada)
        
        nombre_archivo_desc_actual = Path(self.motor.archivo_descripcion_actual).name if self.motor.archivo_descripcion_actual else "N/A"

        if carga_ok and self.motor.datos_diccionario is not None:
            self.config["last_dic_path"] = Path(ruta_seleccionada) # Guardar nueva ruta
            self._guardar_configuracion_app()
            
            df_dicc = self.motor.datos_diccionario
            num_filas_dicc = len(df_dicc)
            
            # Obtener columnas para la vista previa del diccionario
            columnas_preview_dicc, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                df_dicc, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview"
            )
            
            self.lbl_tabla_diccionario.config(text=f"Diccionario ({num_filas_dicc} filas)")
            self._actualizar_tabla_treeview_ui(self.tabla_diccionario, df_dicc, limite_filas=100, columnas_a_mostrar=columnas_preview_dicc)
            
            self.title(f"Buscador - Dic: {nombre_archivo} | Desc: {nombre_archivo_desc_actual}")
            self._actualizar_mensaje_barra_estado(f"Diccionario '{nombre_archivo}' ({num_filas_dicc} filas) cargado exitosamente.")
        else:
            self._actualizar_mensaje_barra_estado(f"Error al cargar diccionario: {mensaje_error or 'Error desconocido'}")
            messagebox.showerror("Error al Cargar Diccionario", mensaje_error or "Ocurrió un error desconocido al cargar el archivo.")
            self.title(f"Buscador - Dic: N/A (Error) | Desc: {nombre_archivo_desc_actual}")
        
        self._actualizar_etiquetas_archivos_cargados()
        self._actualizar_estado_general_botones_y_controles()

    def _cargar_excel_descripcion_ui(self):
        """ Abre un diálogo para cargar el archivo de descripciones. """
        path_config_desc = self.config.get("last_desc_path")
        dir_inicial = str(Path(path_config_desc).parent) if path_config_desc and Path(path_config_desc).exists() else os.getcwd()

        ruta_seleccionada = filedialog.askopenfilename(
            title="Seleccionar Archivo de Descripciones",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
            initialdir=dir_inicial
        )
        if not ruta_seleccionada:
            logger.info("Carga de archivo de descripciones cancelada por el usuario.")
            return

        nombre_archivo = Path(ruta_seleccionada).name
        self._actualizar_mensaje_barra_estado(f"Cargando descripciones: {nombre_archivo}...")

        # Limpiar resultados actuales y tabla de resultados
        self.resultados_actuales = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO # Resetear origen
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, None)
        
        carga_ok, mensaje_error = self.motor.cargar_excel_descripcion(ruta_seleccionada)
        
        nombre_archivo_dicc_actual = Path(self.motor.archivo_diccionario_actual).name if self.motor.archivo_diccionario_actual else "N/A"

        if carga_ok and self.motor.datos_descripcion is not None:
            self.config["last_desc_path"] = Path(ruta_seleccionada)
            self._guardar_configuracion_app()
            
            df_desc = self.motor.datos_descripcion
            num_filas_desc = len(df_desc)
            
            self._actualizar_mensaje_barra_estado(f"Archivo de descripciones '{nombre_archivo}' ({num_filas_desc} filas) cargado. Mostrando vista previa...")
            # Mostrar una vista previa de las descripciones cargadas en la tabla de resultados
            self._actualizar_tabla_treeview_ui(self.tabla_resultados, df_desc, limite_filas=200) # Limitar a 200 filas para preview
            
            self.title(f"Buscador - Dic: {nombre_archivo_dicc_actual} | Desc: {nombre_archivo}")
        else:
            error_a_mostrar = mensaje_error or "Ocurrió un error desconocido al cargar el archivo de descripciones."
            self._actualizar_mensaje_barra_estado(f"Error al cargar descripciones: {error_a_mostrar}")
            messagebox.showerror("Error al Cargar Archivo de Descripciones", error_a_mostrar)
            self.title(f"Buscador - Dic: {nombre_archivo_dicc_actual} | Desc: N/A (Error)")
            
        self._actualizar_etiquetas_archivos_cargados()
        self._actualizar_estado_general_botones_y_controles()

    def _ejecutar_busqueda_ui(self):
        """ Ejecuta la búsqueda con el término ingresado y actualiza la UI. """
        if self.motor.datos_diccionario is None or self.motor.datos_descripcion is None:
            messagebox.showwarning("Archivos Faltantes", "Cargue el archivo de Diccionario y el archivo de Descripciones antes de buscar.")
            return

        termino_busqueda_actual = self.texto_busqueda_var.get()
        self.ultimo_termino_buscado = termino_busqueda_actual # Guardar para posible "Salvar Regla"

        # Resetear estados de resultados previos de esta búsqueda específica
        self.resultados_actuales = None
        self.fcds_de_ultima_busqueda = None
        self.desc_finales_de_ultima_busqueda = None
        self.origen_principal_resultados = OrigenResultados.NINGUNO
        self.indices_fcds_resaltados = None
        
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, None) # Limpiar tabla de resultados antes de nueva búsqueda
        self._actualizar_mensaje_barra_estado(f"Buscando '{termino_busqueda_actual}'...")

        # --- Ejecutar la búsqueda en el motor ---
        # Por defecto, intentar vía diccionario (buscar_via_diccionario_flag=True)
        resultados_df, origen_actual, fcds_encontrados, indices_resaltar, msg_error_motor = self.motor.buscar(
            termino_busqueda_original=termino_busqueda_actual, 
            buscar_via_diccionario_flag=True # Flag para indicar que se intente la lógica vía diccionario
        )

        # Actualizar estado interno con los resultados de la búsqueda
        self.fcds_de_ultima_busqueda = fcds_encontrados
        self.origen_principal_resultados = origen_actual
        self.indices_fcds_resaltados = indices_resaltar
        
        # Columnas de referencia para DataFrames vacíos
        columnas_df_desc_ref = self.motor.datos_descripcion.columns if self.motor.datos_descripcion is not None else []

        # Actualizar tabla de diccionario (resaltando FCDs si los hay)
        if self.motor.datos_diccionario is not None:
            num_fcds_resaltados = len(self.indices_fcds_resaltados) if self.indices_fcds_resaltados else 0
            etiqueta_diccionario = f"Diccionario ({len(self.motor.datos_diccionario)} filas)"
            if num_fcds_resaltados > 0 and origen_actual.es_via_diccionario and origen_actual != OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS:
                etiqueta_diccionario += f" - {num_fcds_resaltados} FCDs resaltados"
            self.lbl_tabla_diccionario.config(text=etiqueta_diccionario)
            
            columnas_preview_dicc_actual, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                self.motor.datos_diccionario, 
                self.motor.indices_columnas_busqueda_dic_preview, 
                "diccionario_preview"
            )
            # Mostrar todas las filas del diccionario si hay FCDs resaltados, sino, el límite normal
            limite_filas_dic_preview = None if self.indices_fcds_resaltados and num_fcds_resaltados > 0 else 100
            self._actualizar_tabla_treeview_ui(
                self.tabla_diccionario, 
                self.motor.datos_diccionario, 
                limite_filas=limite_filas_dic_preview, 
                columnas_a_mostrar=columnas_preview_dicc_actual,
                indices_a_resaltar=self.indices_fcds_resaltados
            )

        # --- Procesar resultados y mensajes según el origen de los mismos ---
        if msg_error_motor and origen_actual.es_error_operacional:
            messagebox.showerror("Error del Motor de Búsqueda", f"Ocurrió un error interno en el motor: {msg_error_motor}")
            self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_ref) # Mostrar tabla vacía

        elif origen_actual.es_error_carga or origen_actual.es_error_configuracion or origen_actual.es_termino_invalido:
            messagebox.showerror("Error de Búsqueda", msg_error_motor or f"Se produjo un error durante la búsqueda: {origen_actual.name}")
            self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_ref)

        elif origen_actual in [OrigenResultados.VIA_DICCIONARIO_CON_RESULTADOS_DESC, 
                               OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_CON_RESULTADOS_DESC,
                               OrigenResultados.VIA_DICCIONARIO_UNIDAD_Y_NUMERICO_EN_DESC]:
            self.resultados_actuales = resultados_df
            num_fcds = len(fcds_encontrados) if fcds_encontrados is not None else 0
            num_desc = len(resultados_df) if resultados_df is not None else 0
            self._actualizar_mensaje_barra_estado(f"Búsqueda para '{termino_busqueda_actual}': {num_fcds} FCDs en Diccionario, resultando en {num_desc} filas en Descripciones.")

        elif origen_actual == OrigenResultados.DICCIONARIO_SIN_COINCIDENCIAS:
            self.resultados_actuales = resultados_df # Debería ser un DataFrame vacío
            self._actualizar_mensaje_barra_estado(f"El término '{termino_busqueda_actual}' no se encontró en el Diccionario.")
            # Ofrecer búsqueda directa en descripciones
            if messagebox.askyesno("Búsqueda Alternativa", 
                                   f"El término '{termino_busqueda_actual}' no fue encontrado en el Diccionario.\n\n"
                                   f"¿Desea buscar '{termino_busqueda_actual}' directamente en el archivo de Descripciones?"):
                self._try_except_wrapper(self._buscar_directo_en_descripciones_y_actualizar_ui, termino_busqueda_actual, columnas_df_desc_ref)
                return # La actualización de UI se hará en la función llamada
            # else: El usuario dijo no, se mostrará tabla de resultados vacía (ya está)

        elif origen_actual in [OrigenResultados.VIA_DICCIONARIO_SIN_RESULTADOS_DESC, 
                               OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS,
                               OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC,
                               OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC]:
            self.resultados_actuales = resultados_df # Debería ser DataFrame vacío
            num_fcds_intermedios = len(fcds_encontrados) if fcds_encontrados is not None else 0
            msg_fcds_intermedios = f"{num_fcds_intermedios} FCDs encontrados en Diccionario para '{termino_busqueda_actual}'"
            
            msg_detalle_falta_desc = ""
            if origen_actual == OrigenResultados.VIA_DICCIONARIO_UNIDAD_SIN_RESULTADOS_DESC:
                msg_detalle_falta_desc = "pero no se encontraron coincidencias numéricas o de unidad en Descripciones."
            elif origen_actual in [OrigenResultados.VIA_DICCIONARIO_SIN_TERMINOS_VALIDOS, OrigenResultados.VIA_DICCIONARIO_PURAMENTE_NEGATIVA_SIN_RESULTADOS_DESC]:
                 msg_detalle_falta_desc = "pero no se pudieron extraer términos válidos de ellos para buscar en Descripciones."
            else: # VIA_DICCIONARIO_SIN_RESULTADOS_DESC
                msg_detalle_falta_desc = "pero esto no produjo resultados en el archivo de Descripciones."

            self._actualizar_mensaje_barra_estado(f"{msg_fcds_intermedios}, {msg_detalle_falta_desc.replace('.','')} en Descripciones.")
            if messagebox.askyesno("Búsqueda Alternativa", 
                                   f"{msg_fcds_intermedios}, {msg_detalle_falta_desc}\n\n"
                                   f"¿Desea buscar '{termino_busqueda_actual}' directamente en el archivo de Descripciones?"):
                self._try_except_wrapper(self._buscar_directo_en_descripciones_y_actualizar_ui, termino_busqueda_actual, columnas_df_desc_ref)
                return # La actualización de UI se hará en la función llamada

        elif origen_actual == OrigenResultados.DIRECTO_DESCRIPCION_CON_RESULTADOS: # Ya es el resultado de una búsqueda directa
            self.resultados_actuales = resultados_df
            num_res_dir = len(resultados_df) if resultados_df is not None else 0
            self._actualizar_mensaje_barra_estado(f"Búsqueda directa de '{termino_busqueda_actual}': {num_res_dir} resultados encontrados.")
        
        elif origen_actual == OrigenResultados.DIRECTO_DESCRIPCION_VACIA: # Búsqueda directa sin resultados, o query vacía mostrando todo
            self.resultados_actuales = resultados_df
            num_filas_mostradas = len(resultados_df) if resultados_df is not None else 0
            if not termino_busqueda_actual.strip(): # Si la query era vacía, se muestran todas las descripciones
                self._actualizar_mensaje_barra_estado(f"Mostrando todas las descripciones ({num_filas_mostradas} filas).")
            else: # Búsqueda directa que no encontró nada
                self._actualizar_mensaje_barra_estado(f"Búsqueda directa de '{termino_busqueda_actual}': 0 resultados.")
                if num_filas_mostradas == 0 : # Confirmar que no hay resultados y la query no era vacía
                    messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_busqueda_actual}' en la búsqueda directa.")

        # Asegurar que resultados_actuales sea un DataFrame para la UI, incluso si es vacío
        if self.resultados_actuales is None:
            self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_ref)
        
        self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy() # Guardar para "Salvar Regla"
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, self.resultados_actuales)
        self._actualizar_estado_general_botones_y_controles()

    def _buscar_directo_en_descripciones_y_actualizar_ui(self, termino_ui_original: str, columnas_df_desc_referencia: List[str]):
        """ Realiza una búsqueda directa en descripciones y actualiza la UI. Se llama como alternativa. """
        self._actualizar_mensaje_barra_estado(f"Iniciando búsqueda directa de '{termino_ui_original}' en descripciones...")
        
        # Resetear resaltado de FCDs ya que es búsqueda directa
        self.indices_fcds_resaltados = None
        if self.motor.datos_diccionario is not None:
            # Actualizar la vista del diccionario para quitar cualquier resaltado previo
            columnas_preview_dicc_directo, _ = self.motor._obtener_nombres_columnas_busqueda_df(
                self.motor.datos_diccionario, self.motor.indices_columnas_busqueda_dic_preview, "diccionario_preview"
            )
            self.lbl_tabla_diccionario.config(text=f"Vista Previa Diccionario ({len(self.motor.datos_diccionario)} filas)")
            self._actualizar_tabla_treeview_ui(
                self.tabla_diccionario, self.motor.datos_diccionario, 
                limite_filas=100, columnas_a_mostrar=columnas_preview_dicc_directo, indices_a_resaltar=None
            )
        
        # Ejecutar búsqueda directa
        res_df_directo, orig_directo, _, _, msg_error_directo = self.motor.buscar(
            termino_busqueda_original=termino_ui_original, 
            buscar_via_diccionario_flag=False # Indicar búsqueda directa
        )
        
        self.origen_principal_resultados = orig_directo
        self.fcds_de_ultima_busqueda = None # No hay FCDs en búsqueda directa

        if msg_error_directo and (orig_directo.es_error_operacional or orig_directo.es_termino_invalido):
            messagebox.showerror("Error en Búsqueda Directa", f"Ocurrió un error: {msg_error_directo}")
            self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_referencia)
        else:
            self.resultados_actuales = res_df_directo

        num_resultados_directos = len(self.resultados_actuales) if self.resultados_actuales is not None else 0
        self._actualizar_mensaje_barra_estado(f"Búsqueda directa de '{termino_ui_original}': {num_resultados_directos} resultados encontrados.")
        
        if num_resultados_directos == 0 and orig_directo == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and termino_ui_original.strip():
            messagebox.showinfo("Información", f"No se encontraron resultados para '{termino_ui_original}' en la búsqueda directa.")
        
        # Asegurar DataFrame para la UI
        if self.resultados_actuales is None:
            self.resultados_actuales = pd.DataFrame(columns=columnas_df_desc_referencia)
            
        self.desc_finales_de_ultima_busqueda = self.resultados_actuales.copy()
        self._actualizar_tabla_treeview_ui(self.tabla_resultados, self.resultados_actuales)
        self._actualizar_estado_general_botones_y_controles()

    def _salvar_regla_actual_ui(self):
        """ Guarda metadatos de la búsqueda actual (no los datos en sí). """
        origen_actual_nombre = self.origen_principal_resultados.name

        # Validar si hay algo que salvar
        if not self.ultimo_termino_buscado and not \
           (self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and self.desc_finales_de_ultima_busqueda is not None):
            messagebox.showerror("Error al Salvar", "No hay una búsqueda activa o resultados para salvar.")
            return

        df_a_considerar_para_salvar: Optional[pd.DataFrame] = None
        tipo_datos_salvados = "DESCONOCIDO"

        # Determinar qué DataFrame y tipo de datos se están "salvando" (solo metadatos)
        if self.origen_principal_resultados.es_via_diccionario:
            if self.desc_finales_de_ultima_busqueda is not None and not self.desc_finales_de_ultima_busqueda.empty:
                df_a_considerar_para_salvar = self.desc_finales_de_ultima_busqueda
                tipo_datos_salvados = "DESC_VIA_DICC"
            elif self.fcds_de_ultima_busqueda is not None and not self.fcds_de_ultima_busqueda.empty:
                df_a_considerar_para_salvar = self.fcds_de_ultima_busqueda
                tipo_datos_salvados = "FCDS_DICC_SIN_DESC" # FCDs encontrados, pero no llevaron a resultados en desc
        
        elif self.origen_principal_resultados.es_directo_descripcion or \
             self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA:
            if self.desc_finales_de_ultima_busqueda is not None: # Siempre debería haber un DataFrame aquí, aunque sea vacío
                df_a_considerar_para_salvar = self.desc_finales_de_ultima_busqueda
                tipo_datos_salvados = "DESC_DIRECTA"
                # Caso especial: si la query era vacía y se mostraron todas las descripciones
                if self.origen_principal_resultados == OrigenResultados.DIRECTO_DESCRIPCION_VACIA and \
                   not (self.ultimo_termino_buscado or "").strip():
                    tipo_datos_salvados = "TODAS_LAS_DESCRIPCIONES"
        
        if df_a_considerar_para_salvar is not None:
            regla_info = {
                "termino_buscado": self.ultimo_termino_buscado or "N/A (Query vacía)",
                "origen_resultados": origen_actual_nombre,
                "tipo_datos_guardados": tipo_datos_salvados,
                "numero_filas_resultantes": len(df_a_considerar_para_salvar),
                "timestamp_guardado": pd.Timestamp.now().isoformat() # Fecha y hora del guardado
            }
            self.reglas_guardadas.append(regla_info)
            self._actualizar_mensaje_barra_estado(f"Búsqueda '{self.ultimo_termino_buscado}' registrada (metadatos).")
            messagebox.showinfo("Regla Salvada", f"Metadatos de la búsqueda para '{self.ultimo_termino_buscado}' han sido guardados en memoria.")
            logger.info(f"Regla guardada en memoria: {regla_info}")
        else:
            messagebox.showwarning("Nada que Salvar", "No hay datos de resultados claros asociados con la última búsqueda para salvar.")
        
        self._actualizar_estado_general_botones_y_controles() # El estado del botón de salvar no cambia aquí

    def _exportar_resultados_ui(self):
        """ Exporta los resultados actuales (de la tabla de descripciones) a un archivo Excel o CSV. """
        if self.resultados_actuales is None or self.resultados_actuales.empty:
            messagebox.showinfo("Exportar Resultados", "No hay resultados para exportar.")
            return

        # Sugerir un nombre de archivo con timestamp
        nombre_archivo_sugerido = f"resultados_busqueda_{pd.Timestamp.now():%Y%m%d_%H%M%S}"
        
        ruta_archivo_exportar = filedialog.asksaveasfilename(
            defaultextension=".xlsx", # Extensión por defecto
            filetypes=[("Archivo Excel", "*.xlsx"), ("Archivo CSV", "*.csv")],
            title="Guardar Resultados Como...",
            initialfile=nombre_archivo_sugerido
        )

        if not ruta_archivo_exportar: # Usuario canceló
            return

        try:
            if ruta_archivo_exportar.endswith(".xlsx"):
                self.resultados_actuales.to_excel(ruta_archivo_exportar, index=False)
            elif ruta_archivo_exportar.endswith(".csv"):
                self.resultados_actuales.to_csv(ruta_archivo_exportar, index=False, encoding='utf-8-sig') # utf-8-sig para mejor compatibilidad con Excel para CSVs
            else: # Debería ser manejado por defaultextension, pero por si acaso
                messagebox.showerror("Error de Formato", "Formato de archivo no soportado. Use .xlsx o .csv.")
                return
            
            messagebox.showinfo("Exportación Exitosa", f"Resultados exportados exitosamente a:\n{ruta_archivo_exportar}")
            self._actualizar_mensaje_barra_estado(f"Resultados exportados a {Path(ruta_archivo_exportar).name}")
        except Exception as e_export:
            logger.exception(f"Error al exportar resultados a '{ruta_archivo_exportar}'.")
            messagebox.showerror("Error de Exportación", f"No se pudieron exportar los resultados:\n{e_export}")

    def _actualizar_estado_botones_operadores(self):
        """ Habilita o deshabilita los botones de operadores según el contexto del texto de búsqueda. """
        # Si no hay archivos cargados, todos deshabilitados
        if self.motor.datos_diccionario is None and self.motor.datos_descripcion is None:
            self._deshabilitar_botones_operadores()
            return

        # Habilitar todos por defecto si hay archivos
        for btn in self.op_buttons.values():
            btn.config(state="normal")

        texto_actual = self.texto_busqueda_var.get()
        posicion_cursor = self.entrada_busqueda.index(tk.INSERT)
        
        # Carácter relevante antes del cursor (ignorando espacios al final de esa subcadena)
        subtexto_antes_cursor = texto_actual[:posicion_cursor].strip()
        ultimo_caracter_relevante = subtexto_antes_cursor[-1:] if subtexto_antes_cursor else ""

        # Lógica para deshabilitar según el último carácter
        operadores_logicos = ["+", "|"] # "/" ya no es OR
        operadores_comparacion_prefijo = [">", "<"]

        # No se puede empezar con operadores lógicos o relacionales de sufijo
        if not ultimo_caracter_relevante or ultimo_caracter_relevante in operadores_logicos + ["#", "<", ">", "=", "-"]:
            if self.op_buttons.get("+"): self.op_buttons["+"]["state"] = "disabled"
            if self.op_buttons.get("|"): self.op_buttons["|"]["state"] = "disabled"
            # No se puede poner ">=" o "<=" si ya hay un ">" o "<"
        
        # No se puede poner # si el último carácter no es espacio o inicio de query o operador lógico
        if ultimo_caracter_relevante and ultimo_caracter_relevante not in operadores_logicos + [" "]:
             if self.op_buttons.get("#"): self.op_buttons["#"]["state"] = "disabled"
        
        # Lógica para operadores de comparación y rango
        if ultimo_caracter_relevante in [">", "<", "="]: # Si ya se escribió un operador de comparación
            for op_key in operadores_comparacion_prefijo + ["=", "-"]: # Deshabilitar otros operadores de comparación y el de rango
                if self.op_buttons.get(op_key): self.op_buttons[op_key]["state"] = "disabled"
            # Deshabilitar >= si ya hay >
            if ultimo_caracter_relevante == ">" and self.op_buttons.get(">="): self.op_buttons[">="]["state"] = "disabled"
            # Deshabilitar <= si ya hay <
            if ultimo_caracter_relevante == "<" and self.op_buttons.get("<="): self.op_buttons["<="]["state"] = "disabled"
        
        # Si el último carácter es un dígito, no se pueden poner operadores de prefijo o #
        if ultimo_caracter_relevante.isdigit():
            for op_key_prefijo in operadores_comparacion_prefijo + ["=", "#"]:
                 if self.op_buttons.get(op_key_prefijo): self.op_buttons[op_key_prefijo]["state"] = "disabled"
            # Permitir "-" para rangos (ej. 10-)
        elif not ultimo_caracter_relevante or ultimo_caracter_relevante in [" ", "+", "|"]: # Si es inicio, o después de espacio u op lógico
            # No se puede poner un "-" de rango si no hay un número antes
            if self.op_buttons.get("-"): self.op_buttons["-"]["state"] = "disabled"

    def _insertar_operador_validado(self, operador_limpio: str):
        """ Inserta el operador en la entrada de búsqueda con espaciado adecuado. """
        operadores_con_espacio_alrededor = ["+", "|"] # Operadores que necesitan espacio antes y después
        
        texto_a_insertar: str
        if operador_limpio in operadores_con_espacio_alrededor:
            texto_a_insertar = f" {operador_limpio} "
        elif operador_limpio == "-": # Para rangos, sin espacio antes si viene de un número
            texto_a_insertar = f"{operador_limpio}" 
        elif operador_limpio in [">=", "<=", ">", "<", "="]: # Comparadores, usualmente sin espacio después si va un número
            texto_a_insertar = f"{operador_limpio}"
        elif operador_limpio == "#": # Negación, con espacio después
            texto_a_insertar = f"{operador_limpio} "
        else: # Caso genérico (no debería ocurrir con los botones actuales)
            texto_a_insertar = operador_limpio
            
        self.entrada_busqueda.insert(tk.INSERT, texto_a_insertar)
        self.entrada_busqueda.focus_set() # Devolver el foco a la entrada
        self._actualizar_estado_botones_operadores() # Re-evaluar estado de botones

    def _deshabilitar_botones_operadores(self):
        """ Deshabilita todos los botones de operadores. """
        for btn in self.op_buttons.values():
            btn.config(state="disabled")

    def on_closing_app(self):
        """ Maneja el evento de cierre de la ventana. """
        try:
            logger.info("Cerrando aplicación Buscador Avanzado...")
            self._guardar_configuracion_app() # Guardar configuración antes de salir
            self.destroy() # Cerrar la ventana de Tkinter
        except Exception as e: # Captura cualquier error durante el cierre
            func_name = "on_closing_app"
            error_type = type(e).__name__
            error_msg = str(e)
            tb_str = traceback.format_exc()
            
            logger.critical(f"Error durante el cierre en {func_name}: {error_type} - {error_msg}\n{tb_str}")
            print(f"--- TRACEBACK COMPLETO (cierre de app) ---\n{tb_str}") # Asegurar que se vea en consola
            self.destroy() # Intentar destruir de todas formas