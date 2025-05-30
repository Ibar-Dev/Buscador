# -*- coding: utf-8 -*-
# main.py (Punto de entrada principal de la aplicación)

import logging
from pathlib import Path
import tkinter as tk # Para el messagebox en la verificación de dependencias
from tkinter import messagebox
import traceback # Para el bloque __main__ try-except final

# Importación desde tu paquete de aplicación
from buscador_app.gui.interfaz_grafica import InterfazGrafica
from typing import List # Necesario para dependencias_faltantes_main

# --- Punto de Entrada Principal de la Aplicación ---
if __name__ == "__main__":
    LOG_FILE_NAME = "Buscador_Avanzado_App_v1.10.3_Mod.log"
    # Configuración básica del logging
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(funcName)s() - %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE_NAME, encoding="utf-8", mode="w"),
            logging.StreamHandler()
        ]
    )
    root_logger = logging.getLogger()
    root_logger.info(f"--- Iniciando Buscador Avanzado v1.10.3 Mod (Modularizado) (Script: {Path(__file__).name}) ---")
    root_logger.info(f"Logs siendo guardados en: {Path(LOG_FILE_NAME).resolve()}")

    # Verificación de dependencias
    dependencias_faltantes_main: List[str] = []
    try: import pandas as pd_check_main; root_logger.info(f"Pandas: {pd_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("pandas")
    try: import openpyxl as opxl_check_main; root_logger.info(f"openpyxl: {opxl_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("openpyxl")
    try: import numpy as np_check_main; root_logger.info(f"Numpy: {np_check_main.__version__}")
    except ImportError: dependencias_faltantes_main.append("numpy")
    try: import xlrd as xlrd_check_main; root_logger.info(f"xlrd: {xlrd_check_main.__version__}")
    except ImportError: root_logger.warning("xlrd no encontrado. Carga de .xls antiguos podría fallar.")

    if dependencias_faltantes_main:
        mensaje_error_deps_main = (f"Faltan dependencias críticas: {', '.join(dependencias_faltantes_main)}.\nInstale con: pip install {' '.join(dependencias_faltantes_main)}")
        root_logger.critical(mensaje_error_deps_main)
        try:
            root_error_tk_main = tk.Tk(); root_error_tk_main.withdraw()
            messagebox.showerror("Dependencias Faltantes", mensaje_error_deps_main); root_error_tk_main.destroy()
        except Exception as e_tk_dep_main: print(f"ERROR CRITICO (Error al mostrar mensaje Tkinter: {e_tk_dep_main}): {mensaje_error_deps_main}")
        exit(1)
    
    # Inicia la aplicación
    try: 
        app=InterfazGrafica()
        app.mainloop()
    except Exception as e_main_app_exc:
        root_logger.critical("Error fatal no controlado en la aplicación principal:", exc_info=True)
        tb_str_fatal = traceback.format_exc()
        print(f"--- TRACEBACK FATAL (desde bloque __main__) ---\n{tb_str_fatal}")
        try:
            root_fatal_tk_main = tk.Tk(); root_fatal_tk_main.withdraw()
            messagebox.showerror("Error Fatal Inesperado", f"Error crítico: {e_main_app_exc}\nConsulte '{LOG_FILE_NAME}' y la consola."); root_fatal_tk_main.destroy()
        except: print(f"ERROR FATAL: {e_main_app_exc}. Revise '{LOG_FILE_NAME}'.")
    finally: 
        root_logger.info(f"--- Finalizando Buscador ---")