import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import Optional

class ManejadorExcel:
    """
    Clase para manejar operaciones con archivos Excel.
    """

    @staticmethod
    def cargar_excel(ruta: str) -> Optional[pd.DataFrame]:
        try:
            return pd.read_excel(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar archivo:\n{e}")
            return None

    @staticmethod
    def comparar_dataframes(df1: pd.DataFrame, df2: pd.DataFrame) -> bool:
        try:
            return df1.equals(df2)
        except Exception as e:
            messagebox.showerror("Error", f"Error al comparar:\n{e}")
            return False

class MotorBusqueda:
    """
    Gestiona la lógica de búsqueda y manipulación de datos.
    """

    def __init__(self):
        self.datos_buscador: Optional[pd.DataFrame] = None
        self.datos_comparar: Optional[pd.DataFrame] = None
        self.archivo_actual: Optional[str] = None
        self.resultados: Optional[pd.DataFrame] = None

    def cargar_excel_buscador(self, ruta: str) -> bool:
        self.datos_buscador = ManejadorExcel.cargar_excel(ruta)
        self.archivo_actual = ruta if self.datos_buscador is not None else None
        return self.datos_buscador is not None

    def cargar_excel_comparar(self, ruta: str) -> bool:
        self.datos_comparar = ManejadorExcel.cargar_excel(ruta)
        return self.datos_comparar is not None

    def buscar(self, termino: str) -> Optional[pd.DataFrame]:
        if self.datos_buscador is None:
            messagebox.showwarning("Advertencia", "Primero cargue el archivo del buscador")
            return None

        termino = termino.strip().upper()
        if not termino:
            return self.datos_buscador.copy()

        # Primero buscar en datos_buscador
        resultados_buscador = self._buscar_en_dataframe(self.datos_buscador, termino)

        # Si no hay resultados o no hay datos para comparar, devolver los resultados del buscador
        if resultados_buscador.empty or self.datos_comparar is None:
            return resultados_buscador

        # Buscar coincidencias entre los resultados del buscador y datos_comparar
        # Convertir cada fila a texto para buscar coincidencias de palabras
        resultados_finales = []

        for _, fila_buscador in resultados_buscador.iterrows():
            texto_fila = ' '.join(fila_buscador.astype(str)).upper()

            # Buscar coincidencias en datos_comparar
            for _, fila_comparar in self.datos_comparar.iterrows():
                texto_comparar = ' '.join(fila_comparar.astype(str)).upper()

                # Si hay alguna palabra en común
                palabras_fila = set(texto_fila.split())
                palabras_comparar = set(texto_comparar.split())

                if palabras_fila.intersection(palabras_comparar):
                    resultados_finales.append(fila_comparar)

        if resultados_finales:
            return pd.DataFrame(resultados_finales)
        else:
            return pd.DataFrame()

    def _buscar_en_dataframe(self, df: pd.DataFrame, termino: str) -> pd.DataFrame:
        """Lógica de búsqueda optimizada"""
        try:
            if '+' in termino:
                palabras = [p.strip() for p in termino.split('+') if p]
                mascara = df.apply(
                    lambda fila: all(p in ' '.join(fila.astype(str)).upper() for p in palabras),
                    axis=1
                )
            elif '-' in termino:
                palabras = [p.strip() for p in termino.split('-') if p]
                mascara = df.apply(
                    lambda fila: any(p in ' '.join(fila.astype(str)).upper() for p in palabras),
                    axis=1
                )
            else:
                mascara = df.apply(
                    lambda fila: termino in ' '.join(fila.astype(str)).upper(),
                    axis=1
                )

            return df[mascara].copy()
        except Exception as e:
            messagebox.showerror("Error", f"Error en búsqueda: {e}")
            return pd.DataFrame()

class InterfazGrafica(tk.Tk):
    """
    Maneja la interfaz gráfica de la aplicación.
    """

    def __init__(self):
        super().__init__()
        self.title("Buscador Avanzado Optimizado")
        self.geometry("1000x800")
        self.motor = MotorBusqueda()

        # Configuración de la interfaz
        self._crear_widgets()
        self._configurar_grid()
        self._configurar_eventos()

    def _crear_widgets(self):
        # Marco de controles
        self.marco_controles = ttk.LabelFrame(self, text="Controles")

        # Botones de control
        self.btn_cargar = ttk.Button(
            self.marco_controles,
            text="Cargar Diccionario",
            command=self._cargar_buscador
        )

        self.btn_comparar = ttk.Button(
            self.marco_controles,
            text="Cargar Descripciones",
            command=self._cargar_excel_a_comparar,
            state="disabled"
        )

        # Entrada de búsqueda
        self.lbl_busqueda = ttk.Label(self.marco_controles, text="REGLAS a ensayar:")
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50)
        self.btn_buscar = ttk.Button(
            self.marco_controles,
            text="Buscar",
            command=self._ejecutar_busqueda,
            state="disabled"
        )

        # Botón de exportación
        self.btn_exportar = ttk.Button(
            self.marco_controles,
            text="Exportar REGLAS",
            command=self._exportar_resultados,
            state="disabled"
        )

        # Etiquetas para las tablas
        self.lbl_datos = ttk.Label(self, text="Datos cargados (Diccionario):")
        self.lbl_resultados = ttk.Label(self, text="Resultados de búsqueda:")

        # Tablas con scrollbars
        self.frame_tabla_principal = ttk.Frame(self)
        self.tabla_principal = ttk.Treeview(self.frame_tabla_principal)
        self.scrolly_principal = ttk.Scrollbar(self.frame_tabla_principal, orient="vertical",
                                             command=self.tabla_principal.yview)
        self.scrollx_principal = ttk.Scrollbar(self.frame_tabla_principal, orient="horizontal",
                                             command=self.tabla_principal.xview)
        self.tabla_principal.configure(yscrollcommand=self.scrolly_principal.set,
                                     xscrollcommand=self.scrollx_principal.set)

        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados)
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical",
                                              command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal",
                                              command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set,
                                      xscrollcommand=self.scrollx_resultados.set)

        # Barra de estado
        self.barra_estado = ttk.Label(self, text="Listo", relief=tk.SUNKEN, anchor=tk.W)

    def _configurar_grid(self):
        # Configuración principal
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(4, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Marco de controles
        self.marco_controles.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        self.marco_controles.grid_columnconfigure(1, weight=1)

        # Controles
        self.btn_cargar.grid(row=0, column=0, padx=5, pady=5)
        self.btn_comparar.grid(row=0, column=1, padx=5, pady=5)
        self.btn_exportar.grid(row=0, column=2, padx=5, pady=5)
        self.lbl_busqueda.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entrada_busqueda.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.btn_buscar.grid(row=1, column=2, padx=5, pady=5)

        # Etiquetas y tablas
        self.lbl_datos.grid(row=1, column=0, sticky="sw", padx=10, pady=(5, 0))
        self.lbl_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(10, 0))
        self.frame_tabla_principal.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Configuración interna de frames
        self.frame_tabla_principal.grid_rowconfigure(0, weight=1)
        self.frame_tabla_principal.grid_columnconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)

        # Posicionamiento de widgets en frames
        self.tabla_principal.grid(row=0, column=0, sticky="nsew")
        self.scrolly_principal.grid(row=0, column=1, sticky="ns")
        self.scrollx_principal.grid(row=1, column=0, sticky="ew")
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.scrolly_resultados.grid(row=0, column=1, sticky="ns")
        self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # Barra de estado
        self.barra_estado.grid(row=5, column=0, sticky="ew", padx=5, pady=5)

    def _configurar_eventos(self):
        self.bind("<Return>", lambda event: self._ejecutar_busqueda())
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], mostrar_limitado: bool = False):
        tabla.delete(*tabla.get_children())
        tabla["columns"] = ()

        if datos is None or datos.empty:
            self.barra_estado.config(text="No hay datos para mostrar")
            return

        columnas = list(datos.columns)
        tabla["columns"] = columnas
        tabla["show"] = "headings"

        for col in columnas:
            tabla.heading(col, text=str(col))
            max_width = len(str(col)) * 10
            for i, valor in enumerate(datos[col].astype(str)):
                if i > 100:
                    break
                max_width = max(max_width, len(str(valor)) * 7)
            tabla.column(col, width=min(max_width, 300), minwidth=50)

        if mostrar_limitado:
            datos_a_mostrar = datos.head(3)
            for idx, fila in datos_a_mostrar.iterrows():
                valores = [str(v) for v in fila.values]
                tabla.insert("", "end", values=valores)
            self.barra_estado.config(text=f"Mostrando las primeras {len(datos_a_mostrar)} de {len(datos)} coincidencias")
        else:
            for idx, fila in datos.iterrows():
                valores = [str(v) for v in fila.values]
                tabla.insert("", "end", values=valores)
            self.barra_estado.config(text=f"Mostrando {len(datos)} filas")

    def _cargar_buscador(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if not ruta:
            return

        self.barra_estado.config(text="Cargando diccionario...")
        self.update_idletasks()

        if self.motor.cargar_excel_buscador(ruta):
            self._actualizar_tabla(self.tabla_principal, self.motor.datos_buscador)
            self.title(f"Buscador - {ruta}")
            self.btn_comparar["state"] = "normal"
            self.btn_buscar["state"] = "normal"
            messagebox.showinfo("Éxito", f"Archivo cargado correctamente\nFilas: {len(self.motor.datos_buscador)}")
        else:
            self.barra_estado.config(text="Error al cargar diccionario")

    def _cargar_excel_a_comparar(self):
        if self.motor.datos_buscador is None:
            messagebox.showwarning("Advertencia", "Primero cargue un archivo con 'Cargar Diccionario'")
            return

        self.ruta_descripciones = filedialog.askopenfilename(
            title="Seleccionar archivo de descripciones",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )

        if not self.ruta_descripciones:
            return

        self.barra_estado.config(text="Cargando descripciones...")
        self.update_idletasks()

        if self.motor.cargar_excel_comparar(self.ruta_descripciones):
            self.df_comparar = self.motor.datos_comparar
            if self.df_comparar is not None:
                if self.motor.datos_buscador.equals(self.df_comparar):
                    messagebox.showinfo("Comparación", "Los archivos son idénticos.")
                else:
                    messagebox.showinfo("Comparación", "Los archivos son diferentes.")
                self._actualizar_tabla(self.tabla_resultados, self.df_comparar)
                self.barra_estado.config(text="Descripciones cargadas en Resultados de búsqueda")
        else:
            self.barra_estado.config(text="Error al cargar descripciones")

    def _ejecutar_busqueda(self):
        termino = self.entrada_busqueda.get()
        self.barra_estado.config(text="Buscando...")
        self.update_idletasks()

        if not termino.strip():
            messagebox.showwarning("Advertencia", "Debe ingresar un término de búsqueda válido.")
            # Corregido: Usar el DataFrame directamente en lugar de llamar al método
            self._actualizar_tabla(self.tabla_resultados, self.motor.datos_comparar)
            return None

        resultados = self.motor.buscar(termino)

        if resultados is not None and not resultados.empty:
            self._actualizar_tabla(self.tabla_resultados, resultados, mostrar_limitado=True)
            self.btn_exportar["state"] = "normal"
            self.barra_estado.config(text=f"Búsqueda completada: {len(resultados)} resultados")
        else:
            self.tabla_resultados.delete(*self.tabla_resultados.get_children())
            self.btn_exportar["state"] = "disabled"
            self.barra_estado.config(text="No se encontraron resultados")
        return resultados

    def _exportar_resultados(self):
        resultados = self._ejecutar_busqueda()
        if resultados is None or resultados.empty:
            messagebox.showwarning("Advertencia", "No hay resultados para exportar")
            return

        ruta = filedialog.asksaveasfilename(
            title="Guardar resultados",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("CSV", "*.csv")]
        )

        if not ruta:
            return

        self.barra_estado.config(text="Exportando REGLAS...")
        self.update_idletasks()

        try:
            extension = ruta.split('.')[-1].lower()

            # Añadir información de depuración
            print(f"DataFrame info:")
            print(resultados.info())
            print(f"DataFrame shape: {resultados.shape}")
            print(f"DataFrame dtypes: {resultados.dtypes}")
            print(f"Primeras 5 filas: {resultados.head()}")

            if extension == 'csv':
                resultados.to_csv(ruta, index=False)
            elif extension in ['xls', 'xlsx']:
                with pd.ExcelWriter(ruta, engine='xlsxwriter') as writer:
                    resultados.to_excel(writer, index=False)
                    # Asegurarse de que se guarde
                    print("Excel creado correctamente")

            messagebox.showinfo("Éxito", f"Archivo exportado correctamente a:\n{ruta}")
            self.barra_estado.config(text=f"REGLAS exportadas a {ruta}")
        except Exception as e:
            # Mostrar información más detallada del error
            import traceback
            error_detallado = traceback.format_exc()
            print(f"Error detallado: {error_detallado}")
            messagebox.showerror("Error", f"Error al exportar:\n{e}\n\nDetalles: {error_detallado}")
            self.barra_estado.config(text="Error al exportar REGLAS")

if __name__ == "__main__":
    app = InterfazGrafica()
    app.mainloop()