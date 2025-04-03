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
            # Usamos equals en lugar de compare para una comparación más robusta
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
            return self.datos_buscador  # Devuelve todos los datos si no hay término de búsqueda

        # Lógica de operadores mejorada
        if '+' in termino:
            palabras = [p.strip() for p in termino.split('+') if p]
            mascara = self.datos_buscador.apply(
                lambda fila: all(p in ' '.join(fila.astype(str)).upper() for p in palabras),
                axis=1
            )
        elif '-' in termino:
            palabras = [p.strip() for p in termino.split('-') if p]
            mascara = self.datos_buscador.apply(
                lambda fila: any(p in ' '.join(fila.astype(str)).upper() for p in palabras),
                axis=1
            )
        else:
            mascara = self.datos_buscador.apply(
                lambda fila: termino in ' '.join(fila.astype(str)).upper(),
                axis=1
            )

        self.resultados = self.datos_buscador[mascara]
        return self.resultados if not self.resultados.empty else None


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
            text="Cargar Excel Buscador",
            command=self._cargar_buscador
        )

        self.btn_comparar = ttk.Button(
            self.marco_controles,
            text="Cargar Excel a Comparar",
            command=self._cargar_excel_a_comparar,
            state="disabled"  # Inicialmente deshabilitado
        )

        # Entrada de búsqueda y botón
        self.lbl_busqueda = ttk.Label(self.marco_controles, text="Término/s de búsqueda:")
        self.entrada_busqueda = ttk.Entry(self.marco_controles, width=50)

        self.btn_buscar = ttk.Button(
            self.marco_controles,
            text="Buscar",
            command=self._ejecutar_busqueda,
            state="disabled"  # Inicialmente deshabilitado
        )

        # Botón de exportación
        self.btn_exportar = ttk.Button(
            self.marco_controles,
            text="Exportar Resultados",
            command=self._exportar_resultados,
            state="disabled"  # Inicialmente deshabilitado
        )

        # Etiquetas para las tablas
        self.lbl_datos = ttk.Label(self, text="Datos cargados (Buscador):")
        self.lbl_resultados = ttk.Label(self, text="Resultados de búsqueda:")

        # Tablas con scrollbars
        # Tabla principal (datos del buscador)
        self.frame_tabla_principal = ttk.Frame(self)
        self.tabla_principal = ttk.Treeview(self.frame_tabla_principal)
        self.scrolly_principal = ttk.Scrollbar(self.frame_tabla_principal, orient="vertical",
                                               command=self.tabla_principal.yview)
        self.scrollx_principal = ttk.Scrollbar(self.frame_tabla_principal, orient="horizontal",
                                               command=self.tabla_principal.xview)
        self.tabla_principal.configure(yscrollcommand=self.scrolly_principal.set,
                                       xscrollcommand=self.scrollx_principal.set)

        # Tabla de resultados (para comparar y mostrar resultados de búsqueda)
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
        # Configuración de filas y columnas de la ventana principal
        self.grid_rowconfigure(2, weight=1)  # Tabla principal
        self.grid_rowconfigure(4, weight=1)  # Tabla resultados
        self.grid_columnconfigure(0, weight=1)

        # Marco de controles
        self.marco_controles.grid(row=0, column=0, sticky="ew", padx=10, pady=5)

        # Configuración de columnas dentro del marco de controles
        self.marco_controles.grid_columnconfigure(1, weight=1)

        # Controles en la primera fila
        self.btn_cargar.grid(row=0, column=0, padx=5, pady=5)
        self.btn_comparar.grid(row=0, column=1, padx=5, pady=5)
        self.btn_exportar.grid(row=0, column=2, padx=5, pady=5)

        # Controles en la segunda fila
        self.lbl_busqueda.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entrada_busqueda.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.btn_buscar.grid(row=1, column=2, padx=5, pady=5)

        # Etiquetas de las tablas
        self.lbl_datos.grid(row=1, column=0, sticky="sw", padx=10, pady=(5, 0))
        self.lbl_resultados.grid(row=3, column=0, sticky="sw", padx=10, pady=(10, 0))

        # Configuración de frames de tablas
        self.frame_tabla_principal.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.frame_tabla_resultados.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Configuración de frames de tablas
        self.frame_tabla_principal.grid_rowconfigure(0, weight=1)
        self.frame_tabla_principal.grid_columnconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_rowconfigure(0, weight=1)
        self.frame_tabla_resultados.grid_columnconfigure(0, weight=1)

        # Colocación de tablas y scrollbars
        self.tabla_principal.grid(row=0, column=0, sticky="nsew")
        self.scrolly_principal.grid(row=0, column=1, sticky="ns")
        self.scrollx_principal.grid(row=1, column=0, sticky="ew")

        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.scrolly_resultados.grid(row=0, column=1, sticky="ns")
        self.scrollx_resultados.grid(row=1, column=0, sticky="ew")

        # Barra de estado
        self.barra_estado.grid(row=5, column=0, sticky="ew", padx=5, pady=5)

    def _configurar_eventos(self):
        # Configurar eventos adicionales como atajos de teclado
        self.bind("<Return>", lambda event: self._ejecutar_busqueda())
        self.entrada_busqueda.bind("<Return>", lambda event: self._ejecutar_busqueda())

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: Optional[pd.DataFrame], mostrar_limitado: bool = False):
        # Limpiar tabla existente
        tabla.delete(*tabla.get_children())

        # Remover columnas existentes
        tabla["columns"] = ()

        if datos is None or datos.empty:
            self.barra_estado.config(text="No hay datos para mostrar")
            return

        # Configurar columnas nuevas
        columnas = list(datos.columns)
        tabla["columns"] = columnas
        tabla["show"] = "headings"  # Oculta la columna de ID

        # Configurar encabezados
        for col in columnas:
            tabla.heading(col, text=str(col))
            # Estimar ancho basado en datos
            max_width = len(str(col)) * 10  # Ancho mínimo basado en el nombre de la columna

            # Revisar los primeros 100 valores para estimar ancho
            for i, valor in enumerate(datos[col].astype(str)):
                if i > 100:  # Limitar para no revisar toda la columna
                    break
                max_width = max(max_width, len(str(valor)) * 7)

            tabla.column(col, width=min(max_width, 300), minwidth=50)  # Limitar ancho máximo

        # Insertar filas
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
        if ruta:
            self.motor.cargar_excel_buscador(ruta)
            if self.motor.datos_buscador is not None:
                self.btn_comparar["state"] = "normal"  # Habilitar el botón

        self.barra_estado.config(text="Cargando archivo del buscador...")
        self.update_idletasks()  # Actualizar la interfaz para mostrar el mensaje

        if self.motor.datos_buscador is not None:
            self._actualizar_tabla(self.tabla_principal, self.motor.datos_buscador)
            self.title(f"Buscador - {ruta}")
            self._habilitar_componentes()
            messagebox.showinfo("Éxito", f"Archivo del buscador cargado correctamente\nFilas: {len(self.motor.datos_buscador)}")
        else:
            self.barra_estado.config(text="Error al cargar archivo del buscador")

    def _cargar_excel_a_comparar(self):
        if self.motor.datos_buscador is None:
            messagebox.showwarning("Advertencia", "Primero cargue un archivo con 'Cargar Excel Buscador'")
            return

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel a comparar",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )

        if not ruta:
            return

        self.barra_estado.config(text="Cargando archivo a comparar...")
        self.update_idletasks()

        if self.motor.cargar_excel_comparar(ruta):
            df_comparar = self.motor.datos_comparar
            if df_comparar is not None:
                # Verificar si el archivo a comparar es el mismo que el del buscador
                if self.motor.datos_buscador.equals(df_comparar):
                    messagebox.showinfo("Comparación", "Los archivos son idénticos.")
                    self._actualizar_tabla(self.tabla_resultados, df_comparar) # Mostrar completo si son iguales
                    self.btn_buscar.config(state="normal")
                    self.barra_estado.config(text="Archivos idénticos cargados en Resultados de búsqueda")
                else:
                    messagebox.showinfo("Comparación", "Los archivos son diferentes. El archivo a comparar se ha cargado en 'Resultados de búsqueda'.")
                    self._actualizar_tabla(self.tabla_resultados, df_comparar) # Mostrar completo si son diferentes
                    self.btn_buscar.config(state="normal")
                    self.barra_estado.config(text="Archivo a comparar cargado en Resultados de búsqueda")
        else:
            self.barra_estado.config(text="Error al cargar archivo a comparar")

    def _ejecutar_busqueda(self):
        termino = self.entrada_busqueda.get()

        self.barra_estado.config(text="Buscando...")
        self.update_idletasks()  # Actualizar la interfaz

        resultados = self.motor.buscar(termino)

        if resultados is not None:
            self._actualizar_tabla(self.tabla_resultados, resultados, mostrar_limitado=True)
            self.btn_exportar.config(state="normal")
            self.barra_estado.config(text=f"Búsqueda completada: {len(resultados)} resultados (mostrando los primeros 3)")
        else:
            self.tabla_resultados.delete(*self.tabla_resultados.get_children())
            self.btn_exportar.config(state="disabled")
            messagebox.showinfo("Información", "No se encontraron resultados.")
            self.barra_estado.config(text="No se encontraron resultados")

    def _exportar_resultados(self):
        if self.motor.resultados is None or self.motor.resultados.empty:
            messagebox.showwarning("Advertencia", "No hay resultados para exportar")
            return

        ruta = filedialog.asksaveasfilename(
            title="Guardar resultados",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("CSV", "*.csv")]
        )

        if not ruta:
            return

        self.barra_estado.config(text="Exportando resultados...")
        self.update_idletasks()

        try:
            # Elegir el método de exportación según la extensión
            extension = ruta.split('.')[-1].lower()
            if extension == 'csv':
                self.motor.resultados.to_csv(ruta, index=False)
            elif extension in ['xls', 'xlsx']:
                writer = pd.ExcelWriter(ruta, engine='xlsxwriter')
                self.motor.resultados.to_excel(writer, index=False)
                writer.close()

            messagebox.showinfo("Éxito", f"Archivo exportado correctamente a:\n{ruta}")
            self.barra_estado.config(text=f"Resultados exportados a {ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar:\n{e}")
            self.barra_estado.config(text="Error al exportar resultados")

    def _habilitar_componentes(self):
        """Habilita los componentes cuando se ha cargado un archivo"""
        self.btn_comparar.config(state="normal")
        # El botón buscar se habilita después de cargar el archivo para comparar
        # El botón exportar se habilita solo cuando hay resultados


if __name__ == "__main__":
    app = InterfazGrafica()
    app.mainloop()