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
        self.datos: Optional[pd.DataFrame] = None
        self.archivo_actual: Optional[str] = None
        self.resultados: Optional[pd.DataFrame] = None

    def cargar_excel(self, ruta: str) -> bool:
        self.datos = ManejadorExcel.cargar_excel(ruta)
        self.archivo_actual = ruta if self.datos is not None else None
        return self.datos is not None

    def buscar(self, termino: str) -> Optional[pd.DataFrame]:
        if self.datos is None:
            messagebox.showwarning("Advertencia", "Primero cargue un archivo")
            return None

        termino = termino.strip().upper()
        if not termino:
            return self.datos  # Devuelve todos los datos si no hay término de búsqueda

        # Lógica de operadores mejorada
        if '+' in termino:
            palabras = [p.strip() for p in termino.split('+') if p]
            mascara = self.datos.apply(
                lambda fila: all(p in ' '.join(fila.astype(str)).upper() for p in palabras),
                axis=1
            )
        elif '-' in termino:
            palabras = [p.strip() for p in termino.split('-') if p]
            mascara = self.datos.apply(
                lambda fila: any(p in ' '.join(fila.astype(str)).upper() for p in palabras),
                axis=1
            )
        else:
            mascara = self.datos.apply(
                lambda fila: termino in ' '.join(fila.astype(str)).upper(),
                axis=1
            )

        self.resultados = self.datos[mascara]
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
            text="Cargar Excel Principal",
            command=self._cargar_principal
        )
        
        self.btn_comparar = ttk.Button(
            self.marco_controles,
            text="Cargar Excel a Comparar",
            command=self._comparar_archivos,
            state="disabled"  # Inicialmente deshabilitado
        )
        
        # Entrada de búsqueda y botón
        self.lbl_busqueda = ttk.Label(self.marco_controles, text="Término de búsqueda:")
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
        self.lbl_datos = ttk.Label(self, text="Datos cargados:")
        self.lbl_resultados = ttk.Label(self, text="Resultados de búsqueda:")
        
        # Tablas con scrollbars
        # Tabla principal
        self.frame_tabla_principal = ttk.Frame(self)
        self.tabla_principal = ttk.Treeview(self.frame_tabla_principal)
        self.scrolly_principal = ttk.Scrollbar(self.frame_tabla_principal, orient="vertical", command=self.tabla_principal.yview)
        self.scrollx_principal = ttk.Scrollbar(self.frame_tabla_principal, orient="horizontal", command=self.tabla_principal.xview)
        self.tabla_principal.configure(yscrollcommand=self.scrolly_principal.set, xscrollcommand=self.scrollx_principal.set)
        
        # Tabla de resultados
        self.frame_tabla_resultados = ttk.Frame(self)
        self.tabla_resultados = ttk.Treeview(self.frame_tabla_resultados)
        self.scrolly_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="vertical", command=self.tabla_resultados.yview)
        self.scrollx_resultados = ttk.Scrollbar(self.frame_tabla_resultados, orient="horizontal", command=self.tabla_resultados.xview)
        self.tabla_resultados.configure(yscrollcommand=self.scrolly_resultados.set, xscrollcommand=self.scrollx_resultados.set)
        
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

    def _actualizar_tabla(self, tabla: ttk.Treeview, datos: pd.DataFrame):
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
        for idx, fila in datos.iterrows():
            valores = [str(v) for v in fila.values]
            tabla.insert("", "end", values=valores)
        
        # Actualizar barra de estado
        self.barra_estado.config(text=f"Mostrando {len(datos)} filas")

    def _cargar_principal(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if not ruta:
            return
            
        self.barra_estado.config(text="Cargando archivo...")
        self.update_idletasks()  # Actualizar la interfaz para mostrar el mensaje
        
        if self.motor.cargar_excel(ruta):
            self._actualizar_tabla(self.tabla_principal, self.motor.datos)
            self.title(f"Buscador - {ruta}")
            self._habilitar_componentes()
            messagebox.showinfo("Éxito", f"Archivo cargado correctamente\nFilas: {len(self.motor.datos)}")
        else:
            self.barra_estado.config(text="Error al cargar archivo")

    def _ejecutar_busqueda(self):
        termino = self.entrada_busqueda.get()
        
        self.barra_estado.config(text="Buscando...")
        self.update_idletasks()  # Actualizar la interfaz
        
        resultados = self.motor.buscar(termino)
        
        if resultados is not None:
            self._actualizar_tabla(self.tabla_resultados, resultados)
            self.btn_exportar.config(state="normal")
            self.barra_estado.config(text=f"Búsqueda completada: {len(resultados)} resultados")
        else:
            self.tabla_resultados.delete(*self.tabla_resultados.get_children())
            self.btn_exportar.config(state="disabled")
            self.barra_estado.config(text="No se encontraron resultados")

    def _comparar_archivos(self):
        if self.motor.datos is None:
            messagebox.showwarning("Advertencia", "Primero cargue un archivo principal")
            return
            
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo para comparar",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if not ruta:
            return
            
        self.barra_estado.config(text="Comparando archivos...")
        self.update_idletasks()
        
        df = ManejadorExcel.cargar_excel(ruta)
        if df is not None:
            # Verificar que tienen las mismas columnas
            if set(self.motor.datos.columns) != set(df.columns):
                messagebox.showinfo("Comparación", "Los archivos tienen columnas diferentes")
                self.barra_estado.config(text="Comparación finalizada - Columnas diferentes")
                return
                
            if ManejadorExcel.comparar_dataframes(self.motor.datos, df):
                messagebox.showinfo("Comparación", "Los archivos son idénticos")
                self.barra_estado.config(text="Comparación finalizada - Archivos idénticos")
            else:
                messagebox.showinfo("Comparación", "Los archivos son diferentes")
                self.barra_estado.config(text="Comparación finalizada - Archivos diferentes")

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
            else:
                self.motor.resultados.to_excel(ruta, index=False)
                
            messagebox.showinfo("Éxito", f"Archivo exportado correctamente a:\n{ruta}")
            self.barra_estado.config(text=f"Resultados exportados a {ruta}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar:\n{e}")
            self.barra_estado.config(text="Error al exportar resultados")

    def _habilitar_componentes(self):
        """Habilita los componentes cuando se ha cargado un archivo"""
        self.btn_comparar.config(state="normal")
        self.btn_buscar.config(state="normal")
        # El botón exportar se habilita solo cuando hay resultados


if __name__ == "__main__":
    app = InterfazGrafica()
    app.mainloop()