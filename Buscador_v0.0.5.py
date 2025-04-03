from httpx import delete
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional, List, Any




class ManejadorExcel:
    """
    Clase para manejar todas las operaciones relacionadas con archivos Excel.
    """

    @staticmethod
    def cargar_excel(ruta_archivo: str) -> Optional[pd.DataFrame]:
        """
        Carga un archivo Excel y devuelve un DataFrame.

        Args:
            ruta_archivo (str): La ruta al archivo Excel.

        Returns:
            Optional[pd.DataFrame]: El DataFrame si se carga correctamente, None en caso contrario.
        """
        try:
            return pd.read_excel(ruta_archivo)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")
            return None

    @staticmethod
    def comparar_dataframes(df1: pd.DataFrame, df2: pd.DataFrame) -> bool:
        """
        Compara dos DataFrames y devuelve True si son iguales, False en caso contrario.

        Args:
            df1 (pd.DataFrame): El primer DataFrame.
            df2 (pd.DataFrame): El segundo DataFrame.

        Returns:
            bool: True si son iguales, False en caso contrario.
        """
        try:
            return df1.equals(df2)
        except Exception as e:
            messagebox.showerror("Error", f"Error al comparar archivos: {e}")
            return False




class MotorBusqueda:
    """
    Clase que implementa el motor de búsqueda para realizar operaciones en DataFrames.
    """

    def __init__(self):
        """
        Inicializa el motor de búsqueda con un DataFrame vacío.
        """
        self.datos: Optional[pd.DataFrame] = None
        self.archivo_excel: Optional[str] = None
        self.resultado_filtrado: Optional[pd.DataFrame] = None
        self.manejador_excel = ManejadorExcel()  # Agregamos una instancia de ManejadorExcel

    def buscar_en_dataframe(self, termino_busqueda: str, max_resultados: int = 3) -> Optional[pd.DataFrame]:
        """
        Busca un término en las columnas especificadas del DataFrame.

        Args:
            termino_busqueda (str): El término a buscar.
            max_resultados (int): El número máximo de resultados a devolver.

        Returns:
            Optional[pd.DataFrame]: Un DataFrame con los resultados, o None si no hay datos o hay un error.
        """
        if self.datos is None:
            messagebox.showwarning("Advertencia", "No hay datos cargados")
            return None

        try:
            mascara = self.datos.loc[:, self.datos.columns].apply(
                lambda fila: fila.astype(str).str.contains(termino_busqueda, case=False, na=False).any(), axis=1
            )
            resultados = self.datos[mascara].head(max_resultados)
            if resultados.empty:
                messagebox.showinfo("Información", "No se encontraron coincidencias.")
                return None
            return resultados
        except Exception as e:
            messagebox.showerror("Error", f"Error en la búsqueda: {e}")
            return None

    def cargar_excel(self) -> Optional[pd.DataFrame]:
        """
        Abre un diálogo para seleccionar un archivo Excel y lo carga usando ManejadorExcel.

        Returns:
            Optional[pd.DataFrame]: El DataFrame si se carga correctamente, None en caso contrario.
        """
        self.archivo_excel = filedialog.askopenfilename(
            title="Seleccione un archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if self.archivo_excel:
            self.datos = self.manejador_excel.cargar_excel(self.archivo_excel)  # Usamos ManejadorExcel
            return self.datos
        else:
            return None




# class reglas(MotorBusqueda):
#     def __init__(self):
#         super().__init__()
#         self.and_mas = None
#         self.or_barra = None

#     def _tambien_debe_estar(self, palabra_añadida):
#         while not palabra_añadida:
#             continue
#         try:
#             pass
#         except Exception as e:
#             messagebox.showerror('Error', f'No se puede ejecutar: {e}')




class Aplicacion(tk.Tk):
    """
    Clase principal de la aplicación con interfaz gráfica.
    """

    def __init__(self):
        """
        Inicializa la aplicación, crea la ventana principal y configura la interfaz.
        """
        super().__init__()
        self.title("Buscador Avanzado")
        self.geometry("800x600")

        # Instancias de las clases de negocio
        self.manejador_excel = ManejadorExcel()
        self.motor_busqueda = MotorBusqueda()

        # Botones para tenerlos localizados
        self.boton_cargar_excel = None
        self.boton_cargar_excel_secundario = None
        self.boton_busqueda = None
        self.boton_crear_excel_resultados = None

        # Crear la interfaz
        self.crear_interfaz()

    def crear_interfaz(self):
        """
        Crea todos los elementos de la interfaz gráfica.
        """

        # Frame principal
        self.marco_principal = tk.Frame(self, padx=10, pady=10)
        self.marco_principal.grid(row=0, column=0, sticky="nsew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Sección de carga de archivos
        self.marco_archivos = tk.LabelFrame(self.marco_principal, text="Carga de Archivos", padx=5, pady=5)
        self.marco_archivos.grid(row=0, column=0, sticky="ew", pady=5)
        self.marco_principal.grid_columnconfigure(0, weight=1)

        self.boton_cargar_excel = tk.Button(
            self.marco_archivos,
            text="Cargar Excel Principal",
            command=self.cargar_archivo_principal
        )
        self.boton_cargar_excel.grid(row=0, column=0, padx=5, sticky="w")

        self.boton_cargar_excel_secundario = tk.Button(
            self.marco_archivos,
            text="Cargar Excel para Comparar",
            command=self.cargar_archivo_comparacion
        )
        self.boton_cargar_excel_secundario.grid(row=0, column=1, padx=5, sticky="w")

        # Sección de búsqueda
        self.marco_busqueda = tk.LabelFrame(self.marco_principal, text="Búsqueda", padx=5, pady=5)
        self.marco_busqueda.grid(row=1, column=0, sticky="ew", pady=5)

        self.entrada_busqueda = tk.Entry(self.marco_busqueda, width=50)
        self.entrada_busqueda.grid(row=0, column=0, padx=5, sticky="ew")
        self.marco_busqueda.grid_columnconfigure(0, weight=1)

        self.boton_busqueda = tk.Button(
            self.marco_busqueda,
            text="Buscar",
            command=self.ejecutar_busqueda
        )
        self.boton_busqueda.grid(row=0, column=1, padx=5, sticky="e")

        # Sección de resultados
        self.marco_resultados = tk.LabelFrame(self.marco_principal, text="Resultados", padx=5, pady=5)
        self.marco_resultados.grid(row=2, column=0, sticky="nsew", pady=5)
        self.marco_principal.grid_rowconfigure(2, weight=1)

        self.tabla_resultados = ttk.Treeview(self.marco_resultados)
        self.tabla_resultados.grid(row=0, column=0, sticky="nsew")
        self.marco_resultados.grid_rowconfigure(0, weight=1)
        self.marco_resultados.grid_columnconfigure(0, weight=1)

        # Sección de acciones
        self.marco_acciones = tk.Frame(self.marco_principal)
        self.marco_acciones.grid(row=3, column=0, sticky="ew", pady=5)

        self.boton_crear_excel_resultados = tk.Button(
            self.marco_acciones,
            text="Exportar Resultados",
            command=self.crear_excel_resultados
        )
        self.boton_crear_excel_resultados.grid(row=0, column=1, padx=5, sticky="e")
        self.marco_acciones.grid_columnconfigure(0, weight=1)

        # Marco para ver el título del Excel
        self.titulo_excel = tk.Label(self.marco_resultados, text="")
        self.titulo_excel.grid(row=1, column=0, sticky="ew")

        self._ocultar_innecesarios()

    def cargar_archivo_principal(self):
        """
        Carga el archivo Excel principal utilizando el motor de búsqueda y muestra los resultados.
        """
        if self.motor_busqueda.cargar_excel() is not None:
            if self.motor_busqueda.archivo_excel:
                self.titulo_excel.config(text=f"Archivo cargado: {self.motor_busqueda.archivo_excel}")
            else:
                self.titulo_excel.config(text="")
            # Mostrar los datos cargados en la tabla de resultados
            self.mostrar_resultados(self.motor_busqueda.datos)
            self._mostrar_necesarios()
        else:
            self.titulo_excel.config(text="")

        # Forzar actualización de la interfaz
        self.update_idletasks()

    def cargar_archivo_comparacion(self):
        """
        Carga un archivo Excel para comparar con el principal.
        """
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta_archivo:
            df2 = self.manejador_excel.cargar_excel(ruta_archivo)
            if df2 is not None and self.motor_busqueda.datos is not None:
                if self.manejador_excel.comparar_dataframes(self.motor_busqueda.datos, df2):
                    messagebox.showinfo("Comparación", "Los archivos son idénticos")
                else:
                    messagebox.showinfo("Comparación", "Los archivos son diferentes")

    def ejecutar_busqueda(self):
        """
        Ejecuta la búsqueda según el término ingresado y muestra los resultados.
        """
        termino_busqueda = self.entrada_busqueda.get().strip().upper()
        if termino_busqueda and self.motor_busqueda.datos is not None:
            if '+' in termino_busqueda:  # Búsqueda con "AND"
                palabras = [palabra.strip() for palabra in termino_busqueda.split('+') if palabra.strip()]
                mascara = self.motor_busqueda.datos.loc[:, self.motor_busqueda.datos.columns].apply(
                    lambda fila: all(palabra in ' '.join(fila.astype(str)).upper() for palabra in palabras), axis=1
                )
            elif '-' in termino_busqueda:  # Búsqueda con "OR"
                palabras = [palabra.strip() for palabra in termino_busqueda.split('-') if palabra.strip()]
                mascara = self.motor_busqueda.datos.loc[:, self.motor_busqueda.datos.columns].apply(
                    lambda fila: any(palabra in ' '.join(fila.astype(str)).upper() for palabra in palabras), axis=1
                )
            else:  # Búsqueda simple
                mascara = self.motor_busqueda.datos.loc[:, self.motor_busqueda.datos.columns].apply(
                    lambda fila: termino_busqueda in ' '.join(fila.astype(str)).upper(), axis=1
                )

            self.resultado_filtrado = self.motor_busqueda.datos[mascara]
            if not self.resultado_filtrado.empty:
                self.mostrar_resultados(None)  # Limpia los resultados previos
                self.mostrar_resultados(self.resultado_filtrado)
                return self.resultado_filtrado
            else:
                messagebox.showinfo("Información", "No se encontraron resultados.")
                return None
        else:
            messagebox.showwarning("Advertencia", "Por favor, ingrese un término de búsqueda.")

    def mostrar_resultados(self, resultados: Optional[pd.DataFrame]):
        """
        Muestra los resultados en el Treeview.

        Args:
            resultados (Optional[pd.DataFrame]): El DataFrame con los resultados, o None si no hay resultados.
        """

        # Limpia el Treeview antes de mostrar nuevos resultados
        for item in self.tabla_resultados.get_children():
            self.tabla_resultados.delete(item)

        if resultados is not None and not resultados.empty:
            # Configura los encabezados de las columnas
            self.tabla_resultados["columns"] = list(resultados.columns)
            self.tabla_resultados["show"] = "headings"

            for col in resultados.columns:
                self.tabla_resultados.heading(col, text=col)
                self.tabla_resultados.column(col, width=100)  # Ajusta el ancho de las columnas

            # Inserta los datos en el Treeview
            for index, fila in resultados.iterrows():
                self.tabla_resultados.insert("", tk.END, values=list(fila))
        elif resultados is not None and resultados.empty:
            messagebox.showinfo("Información", "No se encontraron resultados.")
        elif resultados is None:
            # Ya se mostró un mensaje de advertencia o error al cargar o buscar
            pass

    def crear_excel_resultados(self):
        """
        Exporta los resultados a un archivo (funcionalidad no implementada).
        """
        messagebox.showinfo("Información", "Funcionalidad de exportación no implementada.")

    def _ocultar_innecesarios(self):
        """
        Ocultaremos cualquier boton o marco que no sirvan si no hay un excel cargado
        """
        self.boton_cargar_excel_secundario.grid_remove()
        self.boton_busqueda.grid_remove()
        self.boton_crear_excel_resultados.grid_remove()

        self.marco_busqueda.grid_remove()
        self.marco_resultados.grid_remove()
        self.marco_acciones.grid_remove()

    def _mostrar_necesarios(self):
        """
        Mostraremos cualquier boton o marco que no sirvan si no hay un excel cargado
        """
        self.boton_cargar_excel_secundario.grid(row=0, column=1, padx=5, sticky="w")
        self.boton_busqueda.grid(row=0, column=1, padx=5, sticky="e")
        self.boton_crear_excel_resultados.grid(row=0, column=1, padx=5, sticky="e")

        self.marco_busqueda.grid(row=1, column=0, sticky="ew", pady=5)
        self.marco_resultados.grid(row=2, column=0, sticky="nsew", pady=5)
        self.marco_acciones.grid(row=3, column=0, sticky="ew", pady=5)

    def _añadir_busqueda(self):
        pass



if __name__ == "__main__":
    app = Aplicacion()
    app.mainloop()
