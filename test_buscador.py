import unittest
import pandas as pd
from unittest.mock import patch, MagicMock
from tkinter import Tk
from tkinter import filedialog
import os
import importlib.util

spec = importlib.util.spec_from_file_location("Buscador",
                                             r"C:\Users\Ibarv\.vscode\workspace\Deltacomgroup\buscador\Buscador_v0.1.1.py")
Buscador = importlib.util.module_from_spec(spec)
spec.loader.exec_module(Buscador)

ManejadorExcel = Buscador.ManejadorExcel
MotorBusqueda = Buscador.MotorBusqueda
InterfazGrafica = Buscador.InterfazGrafica


class TestManejadorExcel(unittest.TestCase):
    def setUp(self):
        # Create a dummy Excel file for testing
        self.test_data = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        self.test_file = 'test_excel.xlsx'
        self.test_data.to_excel(self.test_file, index=False)

    def tearDown(self):
        # Remove the dummy Excel file after testing
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_cargar_excel_success(self):
        df = ManejadorExcel.cargar_excel(self.test_file)
        self.assertIsNotNone(df)
        pd.testing.assert_frame_equal(df, self.test_data)

    def test_cargar_excel_file_not_found(self):
        df = ManejadorExcel.cargar_excel('nonexistent_file.xlsx')
        self.assertIsNone(df)

    def test_comparar_dataframes_equal(self):
        df1 = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        df2 = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        result = ManejadorExcel.comparar_dataframes(df1, df2)
        self.assertTrue(result)

    def test_comparar_dataframes_not_equal(self):
        df1 = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        df2 = pd.DataFrame({'col1': [3, 4], 'col2': ['c', 'd']})
        result = ManejadorExcel.comparar_dataframes(df1, df2)
        self.assertFalse(result)

    def test_comparar_dataframes_different_columns(self):
        df1 = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        df2 = pd.DataFrame({'col1': [1, 2], 'col3': ['a', 'b']})
        result = ManejadorExcel.comparar_dataframes(df1, df2)
        self.assertFalse(result)


class TestMotorBusqueda(unittest.TestCase):
    def setUp(self):
        self.motor = MotorBusqueda()
        self.test_data = pd.DataFrame(
            {'col1': [1, 2, 3], 'col2': ['apple', 'banana', 'apple'], 'col3': ['Apple Pie', 'Banana Bread', 'Orange Juice']})
        self.test_file = 'test_excel.xlsx'
        self.test_data.to_excel(self.test_file, index=False)
        self.motor.cargar_excel(self.test_file)

    def tearDown(self):
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    def test_cargar_excel_success(self):
        self.assertIsNotNone(self.motor.datos)
        pd.testing.assert_frame_equal(self.motor.datos, self.test_data)
        self.assertEqual(self.motor.archivo_actual, self.test_file)

    def test_cargar_excel_file_not_found(self):
        motor = MotorBusqueda()
        result = motor.cargar_excel('nonexistent_file.xlsx')
        self.assertFalse(result)
        self.assertIsNone(motor.datos)
        self.assertIsNone(motor.archivo_actual)

    def test_buscar_success(self):
        resultados = self.motor.buscar('apple')
        self.assertIsNotNone(resultados)
        self.assertEqual(len(resultados), 2)
        self.assertTrue(all(resultados['col2'].str.contains('apple', case=False)))

    def test_buscar_no_results(self):
        resultados = self.motor.buscar('grape')
        self.assertIsNone(resultados)

    def test_buscar_empty_term(self):
        resultados = self.motor.buscar('')
        self.assertIsNotNone(resultados)
        pd.testing.assert_frame_equal(resultados, self.test_data)

    def test_buscar_and_operator(self):
        resultados = self.motor.buscar('apple + pie')
        self.assertIsNotNone(resultados)
        self.assertEqual(len(resultados), 1)
        self.assertTrue(all(resultados['col2'].str.contains('apple', case=False)))
        self.assertTrue(all(resultados['col3'].str.contains('pie', case=False)))

    def test_buscar_or_operator(self):
        resultados = self.motor.buscar('apple - banana')
        self.assertIsNotNone(resultados)
        self.assertEqual(len(resultados), 3)
        self.assertTrue(any(resultados['col2'].str.contains('apple', case=False)))
        self.assertTrue(any(resultados['col2'].str.contains('banana', case=False)))

    def test_buscar_no_data_loaded(self):
        motor = MotorBusqueda()
        resultados = motor.buscar('apple')
        self.assertIsNone(resultados)


class TestInterfazGrafica(unittest.TestCase):
    def setUp(self):
        self.root = Tk()
        self.app = InterfazGrafica()
        self.app.update_idletasks()

        # Create a dummy Excel file for testing
        self.test_data = pd.DataFrame({'col1': [1, 2], 'col2': ['a', 'b']})
        self.test_file = 'test_excel.xlsx'
        self.test_data.to_excel(self.test_file, index=False)

    def tearDown(self):
        self.root.destroy()
        if os.path.exists(self.test_file):
            os.remove(self.test_file)

    @patch('tkinter.filedialog.askopenfilename')
    def test_cargar_principal_success(self, mock_askopenfilename):
        mock_askopenfilename.return_value = self.test_file
        self.app._cargar_principal()
        self.assertIsNotNone(self.app.motor.datos)  # Verifica que los datos se cargaron
        self.assertEqual(self.app.motor.archivo_actual, self.test_file)  # Verifica el archivo cargado
        self.assertEqual(str(self.app.btn_comparar["state"]), "normal")  # Verifica que el botón se habilitó  # Verifica que el botón se habilitó
        print("Tipo:", type(self.app.btn_comparar["state"]))
        print("Valor:", self.app.btn_comparar["state"])
    
    @patch('tkinter.filedialog.askopenfilename')
    def test_cargar_principal_cancel(self, mock_askopenfilename):
        mock_askopenfilename.return_value = ''
        self.app._cargar_principal()
        self.assertIsNone(self.app.motor.datos)
        self.assertIsNone(self.app.motor.archivo_actual)

    @patch('tkinter.filedialog.askopenfilename')
    def test_cargar_principal_file_not_found(self, mock_askopenfilename):
        mock_askopenfilename.return_value = 'nonexistent_file.xlsx'
        self.app._cargar_principal()
        self.assertIsNone(self.app.motor.datos)
        self.assertIsNone(self.app.motor.archivo_actual)


    @patch('tkinter.filedialog.askopenfilename')
    @patch('tkinter.messagebox.showinfo')
    def test_comparar_archivos_not_equal(self, mock_showinfo, mock_askopenfilename):
        mock_askopenfilename.side_effect = [self.test_file, 'test_excel_2.xlsx']
        pd.DataFrame({'col1': [3, 4], 'col2': ['c', 'd']}).to_excel('test_excel_2.xlsx', index=False)
        self.app._cargar_principal()
        self.app._comparar_archivos()
        mock_showinfo.assert_called_with("Comparación", "Los archivos son diferentes")
        if os.path.exists('test_excel_2.xlsx'):
            os.remove('test_excel_2.xlsx')

    @patch('tkinter.filedialog.askopenfilename')
    @patch('tkinter.messagebox.showinfo')
    def test_comparar_archivos_different_columns(self, mock_showinfo, mock_askopenfilename):
        mock_askopenfilename.side_effect = [self.test_file, 'test_excel_2.xlsx']
        pd.DataFrame({'col1': [3, 4], 'col3': ['c', 'd']}).to_excel('test_excel_2.xlsx', index=False)
        self.app._cargar_principal()
        self.app._comparar_archivos()
        mock_showinfo.assert_called_with("Comparación", "Los archivos tienen columnas diferentes")
        if os.path.exists('test_excel_2.xlsx'):
            os.remove('test_excel_2.xlsx')

    @patch('tkinter.filedialog.asksaveasfilename')
    def test_exportar_resultados_success(self, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = 'test_export.xlsx'
        self.app.motor.cargar_excel(self.test_file)
        self.app.motor.buscar('a')
        self.app._exportar_resultados()
        self.assertTrue(os.path.exists('test_export.xlsx'))
        if os.path.exists('test_export.xlsx'):
            os.remove('test_export.xlsx')

    @patch('tkinter.filedialog.asksaveasfilename')
    def test_exportar_resultados_cancel(self, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = ''
        self.app.motor.cargar_excel(self.test_file)
        self.app.motor.buscar('a')
        self.app._exportar_resultados()
        self.assertFalse(os.path.exists('test_export.xlsx'))

    @patch('tkinter.filedialog.asksaveasfilename')
    def test_exportar_resultados_no_results(self, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = 'test_export.xlsx'
        self.app.motor.cargar_excel(self.test_file)
        self.app.motor.buscar('z')
        self.app._exportar_resultados()
        self.assertFalse(os.path.exists('test_export.xlsx'))

    @patch('tkinter.filedialog.asksaveasfilename')
    def test_exportar_resultados_csv(self, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = 'test_export.csv'
        self.app.motor.cargar_excel(self.test_file)
        self.app.motor.buscar('a')
        self.app._exportar_resultados()
        self.assertTrue(os.path.exists('test_export.csv'))
        if os.path.exists('test_export.csv'):
            os.remove('test_export.csv')

    @patch('tkinter.filedialog.asksaveasfilename')
    def test_exportar_resultados_xls(self, mock_asksaveasfilename):
        mock_asksaveasfilename.return_value = 'test_export.xls'
        self.app.motor.cargar_excel(self.test_file)
        self.app.motor.buscar('a')
        self.app._exportar_resultados()
        self.assertTrue(os.path.exists('test_export.xls'))
        if os.path.exists('test_export.xls'):
            os.remove('test_export.xls')

    @patch('tkinter.messagebox.showwarning')
    def test_ejecutar_busqueda_no_data_loaded(self, mock_showwarning):
        self.app.motor.datos = None
        self.app._ejecutar_busqueda()
        mock_showwarning.assert_called_with("Advertencia", "Primero cargue un archivo")

    @patch('tkinter.messagebox.showinfo')
    def test_ejecutar_busqueda_no_results(self, mock_showinfo):
        self.app.motor.cargar_excel(self.test_file)
        self.app.entrada_busqueda.insert(0, 'z')
        self.app._ejecutar_busqueda()
        mock_showinfo.assert_called_with("Información", "No se encontraron resultados.")

    @patch('tkinter.messagebox.showinfo')
    def test_ejecutar_busqueda_success(self, mock_showinfo):
        self.app.motor.cargar_excel(self.test_file)
        self.app.entrada_busqueda.insert(0, 'a')
        self.app._ejecutar_busqueda()
        self.assertFalse(mock_showinfo.called)


if __name__ == '__main__':
    unittest.main()
