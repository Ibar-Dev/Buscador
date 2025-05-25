# Buscador v0.7.5

![Python](https://img.shields.io/badge/Python-3.6%2B-blue)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green)
![Pandas](https://img.shields.io/badge/Data-Pandas-yellow)

## 📋 Descripción

Buscador es una aplicación de escritorio para búsquedas avanzadas en archivos Excel. Permite realizar consultas complejas utilizando operadores lógicos, comparaciones numéricas, rangos y negaciones. La aplicación está diseñada para trabajar con dos tipos de archivos:

- **Diccionario**: Archivo Excel que contiene términos de referencia.
- **Descripciones**: Archivo Excel con datos que se desean consultar.

La aplicación puede buscar directamente en las descripciones o utilizar el diccionario como intermediario para encontrar coincidencias más relevantes.

## ✨ Características

- Interfaz gráfica intuitiva basada en Tkinter
- Carga y visualización de archivos Excel (.xlsx, .xls)
- Búsqueda avanzada con múltiples operadores:
  - Operadores lógicos: AND (`+`), OR (`|` o `/`)
  - Comparaciones numéricas: `>`, `<`, `>=`, `<=`, `=`
  - Rangos numéricos: `num1 - num2`
  - Negación (exclusión): `#palabra` o `#"frase completa"`
- Normalización de texto para búsquedas insensibles a mayúsculas/minúsculas y acentos
- Exportación de resultados a Excel o CSV
- Guardado de reglas/búsquedas para uso posterior
- Configuración persistente entre sesiones

## 🔧 Requisitos

- Python 3.6 o superior
- Dependencias:
  - pandas
  - openpyxl (para archivos .xlsx)
  - tkinter (incluido en la mayoría de instalaciones de Python)

## 📦 Instalación

1. Asegúrese de tener Python 3.6+ instalado
2. Instale las dependencias:

```bash
pip install pandas openpyxl
```

3. Clone o descargue este repositorio
4. Ejecute el script:

```bash
python Buscador_v0_7_5.py
```

## 🚀 Uso

### Carga de archivos

1. Inicie la aplicación
2. Haga clic en "Cargar Diccionario" para seleccionar el archivo Excel de diccionario
3. Haga clic en "Cargar Descripciones" para seleccionar el archivo Excel de descripciones

### Sintaxis de búsqueda

- **Texto simple**: Busca la palabra o frase. Ejemplo: `router cisco`
- **Operadores lógicos**:
  - `término1 + término2`: Busca filas con AMBOS términos (AND). Ejemplo: `tarjeta + 16 puertos`
  - `término1 | término2`: Busca filas con AL MENOS UNO de los términos (OR). Ejemplo: `modulo | SFP`
- **Comparaciones numéricas** (unidad opcional):
  - `> num[UNIDAD]`: Mayor que. Ejemplo: `> 1000` o `> 1000W`
  - `< num[UNIDAD]`: Menor que. Ejemplo: `< 50` o `< 50V`
  - `>= num[UNIDAD]` o `≥ num[UNIDAD]`: Mayor o igual que. Ejemplo: `>= 48A`
  - `<= num[UNIDAD]` o `≤ num[UNIDAD]`: Menor o igual que. Ejemplo: `<= 10.5W`
  - `= num[UNIDAD]`: Igual a. Ejemplo: `= 24V`
- **Rangos numéricos** (unidad opcional, ambos extremos incluidos):
  - `num1 - num2[UNIDAD]`: Entre num1 y num2. Ejemplo: `10 - 20` o `50 - 100V`
- **Negación** (excluir):
  - `#palabra`: Excluye filas que contengan `palabra`. Ejemplo: `switch + #gestionable`
  - `#"frase completa"`: Excluye filas con la frase. Ejemplo: `fuente + #"bajo rendimiento"`

### Modos de búsqueda

1. **Vía Diccionario**: La búsqueda se aplica primero al Diccionario. Si hay coincidencias (FCDs), se extraen términos clave que luego se buscan en las Descripciones.
2. **Directa**: Si no hay coincidencias vía Diccionario, se puede buscar directamente en las Descripciones.
3. **Búsqueda vacía**: Muestra todas las descripciones.

### Exportación de resultados

1. Realice una búsqueda
2. Haga clic en "Exportar" para guardar los resultados en formato Excel (.xlsx) o CSV

## 🔍 Flujo de trabajo

1. **Carga de archivos**: Diccionario y Descripciones
2. **Formulación de consulta**: Utilice la sintaxis de búsqueda para crear su consulta
3. **Ejecución de búsqueda**: Haga clic en "Buscar" o presione Enter
4. **Visualización de resultados**: Los resultados se muestran en la tabla inferior
5. **Exportación** (opcional): Guarde los resultados en un archivo

## 🧩 Estructura del código

- `MotorBusqueda`: Clase principal que maneja la lógica de búsqueda
- `ExtractorMagnitud`: Clase para normalizar y extraer unidades de medida
- `ManejadorExcel`: Clase estática para cargar archivos Excel
- `InterfazGrafica`: Clase que implementa la interfaz de usuario con Tkinter
- `OrigenResultados`: Enumeración para rastrear el origen de los resultados

## 📝 Registro y depuración

La aplicación genera un archivo de registro `buscador_app_refactorizado.log` que puede ser útil para depuración.

## 🔄 Configuración persistente

La aplicación guarda la configuración en un archivo JSON (`config_buscador_v0_7_4_mapeo_refactor.json`) que incluye:
- Rutas de los últimos archivos cargados
- Índices de columnas para búsqueda en el diccionario

## 👨‍💻 Desarrollo

### Extensión de funcionalidades

Para añadir nuevas funcionalidades:
1. Extienda las clases existentes o cree nuevas según sea necesario
2. Actualice la interfaz gráfica para exponer las nuevas funcionalidades
3. Mantenga la coherencia con el flujo de trabajo existente

### Mejoras potenciales

- Soporte para más formatos de archivo
- Búsqueda en múltiples archivos
- Guardado y carga de consultas complejas
- Visualización de datos con gráficos
- Filtros adicionales para resultados

## 📄 Licencia

[Incluir información de licencia aquí]

## 👥 Contribuciones

[Incluir información sobre cómo contribuir al proyecto]

## 📞 Contacto

[Incluir información de contacto]
