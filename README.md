```markdown
# Buscador Avanzado (v0.5.8)

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)

Una aplicación de búsqueda avanzada que permite buscar términos en archivos Excel (diccionarios y descripciones) con soporte para operadores lógicos, comparaciones numéricas y exportación de resultados.

## Características Principales

- **Búsqueda en Diccionarios y Descripciones**: Carga archivos Excel separados para términos de referencia (diccionario) y datos a buscar (descripciones).
- **Operadores Avanzados**:
  - Lógicos: `+` (AND), `|` o `/` (OR)
  - Comparaciones: `>`, `<`, `>=`, `<=`
  - Rangos: `num1-num2`
  - Negación: `#término`
- **Exportación de Resultados**: Guarda reglas de búsqueda y resultados en archivos Excel.
- **Interfaz Gráfica Intuitiva**: Desarrollada con Tkinter.

## Requisitos

- Python 3.7+
- Dependencias:
  - `pandas`
  - `openpyxl` (para archivos .xlsx)
  - `tkinter` (normalmente incluido en Python)

Instalar dependencias:
```bash
pip install pandas openpyxl
```

## Uso

1. **Cargar Archivos**:
   - **Diccionario**: Archivo Excel con términos de referencia.
   - **Descripciones**: Archivo Excel con datos donde buscar.

2. **Realizar Búsquedas**:
   - Escriba términos en el campo de búsqueda.
   - Use operadores para refinar la búsqueda (ver Ayuda en la aplicación).

3. **Exportar Resultados**:
   - Guarde reglas de búsqueda con "Salvar Regla".
   - Exporte todas las reglas guardadas con "Exportar".

### Ejemplos de Búsqueda

- `router + cisco`: Busca descripciones que contengan ambos términos.
- `>1000w`: Busca valores mayores a 1000 con unidad "w".
- `10-20 puertos`: Busca rangos numéricos entre 10 y 20 seguidos de "puertos".
- `switch + #gestionable`: Busca "switch" pero excluye "gestionable".

## Configuración

La aplicación guarda automáticamente:
- Últimas rutas de archivos cargados.
- Índices de columnas para búsqueda en diccionario.

## Capturas de Pantalla

*(Incluir imágenes de la interfaz si es posible)*

## Licencia

MIT License. Ver archivo [LICENSE](LICENSE) para más detalles.

## Notas de Versión (v0.5.8)

- Mejoras en la validación de operadores.
- Soporte para unidades en comparaciones numéricas (ej: `>1000w`).
- Optimización del rendimiento en búsquedas grandes.

## Contribuciones

¡Las contribuciones son bienvenidas! Abra un issue o envíe un pull request.

---
> **Nota**: Para ayuda detallada sobre operadores, use el botón `?` en la aplicación.
``` 
