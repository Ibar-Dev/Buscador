# 🔍 Buscador Avanzado - Herramienta de Búsqueda y Análisis de Datos

Una aplicación GUI para realizar búsquedas avanzadas en archivos Excel, diseñada para facilitar el análisis de datos técnicos, inventarios o descripciones complejas.

![Interfaz gráfica](https://via.placeholder.com/800x500.png?text=Captura+de+la+Interfaz) *(Ejemplo de interfaz)*

---

## 🚀 Características Principales

- **Carga de archivos Excel**:
  - **Diccionario**: Define términos clave, especificaciones o condiciones.
  - **Descripciones**: Contiene datos a analizar (ej.: productos, componentes).

- **Sintaxis de búsqueda avanzada**:
  - Operadores lógicos: `+` (AND), `|` o `/` (OR), `#` (NOT).
  - Comparaciones numéricas: `>100`, `<=50`, `>=24.5`.
  - Rangos: `10-20`, `5.5-15.3`.
  - Búsqueda directa en descripciones.

- **Exportación flexible**:
  - Formatos soportados: Excel (`.xlsx`, `.xls`), CSV (UTF-8).
  - Nombres de archivo automáticos basados en la búsqueda.

- **Interfaz intuitiva**:
  - Vista previa de datos.
  - Colores alternos en filas.
  - Ordenación por columnas.

- **Compatibilidad**:
  - Soporte para Excel moderno (`.xlsx`) y legacy (`.xls`).
  - Multiplataforma (Windows, Linux, macOS).

---

## 📦 Requisitos Previos

- Python 3.6 o superior
- Dependencias:
  ```bash
  pip install pandas openpyxl xlwt
  ```

---

## 🛠 Instalación y Uso

1. **Clonar repositorio** (o descargar el script):
   ```bash
   git clone https://github.com/tu-usuario/buscador-avanzado.git
   cd buscador-avanzado
   ```

2. **Ejecutar la aplicación**:
   ```bash
   python Buscador_v0_4_8.py
   ```

3. **Pasos básicos**:
   - Cargar **Diccionario** y **Descripciones** desde archivos Excel.
   - Ingresar términos de búsqueda con sintaxis avanzada (ej: `switch + #gestionable`).
   - Exportar resultados con un clic.

---

## 📖 Ejemplos de Uso

| Búsqueda                 | Descripción                                  |
|--------------------------|---------------------------------------------|
| `router + cisco`         | Filas con "router" Y "cisco" en diccionario |
| `>1000`                  | Valores numéricos mayores a 1000            |
| `#gestionable`           | Excluye términos con "gestionable"          |
| `10-20 | 5-8`         | Rangos numéricos o valores específicos      |
| `"tarjeta red"`          | Búsqueda exacta en descripciones            |

---

## 🛠️ Estructura del Proyecto

```
buscador-avanzado/
├── Buscador_v0_4_8.py    # Código principal
├── config_buscador.json  # Configuración guardada
├── buscador_app.log      # Registro de actividad
└── README.md             # Este archivo
```

---

## 📜 Licencia

Distribuido bajo la licencia MIT. Ver `LICENSE` para más detalles.

---

## 🤝 Contribuir

¡Contribuciones son bienvenidas! Abre un *issue* para reportar errores o un *pull request* para mejoras.

---

**Hecho con ❤️ por Ibar-Dev**  
*¿Preguntas?* ✉️ **IbarVivas@gmail.com**
``` 

Este README incluye:
- Descripción clara del propósito.
- Características destacadas con emojis.
- Requisitos e instalación.
- Ejemplos prácticos.
- Estructura de archivos.
- Licencia y sección de contribución.
- Diseño responsive y elementos visuales (placeholders para imágenes).

Personaliza los enlaces, nombres y datos de contacto según tu proyecto.
