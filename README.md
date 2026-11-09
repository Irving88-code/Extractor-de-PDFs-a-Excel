# Extractor de PDFs a Excel

Herramienta de escritorio que lee **facturas y reportes en PDF** y exporta automáticamente los datos a archivos `.xlsx` con formato profesional.

---

## Funcionalidades

- **Extracción inteligente** de texto, tablas y metadatos con `pdfplumber`
- **Detección automática** de campos de factura: número, fecha, total, IVA, proveedor, cliente
- **Exportación a Excel** con múltiples hojas:
  - `Resumen` — metadatos + campos clave detectados
  - `Texto por Página` — contenido completo de cada página
  - `Tabla N` — una hoja por cada tabla encontrada en el PDF
- **Formato visual** aplicado automáticamente: encabezados coloreados, filas alternas, columnas ajustadas
- **Interfaz gráfica** (tkinter) sin dependencias externas de UI
- Procesamiento **por lotes**: selecciona varios PDFs a la vez
- Opción de **abrir carpeta** destino al finalizar

---

## Capturas de pantalla

> La carpeta `demo/` incluye dos PDFs listos para probar:
> - `factura_demo.pdf` — Factura con tabla de conceptos, totales e IVA
> - `reporte_ventas_demo.pdf` — Reporte trimestral con múltiples tablas y KPIs

---

## Instalación

**Requisitos:** Python 3.10 o superior

```bash
pip install -r requirements.txt
```

---

## Uso

```bash
python extractor.py
```

O doble clic en `lanzar.bat` (Windows).

1. Haz clic en **Agregar PDFs** y selecciona uno o varios archivos
2. Elige la **carpeta de destino**
3. Presiona **Extraer y Exportar**
4. Los archivos `.xlsx` se generan automáticamente

---

## Generar PDFs de demostración

```bash
python generar_demo.py
```

Crea `demo/factura_demo.pdf` y `demo/reporte_ventas_demo.pdf` para pruebas.

---

## Tecnologías

| Librería | Uso |
|---|---|
| `pdfplumber` | Lectura de texto y tablas desde PDF |
| `pandas` | Manipulación de datos tabulares |
| `openpyxl` | Escritura y formato de archivos Excel |
| `fpdf2` | Generación de PDFs de demostración |
| `tkinter` | Interfaz gráfica (incluida en Python) |

---

## Estructura del proyecto

```
Extractor de PDFs a Excel/
├── extractor.py        # Aplicación principal
├── generar_demo.py     # Generador de PDFs de prueba
├── lanzar.bat          # Acceso rápido en Windows
├── requirements.txt
├── README.md
└── demo/
    ├── factura_demo.pdf
    └── reporte_ventas_demo.pdf
```

---

## Personalización

La lógica de extracción está separada en tres funciones independientes:

- `extract_pdf(path)` — extrae todo el contenido del PDF
- `build_dataframes(extracted)` — transforma los datos en DataFrames
- `export_to_excel(frames, output_path)` — genera el `.xlsx` con estilo

Esto permite integrar el extractor en pipelines automatizados o servicios web sin modificar la interfaz gráfica.
