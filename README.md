# 💊 Pharmacy Inventory Manager (Spanish bellow)

A Streamlit web application for managing pharmacy inventory — tracking expiry dates, stock levels, product rotation, and duplicates — all from a single Excel upload.

---

## Features

### 🏠 Dashboard
A quick overview of your inventory health at a glance: total unique products, already-expired items, products missing expiry dates, items with zero rotation, and detected duplicates.

### 📅 Expiry Management
- **Already expired** — products past their expiry date that still have stock (sorted by days overdue)
- **Expiring soon** — set a custom window (1–36 months) and get a filtered list of products expiring within that period, including days remaining
- **No expiry date** — products with stock but no registered expiry date

All three views are exportable to Excel individually, or as a single multi-sheet report.

### 📦 Stock Management
Filter by product families (pre-selected defaults: `DIETETICA Y HERBORISTERIA`, `SOLARES`, `COSMÉTICA`, `CORPORAL`) or add your own via a multi-select dropdown.

- **Stock below minimum** — products with `S.Actual == 1` and sell units > 0 (items that sell and need restocking)
- **Zero rotation with stock** — items sitting in inventory without selling
- **No expiry registered** — selling products with missing expiry data

Exportable individually or as a single multi-sheet report.

### 🔍 Duplicate Detection
Scans the inventory for repeated product codes and shows all affected rows. Duplicates are removed from all other analyses (first occurrence is kept) so your numbers are always accurate.

### 📊 Full Inventory Analysis
- Top N products by selled units (adjustable slider)
- Top N products by rotation (adjustable slider)
- Products with stock but no expiry date (full inventory scope)
- Complete multi-sheet export: most rotated, no expiry, zero rotation with stock, full inventory

---

## Duplicate Handling

Duplicates are resolved **at load time**, before any analysis runs:

- The original file (with duplicates) is preserved and shown only in the **🔍 Duplicados** section
- All other pages use a **deduplicated dataset** (first occurrence of each `Código` is kept)
- The dashboard shows how many duplicate rows and codes were removed

---

## Requirements

```
Python >= 3.9
streamlit
pandas
openpyxl
```

Install dependencies:

```bash
pip install streamlit pandas openpyxl
```

---

## Running the App

```bash
streamlit run farmacia_app.py
```

The app will open automatically at `http://localhost:8501`.

---

## Excel File Format

The app expects an Excel file (`.xlsx` or `.xls`) with the following columns. Column names must match exactly (including dots and accents):

| Column | Description |
|---|---|
| `Código` | Unique product code |
| `Descripción` | Product name |
| `Familia` | Product family / category |
| `S.Actual` | Current stock |
| `S.Minimo` | Minimum stock threshold |
| `S.Maximo` | Maximum stock threshold |
| `Rotación` | Rotation / Margin , variable described by the software entreprise |
| `Caducidad` | Expiry date (`DD/MM/YY` or `DD/MM/YYYY`) |
| `Uds.Vendidas` |Sales whitin the current year |

Columns that don't match these names exactly will be ignored. If your file uses different names, rename them before uploading or update `COLUMNAS_INTERES` in the script.

---

## Project Structure

```
farmacia_app.py   # Main Streamlit application
README.md         # This file
requirements.txt  # File with dependencies

```

---

## Notes

- Expiry dates are parsed in multiple formats (`DD/MM/YY`, `DD/MM/YYYY`, `YYYY-MM-DD`, `MM/DD/YYYY`). Products with unrecognisable or missing dates are treated as `Unknown`.
- Numeric columns (`S.Actual`, `S.Minimo`, `S.Maximo`, `Rotación`, `Uds.Vendidas` ) support both `.` and `,` as decimal separators.
- All exported files use `openpyxl` and are compatible with Excel, LibreOffice, and Google Sheets.


# 💊 Gestor de inventario de farmacia

Una aplicación web de Streamlit para gestionar el inventario de una farmacia —controlando fechas de caducidad, niveles de stock, rotación de productos y duplicados— todo ello a partir de una única importación de Excel.

---

## Características

### 🏠 Panel de control
Una visión general rápida del estado de tu inventario de un solo vistazo: total de productos únicos, artículos ya caducados, productos sin fecha de caducidad, artículos sin rotación y duplicados detectados.

### 📅 Gestión de caducidades
- **Ya caducados**: productos que han superado su fecha de caducidad y que aún tienen existencias (ordenados por días de retraso)
- **Caducan pronto**: establece un intervalo personalizado (1-36 meses) y obtén una lista filtrada de los productos que caducan en ese periodo, incluyendo los días restantes
- **Sin fecha de caducidad**: productos con stock pero sin fecha de caducidad registrada

Las tres vistas se pueden exportar a Excel de forma individual o como un único informe de varias hojas.

### 📦 Gestión de stock
Filtra por familias de productos (opciones predeterminadas preseleccionadas: `DIETETICA Y HERBORISTERIA`, `SOLARES`, `COSMÉTICA`, `CORPORAL`) o añade las tuyas propias mediante un menú desplegable de selección múltiple.

- **Stock por debajo del mínimo** — productos con `S.Actual == 1` y unidades vendidas > 0 (artículos que se venden y necesitan reposición)
- **Rotación nula con stock**: artículos que permanecen en el inventario sin venderse
- **Sin fecha de caducidad registrada**: productos en venta a los que les faltan los datos de caducidad

Se puede exportar individualmente o como un único informe de varias hojas.

### 🔍 Detección de duplicados
Analiza el inventario en busca de códigos de producto repetidos y muestra todas las filas afectadas. Los duplicados se eliminan de todos los demás análisis (se conserva la primera aparición), por lo que tus cifras siempre serán precisas.

### 📊 Análisis completo del inventario
- Los N productos más vendidos por unidades (control deslizante ajustable)
- Los N productos con mayor rotación (control deslizante ajustable)
- Productos con stock pero sin fecha de caducidad (alcance completo del inventario)
- Exportación completa a varias hojas: más vendidos, sin caducidad, sin rotación con stock, inventario completo

---

## Gestión de duplicados

Los duplicados se resuelven **en el momento de la carga**, antes de que se ejecute cualquier análisis:

- El archivo original (con duplicados) se conserva y se muestra únicamente en la sección **🔍 Duplicados**
- Todas las demás páginas utilizan un **conjunto de datos deduplicado** (se conserva la primera aparición de cada `Código`)
- El panel muestra cuántas filas y códigos duplicados se han eliminado

---

## Requisitos

```
Python >= 3.9
streamlit
pandas
openpyxl
```

Instalar dependencias:

```bash
pip install streamlit pandas openpyxl
```

---

## Ejecutar la aplicación

```bash
streamlit run farmacia_app.py
```

La aplicación se abrirá automáticamente en `http://localhost:8501`.

---


## Formato de archivo Excel

La aplicación espera un archivo Excel (`.xlsx` o `.xls`) con las siguientes columnas. Los nombres de las columnas deben coincidir exactamente (incluidos los puntos y los acentos):

| Columna | Descripción |
|---|---|
| `Código` | Código único del producto |
| `Descripción` | Nombre del producto |
| `Familia` | Familia o categoría del producto |
| `S.Actual` | Stock actual |
| `S.Minimo` | Umbral mínimo de stock |
| `S.Maximo` | Umbral máximo de stock |
| `Rotación` | Rotación / Margen, variable descrita por la empresa de software |
| `Caducidad` | Fecha de caducidad (`DD/MM/AA` o `DD/MM/AAAA`) |
| `Uds.Vendidas` | Ventas en el año en curso |

Las columnas que no coincidan exactamente con estos nombres serán ignoradas. Si tu archivo utiliza nombres diferentes, cámbialos antes de subirlo o actualiza `COLUMNAS_INTERES` en el script.

---

## Estructura del proyecto

```
farmacia_app.py   # Aplicación principal de Streamlit
README.md         # Este archivo
requirements.txt  #Archivo con dependencias
```

---

## Notas

- Las fechas de caducidad se analizan en varios formatos (`DD/MM/AA`, `DD/MM/AAAA`, `AAAA-MM-DD`, `MM/DD/AAAA`). Los productos con fechas irreconocibles o que faltan se tratan como `Desconocido`.
- Las columnas numéricas (`S.Actual`, `S.Minimo`, `S.Maximo`, `Rotación`, `Uds.Vendidas`) admiten tanto `.` como `,` como separadores decimales.
- Todos los archivos exportados utilizan `openpyxl` y son compatibles con Excel, LibreOffice y Google Sheets.
