# Sistema de Procesamiento de Datos de MercadoLibre

Este sistema permite procesar estados de cuenta de MercadoLibre, desde archivos PDF hasta un Excel concentrado.

## Scripts Disponibles

### 1. `extract.py` (anteriormente extract_pdf_to_excel.py)
Extrae datos de estados de cuenta en PDF y los guarda en un archivo Excel.

```bash
python extract.py
```

- **Entrada**: Archivos PDF en la carpeta `MERCADOPDF`
- **Salida**: Archivo Excel en `MERCADOEXCEL/estado_cuenta.xlsx`

### 2. `cruce1r.py` (anteriormente process_excel.py)
Cruza datos de Excel de reportes con el archivo concentrado, usando IDs de 16 dígitos.

```bash
python cruce1r.py
```

- **Entrada**: Archivos Excel en `REPORTE-ML/*.xls*` y `RESULTADO-FINAL/CONCENTRADO-MERCADOLIBRE.xlsx`
- **Salida**: Actualización del archivo concentrado y reporte de IDs no encontrados

### 3. `cruce2m.py` (anteriormente cross_excel_data.py)
Cruza datos del Excel generado por `extract.py` con el archivo concentrado, usando IDs de 11 dígitos.

```bash
python cruce2m.py
```

- **Entrada**: Archivo Excel en `MERCADOEXCEL` y `RESULTADO-FINAL/CONCENTRADO-MERCADOLIBRE.xlsx`
- **Salida**: Actualización del archivo concentrado y reporte `IDs_no_encontrados.txt`

## Estructura de Carpetas
- `MERCADOPDF`: Almacena los archivos PDF de estados de cuenta
- `MERCADOEXCEL`: Contiene el Excel generado con los datos extraídos de los PDF
- `REPORTE-ML`: Contiene archivos Excel con reportes adicionales
- `RESULTADO-FINAL`: Contiene el archivo concentrado final

## Flujo de Trabajo Recomendado
1. Ejecutar `extract.py` para procesar los PDF
2. Ejecutar `cruce1r.py` para procesar los reportes Excel
3. Ejecutar `cruce2m.py` para realizar el cruce final de datos

Cada script genera informes detallados en la consola sobre su ejecución.