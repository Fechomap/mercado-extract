# Sistema de Procesamiento de Datos de MercadoLibre

Este sistema permite procesar estados de cuenta de MercadoLibre, desde archivos PDF hasta un Excel concentrado.

## Requisitos

- **Python**: Versión 3.7 o superior (probado hasta Python 3.13.2)
- **Dependencias**: pandas, openpyxl, pdfplumber

## Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/Fechomap/mercado-extract.git
cd mercado-extract
```

### 2. Seleccionar la versión correcta de Python

**Si tienes múltiples versiones de Python instaladas:**

**En Windows:**
```bash
# Verificar versiones disponibles
py --list

# Usar una versión específica para crear el entorno virtual
py -3.11 -m venv venv
```

**En macOS/Linux:**
```bash
# Verificar versiones disponibles
ls -l /usr/bin/python*

# Usar una versión específica
python3.11 -m venv venv
```

### 3. Activar el entorno virtual

**En Windows:**
```bash
venv\Scripts\activate
```

**En macOS/Linux:**
```bash
source venv/bin/activate
```

### 4. Instalar dependencias (forma simplificada)
```bash
pip install -r requirements.txt
```

## Estructura de Carpetas (crear antes de ejecutar)
```bash
mkdir MERCADOPDF MERCADOEXCEL REPORTE-ML RESULTADO-FINAL
```

- `MERCADOPDF`: Almacena los archivos PDF de estados de cuenta
- `MERCADOEXCEL`: Contiene el Excel generado con los datos extraídos de los PDF
- `REPORTE-ML`: Contiene archivos Excel con reportes adicionales
- `RESULTADO-FINAL`: Contiene el archivo concentrado final

## Scripts Disponibles

### 1. `extract.py`
Extrae datos de estados de cuenta en PDF y los guarda en un archivo Excel.

```bash
python extract.py
```

- **Entrada**: Archivos PDF en la carpeta `MERCADOPDF`
- **Salida**: Archivo Excel en `MERCADOEXCEL/estado_cuenta.xlsx`

### 2. `cruce1r.py`
Cruza datos de Excel de reportes con el archivo concentrado, usando IDs de 16 dígitos.

```bash
python cruce1r.py
```

- **Entrada**: Archivos Excel en `REPORTE-ML/*.xls*` y `RESULTADO-FINAL/CONCENTRADO-MERCADOLIBRE.xlsx`
- **Salida**: Actualización del archivo concentrado y reporte de IDs no encontrados

### 3. `cruce2m.py`
Cruza datos del Excel generado por `extract.py` con el archivo concentrado, usando IDs de 11 dígitos.

```bash
python cruce2m.py
```

- **Entrada**: Archivo Excel en `MERCADOEXCEL` y `RESULTADO-FINAL/CONCENTRADO-MERCADOLIBRE.xlsx`
- **Salida**: Actualización del archivo concentrado y reporte `IDs_no_encontrados.txt`

## Flujo de Trabajo Recomendado
1. Colocar los PDFs a procesar en la carpeta `MERCADOPDF`
2. Colocar el Excel concentrado vacío en `RESULTADO-FINAL/CONCENTRADO-MERCADOLIBRE.xlsx`
3. Ejecutar `extract.py` para procesar los PDF
4. Colocar los archivos Excel de reportes en la carpeta `REPORTE-ML`
5. Ejecutar `cruce1r.py` para procesar los reportes Excel
6. Ejecutar `cruce2m.py` para realizar el cruce final de datos

## Solución de Problemas

### Errores con la versión de Python
```bash
# Verificar versión de Python en uso
python --version

# Si estás usando Python 3.12+ y ocurren errores
# Instala Python 3.11 desde python.org y crea un nuevo entorno virtual
```

### Errores con dependencias
```bash
# Actualizar pip primero
pip install --upgrade pip

# Instalar dependencias una por una si hay problemas
pip install pandas==1.5.3
pip install openpyxl==3.1.2
pip install pdfplumber==0.9.0
```

Consulta el archivo `INSTALACION.md` para instrucciones más detalladas y soluciones a problemas comunes.