# Guía de Instalación Detallada

## Requisitos del Sistema
- **Python**: Versión 3.7 o superior (probado hasta Python 3.13.2)
- **Sistema Operativo**: Windows 10/11, macOS o Linux
- **RAM**: 4GB mínimo (8GB recomendado)
- **Espacio en disco**: 500MB para la aplicación y dependencias

## Cambiar de Versión de Python

### Si ya tienes Python 3.11 instalado pero estás usando otra versión

**En Windows:**

```bash
# Verificar todas las versiones instaladas
py --list

# Eliminar el entorno virtual actual si existe
rmdir /s /q venv

# Crear un nuevo entorno virtual con Python 3.11 específicamente
py -3.11 -m venv venv

# Activar el nuevo entorno
venv\Scripts\activate

# Verificar que estás usando Python 3.11
python --version
```

**En macOS/Linux:**

```bash
# Verificar versiones instaladas
ls -la /usr/bin/python*
# o
which python3.11

# Eliminar el entorno virtual actual si existe
rm -rf venv

# Crear un nuevo entorno con la versión específica
python3.11 -m venv venv

# Activar el nuevo entorno
source venv/bin/activate

# Verificar versión
python --version
```

### Cambiar la versión predeterminada de Python

**En Windows:**
1. Busca "Editar las variables de entorno del sistema" en el menú de inicio
2. Haz clic en "Variables de entorno"
3. En "Variables del sistema", edita la variable "Path"
4. Mueve la ruta a Python 3.11 al principio de la lista
5. Reinicia cualquier terminal abierta

**En macOS:**
```bash
# Crear alias en .bash_profile o .zshrc
echo 'alias python=python3.11' >> ~/.zshrc
source ~/.zshrc
```

**En Linux:**
```bash
# Actualizar alternativas
sudo update-alternatives --install /usr/bin/python python /usr/bin/python3.11 1
sudo update-alternatives --config python
```

## Solución de Problemas Comunes

### Error al instalar dependencias

Si encuentras errores al instalar `pdfplumber` u otras dependencias:

**En Windows:**
```bash
# Instala las herramientas de compilación de C++
pip install --upgrade pip
pip install wheel
```

**En todos los sistemas:**
```bash
# Instalar dependencias una por una
pip install pandas==1.5.3
pip install openpyxl==3.1.2
pip install pdfplumber==0.9.0
```

### Error "No module named 'venv'"

```bash
# Instala el módulo venv
python -m pip install virtualenv
# Luego usa virtualenv en lugar de venv
python -m virtualenv venv
```

### Error al procesar PDFs

Si encuentras errores al procesar PDFs específicos, verifica:
1. Que el PDF no esté dañado
2. Que el PDF no esté protegido con contraseña
3. Que el formato del PDF sea compatible (texto extraíble, no escaneado)

### Nota sobre la compatibilidad de versiones de Python

Este sistema ha sido probado y funciona correctamente con:
- Python 3.7 - 3.11 (ampliamente probado)
- Python 3.13.2 (confirmado funcional)

Si encuentras problemas con alguna versión específica, puedes:
1. Probar con otra versión de Python
2. Revisar la compatibilidad de las dependencias con tu versión específica
3. Actualizar las dependencias a versiones más recientes que soporten tu versión de Python

## Instalación Paso a Paso con Anaconda (alternativa)

Si prefieres usar Anaconda:

```bash
# Crear ambiente con conda
conda create -n mercado-extract python=3.9
conda activate mercado-extract

# Instalar dependencias
pip install pandas==1.5.3 openpyxl==3.1.2 pdfplumber==0.9.0
```

## Notas para Administradores de TI

- Los scripts no requieren permisos de administrador para ejecutarse
- Si los usuarios están tras un proxy corporativo, configure las variables de entorno HTTP_PROXY y HTTPS_PROXY
- Para despliegues en múltiples equipos, considere crear un entorno virtual compartido en una unidad de red