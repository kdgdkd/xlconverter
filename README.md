# Excel Transformer

Herramienta para transformar archivos Excel/HTML basada en configuraciones YAML. Convierte archivos con reglas personalizables sin necesidad de programar.

## Instalación

### 1. Instalar Python

**Windows:**
1. Descargar Python desde [python.org](https://www.python.org/downloads/)
2. **IMPORTANTE**: Durante la instalación, marcar "Add Python to PATH"
3. Verificar instalación abriendo CMD y ejecutando:
   ```cmd
   python --version
   ```

**macOS:**
```bash
# Usando Homebrew (recomendado)
brew install python

# O descargar desde python.org
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install python3 python3-pip
```

### 2. Instalar Dependencias

1. **Descargar/clonar este proyecto** en tu ordenador
2. **Abrir terminal/CMD** en la carpeta del proyecto
3. **Instalar librerías** requeridas:

```bash
# Windows
pip install -r requirements.txt

# macOS/Linux  
pip3 install -r requirements.txt
```

### 3. Verificar Instalación

Ejecutar comando de prueba:
```bash
python main.py --help
```

Si aparece la ayuda del programa, la instalación fue exitosa.

## Uso Básico

### Sintaxis Principal

```bash
python main.py -i ARCHIVO_ENTRADA -c CONFIGURACION [opciones]
```

### Ejemplos de Uso

```bash
# Transformar archivo con configuración específica
python main.py -i datos.xlsx -c configs/balance_sheet.yaml

# Especificar archivo de salida personalizado
python main.py -i reporte.xls -c configs/clean_html_export.yaml -o resultado.xlsx

# Modo verbose para ver detalles del procesamiento
python main.py -i archivo.xlsx -c configs/import_reporting.yaml -v
```

### Nombrado Automático de Archivos

Por defecto, los archivos procesados se guardan:
- **En la misma carpeta** que el archivo original
- **Con sufijo '_mod'** añadido al nombre: `datos.xlsx` → `datos_mod.xlsx`
- **Nombres únicos automáticos**: Si `datos_mod.xlsx` ya existe, se crea `datos_mod2.xlsx`

Ejemplos:
- `balance.xlsx` → `balance_mod.xlsx`
- `reporte.xls` → `reporte_mod.txt` (si se exporta como TXT)
- Si ya existe: `balance_mod2.xlsx`, `balance_mod3.xlsx`, etc.

## Comandos Adicionales

```bash
# Ver configuraciones disponibles
python main.py list-configs

# Ver contenido de una configuración específica
python main.py show-config balance_sheet
```

## Configuraciones

Las configuraciones se encuentran en la carpeta `configs/`. Cada archivo YAML define un conjunto de transformaciones específicas para diferentes casos de uso.

### Estructura de Configuración

```yaml
name: "Nombre descriptivo"
description: "Descripción de qué hace esta configuración"

# Opciones de carga (opcional)
load_options:
  sheet_name: 0        # Solo primera hoja (Excel)
  data_only: true      # Solo valores, no fórmulas

# Transformaciones a aplicar
transformations:
  - type: "delete_columns"
    columns: [0, "Column Name"]
  
  - type: "round_numbers"
    columns: ["Amount"]
    decimals: 2

# Configuración de exportación
export:
  format: "xlsx"       # xlsx, txt, csv
  delimiter: "\t"      # Para TXT/CSV
  headers: true        # Incluir encabezados
```

### Tipos de Transformaciones Soportadas

#### Eliminar Datos
```yaml
# Eliminar columnas por índice, nombre o patrón
- type: "delete_columns"
  columns: [0, "Unnamed", "Total"]

# Eliminar filas por contenido, vacías o rango
- type: "delete_rows"
  conditions:
    - contains: ["TOTAL", "---"]
    - empty: true
    - range: [0, 4]  # Eliminar filas 0-4
```

#### Formatear Números
```yaml
# Redondear números
- type: "round_numbers"
  columns: ["Amount", "Price"]
  decimals: 2

# Redondear solo última columna
- type: "round_last_column"
  decimals: 3

# Formato contabilidad (Excel)
- type: "format_accounting"
  columns: ["B", "C", "D"]  # Por letra de columna
  decimals: 2

# Cambiar separador decimal
- type: "replace_decimal_separator"
  columns: "last"  # Solo última columna
  from_separator: "."
  to_separator: ","
```

#### Manipular Columnas
```yaml
# Renombrar columnas
- type: "rename_columns"
  mapping:
    "Old Name": "New Name"
    "Importe": "Amount"

# Reordenar columnas
- type: "reorder_columns"
  order: ["Date", "Description", "Amount"]

# Seleccionar solo ciertas columnas
- type: "select_columns"
  columns: ["Account", "Amount", "Date"]
```

#### Formatear Fechas
```yaml
- type: "date_format"
  columns: ["Date"]
  from_format: "%d/%m/%Y"
  to_format: "%Y%m%d"
```

## Formatos de Exportación

### Excel (.xlsx)
```yaml
export:
  format: "xlsx"
  apply_excel_formatting: true  # Aplicar formatos nativos de Excel
```

### Texto (.txt) / CSV
```yaml
export:
  format: "txt"
  delimiter: "\t"      # Tabulación
  headers: true        # Incluir encabezados
```

## Casos de Uso Comunes

### 1. Limpiar Reportes HTML exportados como XLS
- Eliminar filas de encabezado del sistema
- Quitar columnas vacías
- Formatear números correctamente
- Exportar como Excel limpio

### 2. Preparar Datos para Importar a Sistema Legacy
- Seleccionar solo columnas necesarias
- Formatear números sin separadores de miles
- Convertir fechas al formato requerido
- Exportar como TXT delimitado

### 3. Procesar Archivos con Fórmulas
- Convertir fórmulas a valores puros
- Eliminar hojas adicionales
- Aplicar formato decimal europeo (comas)
- Exportar para uso externo

## Solución de Problemas

### Error: "python no se reconoce como comando"
- **Solución**: Python no está en el PATH del sistema
- **Windows**: Reinstalar Python marcando "Add Python to PATH"
- **Verificar**: `where python` debe mostrar la ruta de instalación

### Error: "No module named 'pandas'"
- **Solución**: Dependencias no instaladas correctamente
- **Ejecutar**: `pip install -r requirements.txt`

### Error: "Permission denied" al guardar
- **Causa**: Archivo de salida está abierto en Excel/otro programa
- **Solución**: Cerrar el archivo antes de procesar, o se creará automáticamente con sufijo _mod2, _mod3, etc.

### Error: "Archivo no encontrado"
- **Verificar**: Ruta del archivo de entrada es correcta
- **Verificar**: Archivo de configuración existe en carpeta `configs/`

## Archivos de Entrada Soportados

- **HTML con extensión .xls**: Exportaciones de sistemas legacy
- **Excel (.xlsx)**: Archivos Excel nativos
- **CSV**: Archivos de texto separados por comas

## Requisitos del Sistema

- **Python**: 3.8 o superior
- **Sistema Operativo**: Windows, macOS, Linux
- **Espacio**: ~50MB para dependencias
- **RAM**: ~100MB durante procesamiento