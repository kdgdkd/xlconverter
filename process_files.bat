@echo off
REM ========================================
REM Excel Transformer - Procesamiento Automatico
REM ========================================
REM Este script procesa multiples archivos automaticamente
REM usando configuraciones predefinidas
REM ========================================

echo.
echo ==========================================
echo    EXCEL TRANSFORMER - PROCESAMIENTO
echo ==========================================
echo.

REM Verificar que Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no esta instalado o no esta en PATH
    echo.
    echo Soluciones:
    echo 1. Instalar Python desde https://www.python.org/
    echo 2. Durante instalacion marcar "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

REM Verificar que estamos en el directorio correcto
if not exist "main.py" (
    echo ERROR: No se encuentra main.py
    echo Ejecutar este script desde la carpeta xlsxconverter
    echo.
    pause
    exit /b 1
)

REM Crear carpeta de entrada si no existe
if not exist "input" mkdir input

REM Verificar si hay archivos para procesar
set "files_found=0"
for %%f in (input\*.xlsx input\*.xls) do (
    set "files_found=1"
    goto :found_files
)

:found_files
if "%files_found%"=="0" (
    echo.
    echo INFO: No se encontraron archivos para procesar
    echo.
    echo INSTRUCCIONES:
    echo 1. Copiar archivos .xlsx o .xls a la carpeta "input"
    echo 2. Ejecutar este script nuevamente
    echo.
    pause
    exit /b 0
)

echo Archivos encontrados para procesar:
echo.
for %%f in (input\*.xlsx input\*.xls) do (
    echo   - %%~nxf
)
echo.

REM Menu de configuraciones
:menu
echo ==========================================
echo SELECCIONAR CONFIGURACION:
echo ==========================================
echo.
echo 1. Limpiar reportes HTML/XLS del sistema legacy
echo 2. Convertir XLSX a TXT con formato europeo 
echo 3. Procesar balance con formato contabilidad
echo 4. Ver configuraciones disponibles
echo 5. Salir
echo.
set /p "choice=Seleccionar opcion (1-5): "

if "%choice%"=="1" set "config=configs\clean_html_export.yaml" & set "desc=Limpieza HTML/XLS" & goto :process
if "%choice%"=="2" set "config=configs\xlsx_to_txt_values.yaml" & set "desc=XLSX a TXT europeo" & goto :process  
if "%choice%"=="3" set "config=configs\test_config.yaml" & set "desc=Balance contabilidad" & goto :process
if "%choice%"=="4" goto :show_configs
if "%choice%"=="5" goto :exit

echo Opcion invalida. Intentar de nuevo.
echo.
goto :menu

:show_configs
echo.
echo ==========================================
echo CONFIGURACIONES DISPONIBLES:
echo ==========================================
python main.py list-configs
echo.
pause
goto :menu

:process
echo.
echo ==========================================
echo PROCESANDO ARCHIVOS...
echo ==========================================
echo Configuracion: %desc%
echo Archivo config: %config%
echo.

set "processed=0"
set "errors=0"

for %%f in (input\*.xlsx input\*.xls) do (
    echo Procesando: %%~nxf
    python main.py -i "%%f" -c "%config%" -v
    
    if !errorlevel! equ 0 (
        set /a processed+=1
        echo   ✓ Procesado exitosamente
    ) else (
        set /a errors+=1
        echo   ✗ Error al procesar
    )
    echo.
)

echo ==========================================
echo RESUMEN DE PROCESAMIENTO:
echo ==========================================
echo Archivos procesados: %processed%
echo Errores: %errors%
echo.

if %errors% gtr 0 (
    echo NOTA: Revisar errores antes de continuar
    echo Los archivos con errores no fueron procesados
)

if %processed% gtr 0 (
    echo Los archivos procesados se guardaron con sufijo "_mod"
    echo en la misma carpeta que los originales
)

echo.
echo ¿Procesar mas archivos?
set /p "continue=Presionar S para volver al menu, cualquier tecla para salir: "
if /i "%continue%"=="s" goto :menu

:exit
echo.
echo ==========================================
echo PROCESAMIENTO COMPLETADO
echo ==========================================
echo.
echo Archivos de salida ubicados en carpeta "input"
echo con sufijo "_mod" añadido al nombre original
echo.
echo Gracias por usar Excel Transformer!
echo.
pause
exit /b 0