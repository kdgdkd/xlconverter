@echo off
REM ========================================
REM Excel Transformer - Procesamiento Rapido
REM ========================================
REM Script simple para procesar un archivo especifico
REM Arrastra un archivo .xlsx/.xls sobre este .bat
REM ========================================

echo.
echo ==========================================
echo    EXCEL TRANSFORMER - PROCESAMIENTO RAPIDO
echo ==========================================

REM Verificar que se paso un archivo como parametro
if "%~1"=="" (
    echo.
    echo USO: Arrastrar un archivo .xlsx o .xls sobre este script
    echo.
    echo Alternativamente:
    echo quick_process.bat "ruta\al\archivo.xlsx"
    echo.
    pause
    exit /b 1
)

REM Verificar que el archivo existe
if not exist "%~1" (
    echo ERROR: Archivo no encontrado: %~1
    echo.
    pause
    exit /b 1
)

REM Obtener extension del archivo
set "file_ext=%~x1"
if /i not "%file_ext%"==".xlsx" if /i not "%file_ext%"==".xls" (
    echo ERROR: Solo se admiten archivos .xlsx y .xls
    echo Archivo recibido: %~nx1
    echo.
    pause
    exit /b 1
)

echo Archivo a procesar: %~nx1
echo Ubicacion: %~dp1
echo.

REM Seleccion rapida de configuracion segun tipo de archivo
if /i "%file_ext%"==".xls" (
    set "default_config=configs\clean_html_export.yaml"
    set "default_desc=Limpieza HTML/XLS (recomendado para .xls)"
) else (
    set "default_config=configs\xlsx_to_txt_values.yaml"
    set "default_desc=XLSX a TXT europeo (recomendado para .xlsx)"
)

echo CONFIGURACION SUGERIDA:
echo %default_desc%
echo.
echo 1. Usar configuracion sugerida
echo 2. Seleccionar otra configuracion
echo.
set /p "choice=Opcion (1-2): "

if "%choice%"=="1" (
    set "config=%default_config%"
    goto :process
)

REM Menu de configuraciones
echo.
echo CONFIGURACIONES DISPONIBLES:
echo 1. Limpiar reportes HTML/XLS
echo 2. XLSX a TXT formato europeo  
echo 3. Balance con formato contabilidad
echo.
set /p "config_choice=Seleccionar (1-3): "

if "%config_choice%"=="1" set "config=configs\clean_html_export.yaml"
if "%config_choice%"=="2" set "config=configs\xlsx_to_txt_values.yaml"
if "%config_choice%"=="3" set "config=configs\test_config.yaml"

:process
echo.
echo ==========================================
echo PROCESANDO...
echo ==========================================
echo Archivo: %~nx1
echo Configuracion: %config%
echo.

REM Procesar el archivo
python main.py -i "%~1" -c "%config%" -v

if %errorlevel% equ 0 (
    echo.
    echo ==========================================
    echo ✓ PROCESAMIENTO EXITOSO
    echo ==========================================
    echo.
    echo El archivo procesado se guardo como:
    echo %~n1_mod%~x1 (o con extension segun configuracion)
    echo.
    echo En la misma carpeta que el archivo original:
    echo %~dp1
) else (
    echo.
    echo ==========================================
    echo ✗ ERROR EN PROCESAMIENTO
    echo ==========================================
    echo.
    echo Revisar:
    echo - Archivo no este abierto en Excel
    echo - Permisos de escritura en carpeta
    echo - Formato del archivo sea valido
)

echo.
pause