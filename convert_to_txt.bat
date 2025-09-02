@echo off
REM ========================================
REM Excel a TXT - Conversion Simple
REM ========================================
REM Convierte archivos XLSX a TXT con formato europeo
REM Arrastra archivos .xlsx sobre este script
REM ========================================

echo.
echo ==========================================
echo    CONVERTIR EXCEL A TXT
echo ==========================================

REM Si no hay parametro, mostrar instrucciones
if "%~1"=="" (
    echo.
    echo COMO USAR:
    echo 1. Arrastra un archivo .xlsx sobre este script, O
    echo 2. Ejecuta: convert_to_txt.bat "archivo.xlsx"
    echo.
    echo QUE HACE:
    echo - Toma solo la primera hoja del Excel
    echo - Convierte formulas a valores
    echo - Redondea ultima columna a 3 decimales  
    echo - Cambia puntos por comas (formato europeo)
    echo - Exporta como TXT con tabulaciones
    echo.
    pause
    exit /b 0
)

REM Verificar archivo existe
if not exist "%~1" (
    echo ERROR: Archivo no encontrado
    pause
    exit /b 1
)

REM Verificar que es .xlsx
if /i not "%~x1"==".xlsx" (
    echo ERROR: Solo archivos .xlsx
    echo Archivo: %~nx1
    pause
    exit /b 1
)

echo Procesando: %~nx1
echo.

REM Convertir archivo
python main.py -i "%~1" -c configs/xlsx_to_txt_values.yaml

if %errorlevel% equ 0 (
    echo.
    echo ✓ CONVERSION EXITOSA
    echo.
    echo Archivo creado: %~n1_mod.txt
    echo Ubicacion: %~dp1
) else (
    echo.
    echo ✗ ERROR - Revisar que:
    echo - Excel este cerrado
    echo - Archivo no este danado
)

echo.
pause