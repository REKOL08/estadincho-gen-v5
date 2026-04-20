@echo off
chcp 65001 >nul
title Estadincho-Gen v5.0

if "%~1"=="" (
    echo.
    echo  =====================================================
    echo    ESTADINCHO-GEN v5.0
    echo  =====================================================
    echo.
    echo  Arrastra tu archivo de datos sobre este .bat
    echo  o escribe la ruta a continuacion:
    echo.
    set /p RUTA_ARCHIVO="  Ruta del archivo: "
    echo.
) else (
    set RUTA_ARCHIVO=%~1
)

python "%~dp0generar_dashboard.py" "%RUTA_ARCHIVO%"

if errorlevel 1 (
    echo.
    echo  ERROR: Algo salio mal. Verifica que Python este instalado.
    pause
)
