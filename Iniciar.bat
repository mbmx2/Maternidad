@echo off
chcp 65001 >nul
title Sistema de Expedientes - Maternidad
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo No se encontro Python en este equipo.
    echo Instala desde python.org marcando "tcl/tk" (opcional en Python 3.13+).
    echo.
    pause
    exit /b 1
)

echo Abriendo el sistema...
python app.py
if errorlevel 1 (
    echo.
    echo El sistema termino con errores. Revisa _ERRORES.log en esta carpeta.
    pause
)
