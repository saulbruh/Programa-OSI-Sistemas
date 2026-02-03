@echo off
REM ====================================================
REM  OSI Arecibo — Inventario, Préstamos y Mantenimientos
REM  Autor: Saúl Medina — Versión 2.0
REM  Script para activar entorno virtual, ejecutar la app
REM  y (opcional) COMPILAR ejecutable con PyInstaller
REM ====================================================

SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

REM === CONFIG: carpeta del proyecto ===
cd /d "%~dp0"

REM === Crear venv si no existe ===
if not exist ".venv\" (
    echo [SETUP] Creando entorno virtual...
    py -3 -m venv .venv
)

REM === Activar venv ===
call .\.venv\Scripts\activate

REM === Instalar/actualizar dependencias ===
python -m pip install --upgrade pip
python -m pip install ttkbootstrap pandas openpyxl
python -m pip install pyinstaller

REM === MODO COMPILACION (opcional) ===
REM Usa:   "Mantenimiento y Reparacion OSI.bat build"  para compilar el .exe
if /I "%~1"=="build" goto BUILD
if /I "%~1"=="/build" goto BUILD
if /I "%~1"=="install" goto INSTALL
if /I "%~1"=="/install" goto INSTALL

REM === MODO EJECUCION (por defecto) ===
echo [RUN] Iniciando OSI Arecibo en Python...
python MANT-REP-TEST-FINAL.py

echo.
pause
GOTO END

:BUILD
REM ----------------------------------------------------
REM  Compila ejecutable Windows institucional
REM  (Program Files + ProgramData + Public Desktop)
REM ----------------------------------------------------
set ICON_FLAG=
if exist "icon.ico" set ICON_FLAG=--icon icon.ico

REM Limpieza previa opcional
if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist OSI_Arecibo.spec del /f /q OSI_Arecibo.spec

pyinstaller --onefile --noconsole --name "OSI_Arecibo" %ICON_FLAG% MANT-REP-TEST-FINAL.py

echo.
echo [INFO] Ejecuta el script con el parámetro INSTALL para instalar el sistema:
echo        MantenimientoReparacionOSI.bat install

if not exist "dist\OSI_Arecibo.exe" (
    echo [ERROR] No se encontró dist\OSI_Arecibo.exe. Revisa la salida de PyInstaller.
    echo.
    pause
    goto END
)

:INSTALL
REM ----------------------------------------------------
REM  Instalación institucional
REM ----------------------------------------------------

set APP_NAME=OSI_Arecibo
set EXE_NAME=OSI_Arecibo.exe

set INSTALL_DIR=C:\Program Files\%APP_NAME%
set DATA_DIR=C:\ProgramData\%APP_NAME%
set PUBLIC_DESKTOP=C:\Users\Public\Desktop

echo [INSTALL] Verificando permisos de administrador...
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Este script debe ejecutarse como ADMINISTRADOR.
    pause
    goto END
)

echo [INSTALL] Creando carpetas...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
if not exist "%DATA_DIR%" mkdir "%DATA_DIR%"

echo [INSTALL] Copiando ejecutable...
if not exist "dist\%EXE_NAME%" (
    echo [ERROR] No se encontró dist\%EXE_NAME%. Ejecuta primero el modo BUILD.
    pause
    goto END
)
copy /y "dist\%EXE_NAME%" "%INSTALL_DIR%\%EXE_NAME%" >nul

echo [INSTALL] Creando acceso directo en Public Desktop...
powershell -NoProfile -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut('%PUBLIC_DESKTOP%\\%APP_NAME%.lnk');" ^
  "$s.TargetPath='%INSTALL_DIR%\\%EXE_NAME%';" ^
  "$s.WorkingDirectory='%INSTALL_DIR%';" ^
  "$s.IconLocation='%INSTALL_DIR%\\%EXE_NAME%,0';" ^
  "$s.Save()"

echo.
echo [OK] OSI Arecibo instalado correctamente.
echo      Programa: %INSTALL_DIR%
echo      Datos:     %DATA_DIR%
echo      Acceso:    Public Desktop
pause
goto END

:END
ENDLOCAL