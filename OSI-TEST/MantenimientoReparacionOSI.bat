@echo off
REM ====================================================
REM  OSI Arecibo — Inventario, Préstamos y Mantenimientos
REM  Autor: Saúl Medina — Versión 2.0
REM  Script para activar entorno virtual, ejecutar la app
REM  y (opcional) COMPILAR ejecutable con PyInstaller
REM ====================================================

SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

REM === CONFIG: carpeta del proyecto ===
cd /d "C:\Users\tecnico01\Desktop\OSI-TEST"

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

REM === MODO EJECUCION (por defecto) ===
echo [RUN] Iniciando OSI Arecibo en Python...
python MANT-REP-TEST-FINAL.py

echo.
pause
GOTO END

:BUILD
REM ----------------------------------------------------
REM  Compila ejecutable Windows (dist\OSI_Arecibo.exe)
REM  y crea carpeta PORTABLE con ./data y recursos
REM ----------------------------------------------------
set ICON_FLAG=
if exist "icon.ico" set ICON_FLAG=--icon icon.ico

REM Limpieza previa opcional
if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist OSI_Arecibo.spec del /f /q OSI_Arecibo.spec

pyinstaller --onefile --noconsole --name "OSI_Arecibo" %ICON_FLAG% MANT-REP-TEST-FINAL.py

if not exist "dist\OSI_Arecibo.exe" (
    echo [ERROR] No se encontró dist\OSI_Arecibo.exe. Revisa la salida de PyInstaller.
    echo.
    pause
    goto END
)

REM === Armar paquete portable ===
set PKG=OSI_Arecibo_Portable
if exist "%PKG%" rmdir /s /q "%PKG%"
mkdir "%PKG%\data"

copy /y "dist\OSI_Arecibo.exe" "%PKG%\OSI_Arecibo.exe" >nul
if exist icon.ico copy /y icon.ico "%PKG%\icon.ico" >nul

REM Copiar datos si están junto al proyecto o en el Desktop del usuario
for %%F in ("Registro Laptops.xlsx" "Registro_Mantenimiento_Reparacion_Laptop.xlsx" "Registro_Prestamos_Laptop.xlsx" "Registro_Decomisados.xlsx") do (
    if exist ".\data\%%~F" copy /y ".\data\%%~F" "%PKG%\data\%%~F" >nul
    if not exist "%PKG%\data\%%~F" if exist "%USERPROFILE%\Desktop\%%~F" copy /y "%USERPROFILE%\Desktop\%%~F" "%PKG%\data\%%~F" >nul
)

REM Archivo README rápido
(
  echo OSI Arecibo — Paquete Portable
  echo.
  echo Estructura:
  echo   OSI_Arecibo_Portable\OSI_Arecibo.exe
  echo   OSI_Arecibo_Portable\data\*.xlsx
  echo.
  echo Puedes mover esta carpeta a cualquier lugar. El .exe usa rutas relativas a .\data
  echo para leer y escribir los Excel.
) > "%PKG%\LEEME.txt"

echo.
echo [BUILD] Paquete portable creado: "%CD%\%PKG%\"
choice /c SN /n /m "¿Abrir carpeta portable ahora? (S/N) > "
if errorlevel 2 goto END
start "" explorer.exe "%CD%\%PKG%\"

:END
ENDLOCAL