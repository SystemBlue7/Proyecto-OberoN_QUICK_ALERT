@echo off
chcp 65001 >nul
title OBERON SYSTEMS - COMPILADOR EXE
color 0A

echo ===========================================================
echo     O B E R O N   S Y S T E M S  -  B U I L D E R  V 1 . 1
echo ===========================================================
echo.
echo [INFO] Iniciando compilación de quick_alert.py...
echo.

REM ============================================================
REM 1. Verificar PyInstaller
REM ============================================================
python -c "import PyInstaller" 2>NUL
if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller no está instalado.
    echo Ejecuta: pip install pyinstaller
    pause
    exit /b
)

echo [OK] PyInstaller detectado.
echo.

REM ============================================================
REM 2. Limpieza previa
REM ============================================================
echo [INFO] Eliminando carpetas previas "build" y "dist"...
rmdir /s /q build 2>NUL
rmdir /s /q dist 2>NUL
echo [OK] Limpieza completada.
echo.

REM ============================================================
REM 3. COMPILACIÓN PRINCIPAL
REM ============================================================
echo [INFO] Compilando quick_alert.py a EXE con ícono personalizado...
echo.

python -m PyInstaller --noconfirm --onefile --windowed ^
 --add-data "O-PEQUEÑA.ico;." ^
 --add-data "Logo.png;." ^
 --add-data "ICONOGRAFIA;ICONOGRAFIA" ^
 --add-data "CABEZOTES;CABEZOTES" ^
 --add-data "EVIDENCIAS;EVIDENCIAS" ^
 --add-data "QUICK ALERT;QUICK ALERT" ^
 quick_alert.py

if %errorlevel% neq 0 (
    echo.
    color 0C
    echo [ERROR FATAL] La compilación falló.
    echo Revisa que quick_alert.py no tenga errores.
    pause
    exit /b
)

REM ============================================================
REM 4. ÉXITO
REM ============================================================
color 0A
echo.
echo ===========================================================
echo       [ÉXITO] Compilación completada correctamente.
echo ===========================================================
echo.
echo [INFO] Tu ejecutable ya está listo en:
echo        dist\quick_alert.exe
echo.
echo Presiona una tecla para abrir la carpeta dist...
pause >nul
start dist
exit

