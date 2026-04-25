@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0.."

echo.
echo ================================================
echo   Build KINETICA
echo ================================================
echo.

:: Activar venv
if not exist "venv\Scripts\activate.bat" (
    echo ERROR: No se encontro el entorno virtual. Ejecuta este script desde la raiz del proyecto.
    pause & exit /b 1
)
call venv\Scripts\activate.bat

:: Instalar pyinstaller si no está
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Instalando PyInstaller...
    python -m pip install pyinstaller --quiet
)

:: Limpiar builds anteriores
echo Limpiando builds anteriores...
if exist "dist\KINETICA" rmdir /s /q "dist\KINETICA"
if exist "build"          rmdir /s /q "build"

:: Compilar con PyInstaller
echo.
echo [1/2] Compilando con PyInstaller...
python -m PyInstaller installer\pami_bot.spec --noconfirm

if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller fallo. Revisa los mensajes anteriores.
    pause & exit /b 1
)

echo.
echo [1/2] OK - Ejecutables generados en dist\KINETICA\

:: Compilar instalador con Inno Setup
echo.
echo [2/2] Compilando instalador con Inno Setup...

set ISCC_PATH=
for %%P in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
) do (
    if exist %%P set ISCC_PATH=%%P
)

if "!ISCC_PATH!"=="" (
    echo.
    echo NOTA: Inno Setup 6 no encontrado.
    echo   Descargalo desde: https://jrsoftware.org/isdl.php
    echo   Luego ejecuta manualmente:
    echo     "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
    echo.
    echo El ejecutable suelto esta disponible en: dist\KINETICA\KINETICA.exe
) else (
    !ISCC_PATH! installer\setup.iss
    if errorlevel 1 (
        echo ERROR: Inno Setup fallo.
    ) else (
        echo.
        echo [2/2] OK - Instalador generado en: dist\KINETICA_setup.exe
    )
)

echo.
echo ================================================
echo   Build finalizado
echo ================================================
echo.
pause
