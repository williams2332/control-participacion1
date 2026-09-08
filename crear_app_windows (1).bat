@echo off
setlocal

title Crear Control de Participacion

echo ==============================================
echo   CREANDO APLICACION CONTROL DE PARTICIPACION
echo ==============================================
echo.

if not exist "Logo.jpg" (
    echo ERROR: No se encontro Logo.jpg en esta carpeta.
    echo Coloca el logo junto a este archivo .bat y al archivo .py.
    pause
    exit /b 1
)

python -m pip install --upgrade pyinstaller openpyxl pillow
if errorlevel 1 (
    echo ERROR: No se pudieron instalar las dependencias.
    pause
    exit /b 1
)

python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "ControlParticipacion" ^
    --add-data "Logo.jpg;." ^
    control_participacion_bd.py

if errorlevel 1 (
    echo ERROR: No se pudo crear el ejecutable.
    pause
    exit /b 1
)

echo.
echo ==============================================
echo APLICACION CREADA CORRECTAMENTE
echo El ejecutable esta en la carpeta dist.
echo ==============================================
pause
