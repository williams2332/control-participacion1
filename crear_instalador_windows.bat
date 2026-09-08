@echo off
setlocal

title Crear instalador de Control de Participacion

echo ==============================================
echo   CREANDO INSTALADOR DE CONTROL DE PARTICIPACION
echo ==============================================
echo.

if not exist "dist\ControlParticipacion.exe" (
    echo ERROR: No se encontro dist\ControlParticipacion.exe
    echo Primero ejecuta crear_app_windows.bat
    pause
    exit /b 1
)

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
    echo ERROR: No se encontro Inno Setup 6.
    echo Instala Inno Setup desde https://jrsoftware.org/isinfo.php
    pause
    exit /b 1
)

"%ISCC%" "instalador_control_participacion.iss"
if errorlevel 1 (
    echo ERROR: No se pudo crear el instalador.
    pause
    exit /b 1
)

echo.
echo ==============================================
echo INSTALADOR CREADO CORRECTAMENTE
echo ==============================================
echo Se encuentra en:
echo instalador\Instalador_ControlParticipacion.exe
echo.
pause
