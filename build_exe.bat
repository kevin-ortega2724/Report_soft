@echo off
REM Genera el ejecutable de Windows para ReportSoft - Consolidados.
REM Debe correrse en Windows, dentro de esta carpeta, con el .venv ya creado.

call "%~dp0.venv\Scripts\activate.bat"
if errorlevel 1 (
    echo No se encontro .venv\Scripts\activate.bat -- crea el entorno primero.
    exit /b 1
)

pyinstaller "%~dp0consolidados.spec" --clean
if errorlevel 1 (
    echo.
    echo La generacion del ejecutable fallo. Revisa el error de arriba.
    exit /b 1
)

echo.
echo Listo. El ejecutable quedo en dist\ReportSoft-Consolidados\ReportSoft-Consolidados.exe
echo Copia las carpetas data\ y config\ junto al .exe antes de correrlo.
