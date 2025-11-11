@echo off

@echo "Script de compilación FactuNabo"
@echo "Asegúrate de ejecutar esto dentro del entorno virtual con los requisitos instalados."
@echo off

REM Limpia artefactos anteriores
if exist build rd /s /q build
if exist dist rd /s /q dist

echo Compilando ejecutable con FactuNabo.spec...
pyinstaller FactuNabo.spec --noconfirm --clean

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Falló la compilación. Revisa los mensajes anteriores.
    exit /b %errorlevel%
)

echo.
echo [OK] Ejecutable generado en dist\FactuNabo\
echo Copia FactuNabo.exe y la carpeta completa si necesitas una versión portable.
