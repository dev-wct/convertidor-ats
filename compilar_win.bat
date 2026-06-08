@echo off
echo ===================================================
echo   Compilador de Convertidor ATS - World Class
echo ===================================================
echo.

:: Verificar si python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no está instalado o no está en el PATH del sistema.
    echo Por favor, instala Python desde python.org y marca la casilla "Add Python to PATH".
    pause
    exit /b
)

echo [1/3] Instalando dependencias necesarias...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/3] Compilando aplicación con PyInstaller (creando archivo único)...
pyinstaller --noconsole --onefile --add-data "worldclass-logo.png;." --hidden-import "python_calamine" --name "ConvertidorATS_WorldClass" app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Ocurrió un error al compilar con PyInstaller.
    pause
    exit /b
)

echo.
echo [3/3] ¡Compilación exitosa!
echo El ejecutable único se encuentra en la carpeta: dist\ConvertidorATS_WorldClass.exe
echo Puedes mover ese archivo .exe a cualquier carpeta o computadora Windows y funcionará solo.
echo.
pause
