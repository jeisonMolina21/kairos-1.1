@echo off
chcp 65001 >nul
title Instalador Automático - Sistema KAIROS V4
color 0A

echo.
echo ================================================
echo    🚀 INSTALADOR AUTOMÁTICO KAIROS V4
echo ================================================
echo.

:: Verificar si estamos en Windows
ver | find "Windows" > nul
if errorlevel 1 (
    echo ❌ Este instalador solo funciona en Windows
    pause
    exit /b 1
)

:: Obtener ruta actual
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo 🔍 Verificando requisitos del sistema...
echo.

:: Verificar arquitectura
echo 📊 Arquitectura del sistema:
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo ✅ Sistema de 64 bits detectado
) else (
    echo ✅ Sistema de 32 bits detectado
)

:: Verificar espacio en disco
for /f "tokens=3" %%a in ('dir /-c ^| find "bytes libres"') do set free=%%a
echo 💾 Espacio libre en disco: %free%

echo.
echo 📦 Iniciando instalación automática...
echo.

:: Crear entorno virtual
echo 🔧 Creando entorno virtual Python...
python -m venv venv --copies
if errorlevel 1 (
    echo ❌ Error creando entorno virtual
    echo 🔧 Intentando con virtualenv...
    pip install virtualenv
    virtualenv venv
)

:: Activar entorno virtual
echo 🔧 Activando entorno virtual...
call venv\Scripts\activate.bat

:: Instalar dependencias
echo 📚 Instalando dependencias de Python...
venv\Scripts\python.exe -m pip install --upgrade pip

echo 📥 Instalando pandas...
venv\Scripts\python.exe -m pip install pandas==1.5.3

echo 📥 Instalando openpyxl...
venv\Scripts\python.exe -m pip install openpyxl==3.0.10

echo 📥 Instalando requests...
venv\Scripts\python.exe -m pip install requests==2.28.2

echo 📥 Instalando pyinstaller...
venv\Scripts\python.exe -m pip install pyinstaller==5.9.0

echo 📥 Instalando tkinter...
venv\Scripts\python.exe -m pip install tkinter

echo 📥 Instalando pillow...
venv\Scripts\python.exe -m pip install pillow==9.4.0

echo.
echo 🔨 Compilando ejecutable...
echo.

:: Crear ejecutable
venv\Scripts\pyinstaller.exe --onefile --console ^
  --name "KairosV4" ^
  --icon=icon.ico ^
  --add-data "core;core" ^
  --add-data "reports;reports" ^
  --hidden-import=pandas ^
  --hidden-import=openpyxl ^
  --hidden-import=requests ^
  --hidden-import=tkinter ^
  --hidden-import=json ^
  --hidden-import=os ^
  --hidden-import=sys ^
  --hidden-import=datetime ^
  --hidden-import=traceback ^
  --hidden-import=shutil ^
  --hidden-import=glob ^
  --hidden-import=tempfile ^
  --hidden-import=io ^
  --hidden-import=collections ^
  --hidden-import=math ^
  main.py

echo.
if exist "dist\KairosV4.exe" (
    echo ✅ EJECUTABLE CREADO EXITOSAMENTE!
    echo.
    echo 📁 El ejecutable se encuentra en: dist\KairosV4.exe
    echo.
    echo 🎯 Para usar el sistema:
    echo    1. Copie la carpeta 'dist' a cualquier computador
    echo    2. Ejecute 'KairosV4.exe'
    echo    3. No necesita Python instalado
    echo.
) else (
    echo ❌ Error creando el ejecutable
    echo 🔧 Intentando método alternativo...
    
    :: Método alternativo
    venv\Scripts\pyinstaller.exe --onefile --console ^
      --name "KairosV4" ^
      main.py
)

echo.
echo 📋 Creando accesos directos...
echo.

:: Crear script de ejecución rápida
echo @echo off > "Ejecutar Kairos.bat"
echo chcp 65001 >> "Ejecutar Kairos.bat"
echo cd /d "%~dp0" >> "Ejecutar Kairos.bat"
echo echo Iniciando Sistema Kairos V4... >> "Ejecutar Kairos.bat"
echo dist\KairosV4.exe >> "Ejecutar Kairos.bat"
echo pause >> "Ejecutar Kairos.bat"

:: Crear archivo de configuración rápida
echo @echo off > "Configuracion Rapida.bat"
echo chcp 65001 >> "Configuracion Rapida.bat"
echo echo ================================================ >> "Configuracion Rapida.bat"
echo echo    📋 CONFIGURACIÓN RÁPIDA KAIROS V4 >> "Configuracion Rapida.bat"
echo echo ================================================ >> "Configuracion Rapida.bat"
echo echo. >> "Configuracion Rapida.bat"
echo echo 🎯 INSTRUCCIONES: >> "Configuracion Rapida.bat"
echo echo. >> "Configuracion Rapida.bat"
echo echo 1. Ejecute 'Ejecutar Kairos.bat' >> "Configuracion Rapida.bat"
echo echo 2. Seleccione los archivos Excel de asistencia >> "Configuracion Rapida.bat"
echo echo 3. El sistema procesará automáticamente >> "Configuracion Rapida.bat"
echo echo 4. Los reportes se guardan en carpeta 'outputs' >> "Configuracion Rapida.bat"
echo echo. >> "Configuracion Rapida.bat"
echo echo ⚠️  REQUISITOS: >> "Configuracion Rapida.bat"
echo echo    - Windows 7 o superior >> "Configuracion Rapida.bat"
echo echo    - Conexión a Internet >> "Configuracion Rapida.bat"
echo echo    - 500MB espacio libre >> "Configuracion Rapida.bat"
echo echo. >> "Configuracion Rapida.bat"
echo pause >> "Configuracion Rapida.bat"

echo.
echo 🎉 INSTALACIÓN COMPLETADA!
echo.
echo 📂 Estructura final:
echo    📄 KairosV4.exe          (Ejecutable principal)
echo    📄 Ejecutar Kairos.bat   (Acceso directo)
echo    📂 core/                 (Módulos del sistema)
echo    📂 reports/              (Generadores de reportes)
echo    📂 outputs/              (Reportes generados)
echo.
echo 🚀 Ejecute 'Ejecutar Kairos.bat' para comenzar
echo.

pause