@echo off
chcp 65001 >nul
title Creando Ejecutable - Sistema de Asistencias

echo 🚀 INICIANDO CREACION DE EJECUTABLE
echo ========================================

echo 🔍 Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python no está instalado o no está en el PATH
    echo 💡 Por favor instala Python desde https://python.org
    pause
    exit /b 1
)

echo ✅ Python detectado correctamente
echo.

echo 📦 Instalando dependencias y creando ejecutable...
python installer.py

if errorlevel 1 (
    echo.
    echo ❌ Error durante la instalacion
    pause
    exit /b 1
)

echo.
echo 🎉 PROCESO COMPLETADO!
echo 📁 El ejecutable esta en: dist\SistemaProcesamientoAsistencia.exe
echo.
pause