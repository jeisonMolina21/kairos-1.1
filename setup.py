# setup.py
from PyInstaller.__main__ import run
import os
import shutil

def build_executable():
    """Construye el ejecutable autónomo"""
    
    # Opciones de PyInstaller
    opts = [
        'main.py',
        '--name=Kairos_Asistencias',
        '--onefile',
        '--windowed',
        '--icon=kairos.ico',
        '--add-data=core;core',
        '--add-data=reports;reports',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=requests',
        '--hidden-import=tkinter',
        '--hidden-import=json',
        '--hidden-import=datetime',
        '--clean',
        '--noconfirm'
    ]
    
    print("🚀 Construyendo ejecutable...")
    run(opts)
    
    # Crear carpeta de distribución si no existe
    if not os.path.exists('dist'):
        os.makedirs('dist')
    
    print("✅ Ejecutable creado en: dist/Kairos_Asistencias.exe")

if __name__ == '__main__':
    build_executable()