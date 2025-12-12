# installer.py
import os
import sys
import subprocess
import importlib.util
from pathlib import Path

def check_and_install_package(package, install_name=None):
    """Verifica si un paquete está instalado y lo instala si no lo está"""
    if install_name is None:
        install_name = package
    
    try:
        spec = importlib.util.find_spec(package)
        if spec is None:
            print(f"📦 Instalando {install_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", install_name])
            print(f"✅ {install_name} instalado correctamente")
        else:
            print(f"✅ {install_name} ya está instalado")
    except Exception as e:
        print(f"❌ Error instalando {install_name}: {e}")
        return False
    return True

def main():
    print("🚀 INSTALADOR AUTOMÁTICO - Sistema de Procesamiento de Asistencia")
    print("=" * 60)
    
    # Verificar que pip está disponible
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "--version"])
    except:
        print("❌ pip no está disponible. Instalando pip...")
        try:
            import ensurepip
            ensurepip.bootstrap()
        except:
            print("❌ No se pudo instalar pip. Por favor instala Python manualmente.")
            input("Presiona Enter para salir...")
            return
    
    # Lista de paquetes necesarios
    packages = [
        ("pandas", "pandas"),
        ("openpyxl", "openpyxl"), 
        ("requests", "requests"),
        ("PyInstaller", "pyinstaller")
    ]
    
    print("\n🔍 Verificando e instalando dependencias...")
    
    all_installed = True
    for package, install_name in packages:
        if not check_and_install_package(package, install_name):
            all_installed = False
    
    if not all_installed:
        print("❌ Algunos paquetes no se pudieron instalar.")
        input("Presiona Enter para salir...")
        return
    
    print("\n✅ Todas las dependencias están instaladas correctamente!")
    
    # Crear el ejecutable
    print("\n🔨 Creando ejecutable...")
    try:
        # Crear archivo .spec para PyInstaller
        spec_content = '''
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('core', 'core'),
        ('reports', 'reports'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'requests',
        'tkinter',
        'json',
        'os',
        'sys',
        'datetime',
        'traceback',
        'shutil',
        'glob',
        'tempfile'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SistemaProcesamientoAsistencia',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SistemaProcesamientoAsistencia',
)
'''
        with open('build.spec', 'w', encoding='utf-8') as f:
            f.write(spec_content)
        
        # Ejecutar PyInstaller
        subprocess.check_call([sys.executable, '-m', 'PyInstaller', 'build.spec', '--onefile', '--windowed'])
        
        print("✅ Ejecutable creado correctamente!")
        print("📁 El ejecutable se encuentra en la carpeta 'dist'")
        
    except Exception as e:
        print(f"❌ Error creando el ejecutable: {e}")
        print("💡 Solución alternativa: Ejecuta manualmente:")
        print('   python -m PyInstaller main.py --onefile --hidden-import=pandas --hidden-import=openpyxl --hidden-import=requests --hidden-import=tkinter --add-data="core;core" --add-data="reports;reports"')
    
    input("\nPresiona Enter para salir...")

if __name__ == "__main__":
    main()