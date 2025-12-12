# requirements_autoinstall.py
import sys
import subprocess
import importlib

def install_package(package):
    """Instala un paquete si no está disponible"""
    try:
        importlib.import_module(package)
        print(f"✅ {package} ya está instalado")
        return True
    except ImportError:
        print(f"📦 Instalando {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"✅ {package} instalado correctamente")
            return True
        except subprocess.CalledProcessError:
            print(f"❌ Error instalando {package}")
            return False

def main():
    """Instala automáticamente todas las dependencias"""
    packages = [
        "pandas>=1.5.0",
        "openpyxl>=3.0.0", 
        "requests>=2.28.0",
        "pyinstaller>=5.0.0"
    ]
    
    print("🔧 INSTALACIÓN AUTOMÁTICA DE DEPENDENCIAS")
    print("=" * 50)
    
    all_installed = True
    for package in packages:
        if not install_package(package.split('>=')[0]):
            all_installed = False
    
    if all_installed:
        print("\n🎉 Todas las dependencias están instaladas")
        print("📦 Puedes construir el ejecutable con: python setup.py")
    else:
        print("\n⚠️  Algunas dependencias no se pudieron instalar automáticamente")
        print("📋 Instala manualmente con: pip install -r requirements.txt")

if __name__ == "__main__":
    main()