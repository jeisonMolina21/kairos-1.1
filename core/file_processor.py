"""
Módulo para procesamiento y gestión de archivos
Funciones auxiliares para limpieza, creación de directorios y procesamiento de archivos
"""
import os
import shutil
import glob
import json
from datetime import datetime
from typing import List, Tuple, Optional
from core.config import paths, departments


# =============================================================================
# MANEJO DE DIRECTORIOS TEMPORALES
# =============================================================================

def limpiar_archivos_temporales():
    """Limpia todos los archivos temporales de ejecuciones anteriores"""
    temp_dirs = [paths.TEMP_DIR, os.path.join(paths.OUTPUT_DIR, "temp")]
    
    for temp_dir in temp_dirs:
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                print(f"🧹 Carpeta temporal '{temp_dir}' eliminada")
            except Exception as e:
                print(f"⚠️ No se pudo eliminar '{temp_dir}': {e}")


def crear_directorios_temporales() -> str:
    """
    Crea los directorios temporales necesarios
    
    Returns:
        Ruta del directorio temporal principal
    """
    temp_dirs = [paths.TEMP_DIR, os.path.join(paths.OUTPUT_DIR, "temp")]
    
    for temp_dir in temp_dirs:
        os.makedirs(temp_dir, exist_ok=True)
    
    return paths.TEMP_DIR


def limpiar_temp_files_en_directorio(directorio: str):
    """
    Limpia solo archivos dentro de un directorio, manteniendo el directorio
    
    Args:
        directorio: Ruta del directorio a limpiar
    """
    if os.path.exists(directorio):
        try:
            for archivo in os.listdir(directorio):
                archivo_path = os.path.join(directorio, archivo)
                if os.path.isfile(archivo_path):
                    os.remove(archivo_path)
            print(f"✅ Archivos temporales eliminados de {directorio}")
        except Exception as e:
            print(f"⚠️ Error limpiando archivos: {e}")


# =============================================================================
# PROCESAMIENTO DE FECHAS Y NOMBRES DE ARCHIVO
# =============================================================================

def obtener_rango_fechas(resumen_path: str) -> Tuple[Optional[datetime], Optional[datetime]]:
    """
    Obtiene el rango de fechas del archivo de resumen
    
    Args:
        resumen_path: Ruta al archivo JSON de resumen
        
    Returns:
        Tupla con (fecha_minima, fecha_maxima) o (None, None) si hay error
    """
    try:
        with open(resumen_path, "r", encoding="utf-8") as f:
            resumen = json.load(f)
        
        if not resumen:
            return None, None
            
        fechas = []
        for registro in resumen:
            fecha_str = registro.get("Fecha")
            if fecha_str and fecha_str != "N/A":
                try:
                    fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
                    fechas.append(fecha)
                except Exception:
                    continue
        
        if not fechas:
            return None, None
            
        return min(fechas), max(fechas)
        
    except Exception as e:
        print(f"⚠️ Error obteniendo rango de fechas: {e}")
        return None, None


def detectar_departamento_archivos(archivos: List[str]) -> List[str]:
    """
    Detecta TODOS los departamentos basados en los nombres de los archivos Excel
    
    Args:
        archivos: Lista de rutas de archivos
        
    Returns:
        Lista de códigos de departamentos detectados
    """
    print("\\n🔍 Detectando departamentos de los archivos...")
    
    departamentos_encontrados = set()
    
    for archivo in archivos:
        nombre_archivo = os.path.basename(archivo).upper()
        print(f"📁 Analizando archivo: {nombre_archivo}")
        
        # Usar configuración centralizada
        for depto_nombre, codigo in departments.MAPEO_ARCHIVOS.items():
            if depto_nombre in nombre_archivo:
                departamentos_encontrados.add(codigo)
                print(f"   ✅ Detectado: {depto_nombre} -> {codigo}")
                break
    
    if departamentos_encontrados:
        departamentos_ordenados = sorted(list(departamentos_encontrados))
        print(f"🏢 Departamentos detectados: {', '.join(departamentos_ordenados)}")
        return departamentos_ordenados
    else:
        print("⚠️ No se pudo detectar departamentos. Usando 'ND'")
        return ["ND"]


def generar_nombre_archivo_excel(
    departamentos: List[str],
    fecha_min: Optional[datetime],
    fecha_max: Optional[datetime]
) -> str:
    """
    Genera el nombre del archivo Excel basado en departamentos y fechas
    
    Args:
        departamentos: Lista de códigos de departamentos
        fecha_min: Fecha mínima
        fecha_max: Fecha máxima
        
    Returns:
        Nombre del archivo Excel
    """
    codigos_str = "_".join(departamentos)
    
    if fecha_min and fecha_max:
        rango_fechas = f"{fecha_min.strftime('%d-%m')}_{fecha_max.strftime('%d-%m')}"
        return f"{codigos_str}_{rango_fechas}.xlsx"
    else:
        return f"{codigos_str}_reporte.xlsx"


def generar_nombre_archivo_json(
    departamentos: List[str],
    fecha_min: Optional[datetime],
    fecha_max: Optional[datetime]
) -> str:
    """
    Genera el nombre del archivo JSON basado en departamentos y fechas
    
    Args:
        departamentos: Lista de códigos de departamentos
        fecha_min: Fecha mínima
        fecha_max: Fecha máxima
        
    Returns:
        Nombre del archivo JSON
    """
    codigos_str = "_".join(departamentos)
    
    if fecha_min and fecha_max:
        rango_fechas = f"{fecha_min.strftime('%d-%m')}_{fecha_max.strftime('%d-%m')}"
        return f"Asistencia_{codigos_str}_{rango_fechas}.json"
    else:
        return f"Asistencia_{codigos_str}.json"


# =============================================================================
# VERIFICACIÓN DE ARCHIVOS
# =============================================================================

def verificar_archivos_generados(archivos_paths: List[Tuple[str, str]]):
    """
    Verifica que los archivos esperados se hayan generado correctamente
    
    Args:
        archivos_paths: Lista de tuplas (nombre_descriptivo, ruta_archivo)
    """
    print("\\n" + "="*60)
    print("📁 ARCHIVOS GENERADOS")
    print("="*60)
    
    for nombre, ruta in archivos_paths:
        if os.path.exists(ruta):
            tamaño = os.path.getsize(ruta)
            print(f"✅ {nombre}: {ruta} ({tamaño:,} bytes)")
        else:
            print(f"❌ {nombre}: NO ENCONTRADO")


# =============================================================================
# HELPERS DE INFORMACIÓN
# =============================================================================

def mostrar_informacion_fechas(
    fecha_min: Optional[datetime],
    fecha_max: Optional[datetime],
    codigos_departamentos: List[str]
):
    """
    Muestra información sobre el rango de fechas y departamentos detectados
    
    Args:
        fecha_min: Fecha mínima
        fecha_max: Fecha máxima
        codigos_departamentos: Lista de códigos de departamentos
    """
    if fecha_min and fecha_max:
        rango_fechas = f"{fecha_min.strftime('%d-%m')}_{fecha_max.strftime('%d-%m')}"
        codigos_str = "_".join(codigos_departamentos)
        print(f"📅 Rango de fechas detectado: {rango_fechas}")
        print(f"🏢 Códigos de departamentos: {codigos_str}")
    else:
        print("⚠️ No se pudieron detectar las fechas para nombrar archivos")
