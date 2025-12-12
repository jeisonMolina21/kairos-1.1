# core/festivos.py
import requests
from datetime import datetime
import json
import os

def obtener_festivos(año):
    """Obtiene los días festivos para Colombia de un año específico"""
    try:
        url = f"https://date.nager.at/api/v3/PublicHolidays/{año}/CO"
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        festivos = response.json()
        
        # Guardar en cache local
        cache_dir = "temp"
        os.makedirs(cache_dir, exist_ok=True)
        cache_file = os.path.join(cache_dir, f"festivos_{año}.json")
        
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(festivos, f, ensure_ascii=False, indent=2)
        
        return festivos
    except Exception as e:
        print(f"⚠️ Error obteniendo festivos para {año}: {e}")
        # Intentar cargar desde cache
        return cargar_festivos_desde_cache(año)

def cargar_festivos_desde_cache(año):
    """Carga festivos desde cache local si existen"""
    cache_file = f"temp/festivos_{año}.json"
    try:
        with open(cache_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []

def es_dia_festivo(fecha):
    """Verifica si una fecha es festiva y retorna (es_festivo, nombre_festivo)"""
    año = fecha.year
    festivos = obtener_festivos(año)
    
    for festivo in festivos:
        try:
            fecha_festivo = datetime.strptime(festivo['date'], '%Y-%m-%d').date()
            if fecha_festivo == fecha:
                return True, festivo['localName']
        except:
            continue
    
    return False, None