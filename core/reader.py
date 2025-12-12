import pandas as pd
import os
from typing import List, Optional

def leer_archivos_excel(archivos: List[str]) -> List[pd.DataFrame]:
    """
    Lee múltiples archivos Excel y los convierte en DataFrames.
    
    Args:
        archivos: Lista de rutas de archivos Excel
        
    Returns:
        Lista de DataFrames con los datos leídos
    """
    dataframes = []
    
    for archivo in archivos:
        try:
            # Intentar leer con diferentes configuraciones
            df = _leer_archivo_excel(archivo)
            if df is not None and not df.empty:
                dataframes.append(df)
                print(f"📄 {os.path.basename(archivo)} -> {len(df)} filas")
            else:
                print(f"⚠️ {os.path.basename(archivo)} está vacío o no se pudo leer")
                
        except Exception as e:
            print(f"❌ Error leyendo {archivo}: {e}")
            continue
    
    return dataframes

def _leer_archivo_excel(archivo_path: str) -> Optional[pd.DataFrame]:
    """Lee un archivo Excel individual con manejo robusto de errores"""
    try:
        # Primero intentar con header en fila 1 (índice 0)
        df = pd.read_excel(archivo_path, header=0, dtype=str)
        
        # Si no tiene columnas válidas, intentar con header en fila 2
        if df.empty or not _tiene_columnas_validas(df):
            df = pd.read_excel(archivo_path, header=1, dtype=str)
        
        # Limpiar DataFrame
        df = _limpiar_dataframe(df)
        
        return df if not df.empty else None
        
    except Exception as e:
        print(f"Error en lectura de {archivo_path}: {e}")
        return None

def _tiene_columnas_validas(df: pd.DataFrame) -> bool:
    """Verifica si el DataFrame tiene columnas relevantes"""
    columnas_relevantes = ['tiempo', 'fecha', 'id', 'cedula', 'empleado']
    columnas_df = [col.lower() for col in df.columns]
    
    return any(col in ' '.join(columnas_df) for col in columnas_relevantes)

def _limpiar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia y normaliza el DataFrame"""
    # Eliminar columnas completamente vacías
    df = df.dropna(axis=1, how='all')
    
    # Eliminar filas completamente vacías
    df = df.dropna(how='all')
    
    # Normalizar nombres de columnas
    df.columns = [_normalizar_nombre_columna(col) for col in df.columns]
    
    return df

def _normalizar_nombre_columna(nombre: str) -> str:
    """Normaliza el nombre de una columna"""
    if pd.isna(nombre):
        return "columna_desconocida"
    
    nombre = str(nombre).strip().lower()
    reemplazos = {
        ' ': '_', '-': '_', '+': '_', '.': '_',
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u'
    }
    
    for viejo, nuevo in reemplazos.items():
        nombre = nombre.replace(viejo, nuevo)
    
    return nombre