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
    """Lee un archivo Excel individual con búsqueda dinámica de encabezados"""
    try:
        df_final = None
        # Buscar el encabezado correcto en las primeras 10 filas
        for i in range(10):
            df_temp = pd.read_excel(archivo_path, header=i, dtype=str)
            if not df_temp.empty and _tiene_columnas_validas(df_temp):
                df_final = df_temp
                print(f"🎯 Encabezados detectados automáticamente en la fila {i+1}")
                break
                
        if df_final is None:
            # Si no encontró las columnas, usar la fila 1 por defecto para que cleaner arroje el error detallado
            df_final = pd.read_excel(archivo_path, header=0, dtype=str)
            
        df_final = _limpiar_dataframe(df_final)
        return df_final if not df_final.empty else None
        
    except Exception as e:
        print(f"❌ Error en lectura de {archivo_path}: {e}")
        return None

def _tiene_columnas_validas(df: pd.DataFrame) -> bool:
    """Verifica si el DataFrame tiene columnas relevantes biométricas"""
    # Reemplazar espacios por guiones bajos para que coincida con nuestras alternativas
    columnas_df = [col.lower().strip().replace(" ", "_") for col in df.columns if isinstance(col, str)]
    
    # Alternativas de nombres (mismo mapeo que usa cleaner.py)
    tiene_punto = any(alt in col for col in columnas_df for alt in ['punto_del_evento', 'punto_evento', 'evento', 'tipo_evento', 'punto'])
    tiene_tiempo = any(alt in col for col in columnas_df for alt in ['tiempo', 'fecha_hora', 'timestamp', 'fecha', 'hora'])
    tiene_id = any(alt in col for col in columnas_df for alt in ['id', 'cedula', 'documento', 'employee_id', 'user_id'])
    
    # Si tiene al menos el punto del evento y el ID o Tiempo, es válido
    return tiene_punto and (tiene_tiempo or tiene_id)

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