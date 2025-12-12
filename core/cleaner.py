import pandas as pd
import numpy as np
from typing import Dict, Any
from .filters import clasificar_evento, eliminar_duplicados_temporales

def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia y prepara los datos para procesamiento.
    
    Args:
        df: DataFrame con datos brutos
        
    Returns:
        DataFrame limpio y normalizado
    """
    print("\n🧹 Iniciando limpieza de datos...")
    print(f"📋 Columnas detectadas: {list(df.columns)}")
    
    # Verificar columnas obligatorias
    df = _verificar_columnas_obligatorias(df)
    
    # Normalizar datos
    df = _normalizar_datos(df)
    
    # Filtrar eventos relevantes
    df = _filtrar_eventos_relevantes(df)
    
    # Eliminar duplicados
    df = eliminar_duplicados_temporales(df)
    
    # Ordenar datos
    df = _ordenar_datos(df)
    
    print(f"✅ Limpieza completada: {len(df)} filas válidas")
    return df

def _verificar_columnas_obligatorias(df: pd.DataFrame) -> pd.DataFrame:
    """Verifica y adapta las columnas obligatorias"""
    mapeo_columnas = {
        'punto_del_evento': ['punto_evento', 'evento', 'tipo_evento', 'punto'],
        'tiempo': ['fecha_hora', 'timestamp', 'fecha', 'hora'],
        'id': ['cedula', 'documento', 'employee_id', 'user_id'],
        'nombre_de_lector': ['lector', 'dispositivo', 'ubicacion']
    }
    
    for col_obligatoria, alternativas in mapeo_columnas.items():
        if col_obligatoria not in df.columns:
            for alternativa in alternativas:
                if alternativa in df.columns:
                    df = df.rename(columns={alternativa: col_obligatoria})
                    break
            else:
                raise KeyError(f"No se encontró columna: {col_obligatoria}")
    
    return df

def _normalizar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza los datos del DataFrame"""
    # Normalizar punto del evento
    df['punto_del_evento'] = df['punto_del_evento'].astype(str).str.upper().str.strip()
    
    # Convertir tiempo a datetime
    df['tiempo'] = pd.to_datetime(df['tiempo'], errors='coerce')
    
    # Eliminar registros con tiempo inválido
    df = df.dropna(subset=['tiempo'])
    
    # Normalizar ID
    if 'id' in df.columns:
        df['id'] = df['id'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    
    return df

def _filtrar_eventos_relevantes(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra solo eventos de entrada y salida relevantes"""
    # Clasificar eventos
    df['tipo_evento'] = df['punto_del_evento'].apply(clasificar_evento)
    
    # Filtrar solo entradas y salidas
    df_filtrado = df[df['tipo_evento'].isin(['Entrada', 'Salida'])].copy()
    
    print(f"➡️ Eventos relevantes: {len(df_filtrado)}/{len(df)}")
    return df_filtrado

def _ordenar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """Ordena los datos de forma consistente"""
    columnas_orden = ['id', 'nombre', 'apellido', 'tiempo']
    columnas_existentes = [col for col in columnas_orden if col in df.columns]
    
    return df.sort_values(by=columnas_existentes).reset_index(drop=True)