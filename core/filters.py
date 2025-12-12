import pandas as pd
from typing import List

def clasificar_evento(punto_evento: str) -> str:
    """
    Clasifica un evento como Entrada, Salida u Otro basado en palabras clave.
    
    Args:
        punto_evento: Descripción del punto de evento
        
    Returns:
        'Entrada', 'Salida' o 'Otro'
    """
    if pd.isna(punto_evento):
        return "Otro"
    
    punto = str(punto_evento).strip().upper()
    
    # Palabras clave para entrada
    palabras_entrada = [
        "TOR IZ-ENT", "TOR DR-ENT", "ENTRADA", "ENTRY", "IN", "INGRESO",
        "ENT", "CHECK-IN", "ARRIVAL", "ACCESS IN", "ENTRADA PRINCIPAL"
    ]
    
    # Palabras clave para salida
    palabras_salida = [
        "TOR IZ-SAL", "TOR DR-SAL", "SALIDA", "EXIT", "OUT", "EGRESO", 
        "SAL", "CHECK-OUT", "DEPARTURE", "ACCESS OUT", "SALIDA PRINCIPAL"
    ]
    
    # Buscar coincidencias
    for palabra in palabras_entrada:
        if palabra in punto:
            return "Entrada"
    
    for palabra in palabras_salida:
        if palabra in punto:
            return "Salida"
    
    return "Otro"

def eliminar_duplicados_temporales(df: pd.DataFrame, segundos_umbral: int = 10) -> pd.DataFrame:
    """
    Elimina eventos duplicados que ocurren dentro de un umbral de tiempo.
    
    Args:
        df: DataFrame con datos de eventos
        segundos_umbral: Umbral en segundos para considerar duplicados
        
    Returns:
        DataFrame sin duplicados temporales
    """
    if df.empty or "id" not in df.columns or "tiempo" not in df.columns:
        return df
    
    df_ordenado = df.sort_values(["id", "tiempo"]).copy()
    indices_a_eliminar = []
    
    print(f"🧹 Eliminando duplicados (umbral: {segundos_umbral}s)...")
    
    for usuario_id, grupo in df_ordenado.groupby("id"):
        tiempos = grupo["tiempo"].tolist()
        indices = grupo.index.tolist()
        
        for i in range(1, len(tiempos)):
            diferencia = (tiempos[i] - tiempos[i-1]).total_seconds()
            if diferencia <= segundos_umbral:
                indices_a_eliminar.append(indices[i])
    
    df_limpio = df_ordenado.drop(indices_a_eliminar)
    print(f"✅ Eliminados {len(indices_a_eliminar)} duplicados")
    
    return df_limpio

def filtrar_por_ids_validos(df: pd.DataFrame, ids_validos: List[str]) -> pd.DataFrame:
    """
    Filtra el DataFrame para incluir solo IDs válidos.
    
    Args:
        df: DataFrame a filtrar
        ids_validos: Lista de IDs válidos
        
    Returns:
        DataFrame filtrado
    """
    if not ids_validos:
        return df
    
    # Buscar columna de ID
    columnas_id = ['id', 'cedula', 'documento', 'employee_id', 'user_id']
    columna_id = None
    
    for col in columnas_id:
        if col in df.columns:
            columna_id = col
            break
    
    if columna_id:
        # Normalizar IDs
        df[columna_id] = df[columna_id].astype(str).str.strip()
        ids_validos_str = [str(id_val).strip() for id_val in ids_validos]
        
        # Aplicar filtro
        df_filtrado = df[df[columna_id].isin(ids_validos_str)].copy()
        print(f"✅ Filtrado por IDs válidos: {len(df_filtrado)} filas")
        return df_filtrado
    
    print("⚠️ No se encontró columna de ID para filtrar")
    return df