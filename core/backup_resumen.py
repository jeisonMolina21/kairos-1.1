# core/backup_resumen.py
import pandas as pd
import json
from datetime import datetime

def crear_resumen_basico(df_limpio, salida_path="outputs/resumen_asistencias.json"):
    """
    Crea un resumen básico cuando el comparator falla
    """
    print("🔄 Creando resumen básico de asistencias...")
    
    resultados = []
    
    # Agrupar por ID y fecha
    if 'tiempo' in df_limpio.columns and 'id' in df_limpio.columns:
        df_limpio['fecha'] = pd.to_datetime(df_limpio['tiempo']).dt.date
        
        for (empleado_id, fecha), grupo in df_limpio.groupby(['id', 'fecha']):
            # Ordenar por tiempo
            grupo = grupo.sort_values('tiempo')
            
            # Obtener primera y última marca
            primera_entrada = grupo['tiempo'].min()
            ultima_salida = grupo['tiempo'].max()
            
            # Calcular horas trabajadas
            horas_trabajadas = (ultima_salida - primera_entrada).total_seconds() / 3600
            
            # Buscar nombre si existe
            nombre = grupo.iloc[0]['nombre'] if 'nombre' in grupo.columns else "N/A"
            apellido = grupo.iloc[0]['apellido'] if 'apellido' in grupo.columns else "N/A"
            nombre_completo = f"{nombre} {apellido}".strip()
            
            resultados.append({
                "Empleado": nombre_completo if nombre_completo else "Desconocido",
                "Cedula": empleado_id,
                "Fecha": fecha.strftime("%Y-%m-%d"),
                "Ingreso Asignado": "08:00:00",  # Horario por defecto
                "Salida Asignada": "17:00:00",   # Horario por defecto
                "Hora Ingreso": primera_entrada.strftime("%H:%M:%S"),
                "Hora Salida": ultima_salida.strftime("%H:%M:%S"),
                "Horas Netas": round(horas_trabajadas, 2),
                "Horas Fuera": 0,
                "Horas Trabajadas": round(horas_trabajadas, 2),
                "Observaciones": "Procesado básico",
                "Tipo Evento": "N/A",
                "Hora Evento Inicio": "N/A",
                "Hora Evento Fin": "N/A",
                "Horas Evento": 0
            })
    
    # Guardar resultados
    with open(salida_path, "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=4)
    
    print(f"✅ Resumen básico creado: {len(resultados)} registros")
    return resultados