# core/calculations.py
import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Tuple, Dict, List, Any, Optional
import sys
import os

# Agregar el directorio actual al path para importaciones
current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, current_dir)

from utils import segundos_a_hhmmss, hhmmss_a_segundos, calcular_semana_mes
from filters import clasificar_evento

def procesar_calculos(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, List[Dict]]:
    """
    Procesa cálculos avanzados de asistencia.
    
    Args:
        df: DataFrame con datos limpios
        
    Returns:
        Tuple con: datos procesados, resumen diario, datos para API
    """
    print("🔄 Iniciando procesamiento de cálculos avanzados...")
    
    # Preparar datos
    df_procesado = _preparar_datos(df)
    
    # Procesar por empleado y fecha
    registros_resumen = []
    registros_api = []
    
    for (empleado_id, fecha), grupo in df_procesado.groupby(['id', 'fecha']):
        registro = _procesar_grupo_diario(grupo, empleado_id, fecha)
        if registro:
            registros_resumen.append(registro['resumen'])
            registros_api.append(registro['api'])
    
    # Crear DataFrames finales
    df_resumen = pd.DataFrame(registros_resumen) if registros_resumen else pd.DataFrame()
    json_api = registros_api
    
    print(f"✅ Procesamiento completado: {len(registros_resumen)} registros")
    return df_procesado, df_resumen, json_api

def _preparar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara los datos para procesamiento"""
    df_preparado = df.copy()
    
    # Asegurar columnas necesarias
    if 'tipo_evento' not in df_preparado.columns:
        df_preparado['tipo_evento'] = df_preparado['punto_del_evento'].apply(clasificar_evento)
    
    # Crear columnas derivadas
    df_preparado['fecha'] = df_preparado['tiempo'].dt.date
    df_preparado['hora'] = df_preparado['tiempo'].dt.time
    df_preparado['semana_mes'] = df_preparado['tiempo'].apply(calcular_semana_mes)
    
    return df_preparado

def _procesar_grupo_diario(grupo: pd.DataFrame, empleado_id: str, fecha: datetime) -> Optional[Dict[str, Any]]:
    """Procesa un grupo de registros de un día para un empleado"""
    # Ordenar por tiempo
    grupo_ordenado = grupo.sort_values('tiempo')
    
    # Separar entradas y salidas
    entradas = grupo_ordenado[grupo_ordenado['tipo_evento'] == 'Entrada']
    salidas = grupo_ordenado[grupo_ordenado['tipo_evento'] == 'Salida']
    
    if entradas.empty or salidas.empty:
        return None
    
    # Obtener primera entrada y última salida
    primera_entrada = entradas['tiempo'].iloc[0]
    ultima_salida = salidas['tiempo'].iloc[-1]
    
    # Calcular horas
    horas_trabajadas, horas_descanso, horas_netas = _calcular_horas_diarias(
        grupo_ordenado, primera_entrada, ultima_salida
    )
    
    # Obtener horario programado (simplificado - usar valores por defecto)
    horario_entrada, horario_salida = _obtener_horario_programado(empleado_id, fecha)
    
    # Generar observaciones
    observaciones = _generar_observaciones(primera_entrada, horario_entrada, ultima_salida, horario_salida)
    
    # Preparar registros
    registro_resumen = {
        'ID': empleado_id,
        'Nombre': grupo_ordenado['nombre'].iloc[0] if 'nombre' in grupo_ordenado.columns else 'N/A',
        'Apellido': grupo_ordenado['apellido'].iloc[0] if 'apellido' in grupo_ordenado.columns else 'N/A',
        'Fecha': fecha.strftime('%Y-%m-%d'),
        'Entrada': primera_entrada.strftime('%H:%M:%S'),
        'Salida': ultima_salida.strftime('%H:%M:%S'),
        'HorasTrabajadas': segundos_a_hhmmss(horas_trabajadas),
        'HorasDescanso': segundos_a_hhmmss(horas_descanso),
        'HorasNetas': segundos_a_hhmmss(horas_netas),
        'IngresoAsignado': horario_entrada,
        'SalidaAsignada': horario_salida
    }
    
    # NUEVA ESTRUCTURA PARA LA API - ACTUALIZADA
    registro_api = {
        'user_id': empleado_id,
        'attendance_date': fecha.strftime('%Y-%m-%d'),
        'arrival_time': primera_entrada.strftime('%H:%M:%S'),
        'exit_time': ultima_salida.strftime('%H:%M:%S'),
        'assigned_arrival_time': horario_entrada,
        'assigned_exit_time': horario_salida,
        'total_worked_hours': round(horas_netas/ 3600, 2),  # Convertir a horas decimales
        'out_of_institution_hours': round(horas_descanso / 3600, 2),  # Convertir a horas decimales
        'incidents': None,  # Se llenará posteriormente con datos de statistics
        'observations': observaciones,
        'news': None  # Se llenará posteriormente con datos de statistics
    }
    
    return {'resumen': registro_resumen, 'api': registro_api}

def _calcular_horas_diarias(grupo: pd.DataFrame, entrada: datetime, salida: datetime) -> Tuple[int, int, int]:
    """Calcula horas trabajadas, descanso y netas para un día"""
    # Horas totales entre entrada y salida
    horas_totales = int((salida - entrada).total_seconds())
    
    # Calcular periodos de trabajo (simplificado)
    horas_trabajadas = horas_totales
    horas_descanso = 0
    
    # Ajustar por descansos detectados
    for i in range(len(grupo) - 1):
        evento_actual = grupo.iloc[i]
        evento_siguiente = grupo.iloc[i + 1]
        
        if (evento_actual['tipo_evento'] == 'Salida' and 
            evento_siguiente['tipo_evento'] == 'Entrada'):
            
            descanso = (evento_siguiente['tiempo'] - evento_actual['tiempo']).total_seconds()
            if descanso > 0:
                horas_descanso += int(descanso)
    
    # Horas netas (trabajadas - descanso)
    horas_netas = max(horas_trabajadas - horas_descanso, 0)
    
    return horas_trabajadas, horas_descanso, horas_netas

def _obtener_horario_programado(empleado_id: str, fecha: datetime) -> Tuple[str, str]:
    """Obtiene el horario programado para un empleado en una fecha específica"""
    # Por ahora, usar horario por defecto
    # En una implementación real, esto vendría de la API o base de datos
    return "08:00:00", "17:00:00"

def _generar_observaciones(entrada_real: datetime, entrada_asignada: str, 
                          salida_real: datetime, salida_asignada: str) -> str:
    """Genera observaciones basadas en la puntualidad"""
    observaciones = []
    
    try:
        # Convertir horarios asignados a objetos datetime para comparación
        entrada_asignada_dt = datetime.strptime(entrada_asignada, "%H:%M:%S")
        salida_asignada_dt = datetime.strptime(salida_asignada, "%H:%M:%S")
        
        # Comparar entrada
        diferencia_entrada = (entrada_real - datetime.combine(entrada_real.date(), entrada_asignada_dt.time())).total_seconds() / 60
        
        if diferencia_entrada < -10:  # Más de 10 minutos temprano
            observaciones.append("Llegó temprano")
        elif diferencia_entrada > 5:  # Más de 5 minutos tarde
            observaciones.append("Llegó tarde")
        else:
            observaciones.append("Llegó a tiempo")
        
        # Comparar salida
        diferencia_salida = (salida_real - datetime.combine(salida_real.date(), salida_asignada_dt.time())).total_seconds() / 60
        
        if diferencia_salida < -5:  # Más de 5 minutos temprano
            observaciones.append("Salió temprano")
        elif diferencia_salida > 10:  # Más de 10 minutos tarde
            observaciones.append("Salió tarde")
        else:
            observaciones.append("Salió a tiempo")
            
    except Exception as e:
        observaciones.append("Error calculando puntualidad")
    
    return "; ".join(observaciones) if observaciones else "Sin observaciones"