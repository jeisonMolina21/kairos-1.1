"""
Módulo para transformación de datos al formato requerido por la API
"""
import json
from typing import List, Dict, Any
from core.time_utils import segundos_a_hhmmss


def transformar_para_api(resumen_path: str) -> List[Dict[str, Any]]:
    """
    Transforma el resumen de asistencias al formato requerido por la API
    
    Args:
        resumen_path: Ruta al archivo JSON de resumen de asistencias
        
    Returns:
        Lista de dictionaries con datos formateados para la API
    """
    try:
        with open(resumen_path, "r", encoding="utf-8") as f:
            resumen = json.load(f)
    except Exception as e:
        print(f"❌ Error al cargar el archivo de resumen: {e}")
        return []
    
    datos_api = []
    
    for registro in resumen:
        # Obtener información de incidentes
        incidentes = _extraer_incidentes(registro)
        
        # Convertir horas trabajadas y horas fuera a segundos
        horas_trabajadas_decimal = registro.get("Horas_Trabajadas", 0)
        horas_fuera_decimal = registro.get("Horas_Fuera", 0)
        
        # Convertir horas decimales a segundos y luego a formato HH:MM:SS
        segundos_trabajados = int(horas_trabajadas_decimal * 3600)
        segundos_fuera = int(horas_fuera_decimal * 3600)
        
        # Preparar datos para la API
        dato_api = {
            "user_id": registro.get("Cedula", ""),
            "attendance_date": registro.get("Fecha", ""),
            "arrival_time": _obtener_hora_valida(registro.get("Hora_Ingreso")),
            "exit_time": _obtener_hora_valida(registro.get("Hora_Salida")),
            "assigned_arrival_time": _obtener_hora_valida(registro.get("Ingreso_Asignado")),
            "assigned_exit_time": _obtener_hora_valida(registro.get("Salida_Asignada")),
            "total_worked_hours": segundos_a_hhmmss(segundos_trabajados),
            "out_of_institution_hours": segundos_a_hhmmss(segundos_fuera),
            "incidents": incidentes,
            "observations": registro.get("Observaciones", ""),
            "news": None  # Se puede completar si hay información de novedades
        }
        datos_api.append(dato_api)
    
    print(f"📤 Datos transformados para API: {len(datos_api)} registros")
    return datos_api


def _extraer_incidentes(registro: Dict[str, Any]) -> Dict[str, Any]:
    """
    Extrae información de incidentes desde un registro
    
    Args:
        registro: Registro de asistencia
        
    Returns:
        Dictionary con información de incidentes o None
    """
    tipo_evento = registro.get("Tipo_Evento", "N/A")
    
    if tipo_evento == "N/A":
        return None
    
    incidentes = {"tipo_evento": tipo_evento}
    
    # Agregar información específica según el tipo de evento
    if tipo_evento == "Permiso":
        tipo_permiso = registro.get("Tipo_Permiso", "N/A")
        if tipo_permiso != "N/A":
            incidentes["tipo_permiso"] = tipo_permiso
    
    elif tipo_evento == "Campaña":
        campana_entidad = registro.get("Campaña_Entidad", "N/A")
        if campana_entidad != "N/A":
            incidentes["entidad"] = campana_entidad
    
    elif tipo_evento == "Salida Institucional":
        departures_institucion = registro.get("Departures_Institucion", "N/A")
        departures_razon = registro.get("Departures_Razon", "N/A")
        
        if departures_institucion != "N/A":
            incidentes["institucion"] = departures_institucion
        if departures_razon != "N/A":
            incidentes["razon"] = departures_razon
    
    return incidentes


def _obtener_hora_valida(hora: Any) -> Any:
    """
    Retorna la hora si es válida, None si es N/A
    
    Args:
        hora: Valor de hora a validar
        
    Returns:
        Hora válida o None
    """
    if hora == "N/A" or hora is None or hora == "":
        return None
    return hora
