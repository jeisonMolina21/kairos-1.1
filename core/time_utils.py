"""
Módulo de utilidades de tiempo para el proyecto Kairos v4
Funciones para conversión de formatos de tiempo, cálculo de duraciones y puntualidad
"""
from datetime import datetime, time, timedelta
from typing import Tuple, Optional, Union
from core.config import work_schedule


# =============================================================================
# CONVERSIONES DE TIEMPO
# =============================================================================

def segundos_a_hhmmss(segundos_totales: int) -> str:
    """
    Convierte segundos a formato HH:MM:SS
    
    Args:
        segundos_totales: Número de segundos a convertir
        
    Returns:
        String en formato HH:MM:SS
        
    Example:
        >>> segundos_a_hhmmss(3665)
        '01:01:05'
    """
    if not segundos_totales or segundos_totales == 0:
        return "00:00:00"
    
    segundos_totales = int(segundos_totales)
    horas = segundos_totales // 3600
    minutos = (segundos_totales % 3600) // 60
    segundos = segundos_totales % 60
    
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"


def hhmmss_a_segundos(tiempo_str: str) -> int:
    """
    Convierte formato HH:MM:SS a segundos
    
    Args:
        tiempo_str: String en formato HH:MM:SS o H:MM:SS
        
    Returns:
        Número total de segundos
        
    Example:
        >>> hhmmss_a_segundos('01:30:00')
        5400
    """
    if not tiempo_str or tiempo_str in ["N/A", "", "None"]:
        return 0
    
    try:
        partes = tiempo_str.strip().split(':')
        if len(partes) != 3:
            return 0
        
        horas, minutos, segundos = map(int, partes)
        return horas * 3600 + minutos * 60 + segundos
    except (ValueError, AttributeError):
        return 0


def decimal_a_hhmmss(horas_decimal: float) -> str:
    """
    Convierte horas decimales a formato HH:MM:SS
    
    Args:
        horas_decimal: Horas en formato decimal (ej: 8.5 = 8h 30m)
        
    Returns:
        String en formato HH:MM:SS
        
    Example:
        >>> decimal_a_hhmmss(8.5)
        '08:30:00'
    """
    if not horas_decimal or horas_decimal == 0:
        return "00:00:00"
    
    segundos_totales = int(horas_decimal * 3600)
    return segundos_a_hhmmss(segundos_totales)


def hhmmss_a_decimal(tiempo_str: str) -> float:
    """
    Convierte formato HH:MM:SS a horas decimales
    
    Args:
        tiempo_str: String en formato HH:MM:SS
        
    Returns:
        Horas en formato decimal
        
    Example:
        >>> hhmmss_a_decimal('08:30:00')
        8.5
    """
    segundos = hhmmss_a_segundos(tiempo_str)
    return round(segundos / 3600, 2)


# =============================================================================
# PARSING DE TIEMPO
# =============================================================================

def parse_time(value: Union[str, time, None]) -> Optional[time]:
    """
    Convierte texto HH:MM:SS a datetime.time o None
    
    Args:
        value: Valor a parsear (string, time object, o None)
        
    Returns:
        Objeto datetime.time o None si hay error
    """
    if value is None or value == "N/A" or value == "":
        return None
    
    if isinstance(value, time):
        return value
    
    try:
        if isinstance(value, str):
            # Intentar parsear formato HH:MM:SS
            partes = value.strip().split(':')
            if len(partes) == 3:
                h, m, s = map(int, partes)
                return time(h, m, s)
            elif len(partes) == 2:
                h, m = map(int, partes)
                return time(h, m, 0)
    except (ValueError, AttributeError):
        pass
    
    return None


def parse_datetime(fecha_str: str, hora_str: str) -> Optional[datetime]:
    """
    Combina fecha y hora en un objeto datetime
    
    Args:
        fecha_str: Fecha en formato YYYY-MM-DD
        hora_str: Hora en formato HH:MM:SS
        
    Returns:
        Objeto datetime o None si hay error
    """
    try:
        fecha = datetime.strptime(fecha_str, "%Y-%m-%d").date()
        hora = parse_time(hora_str)
        
        if hora:
            return datetime.combine(fecha, hora)
    except (ValueError, AttributeError):
        pass
    
    return None


# =============================================================================
# CÁLCULOS DE DURACIÓN
# =============================================================================

def calcular_duracion_horas(inicio: time, fin: time) -> float:
    """
    Calcula la duración en horas entre dos objetos time
    
    Args:
        inicio: Hora de inicio
        fin: Hora de fin
        
    Returns:
        Duración en horas (decimal)
    """
    if not inicio or not fin:
        return 0.0
    
    # Crear datetimes para facilitar el cálculo
    fecha_base = datetime(2000, 1, 1)
    dt_inicio = datetime.combine(fecha_base, inicio)
    dt_fin = datetime.combine(fecha_base, fin)
    
    # Si fin es menor que inicio, asumimos que cruza medianoche
    if dt_fin < dt_inicio:
        dt_fin += timedelta(days=1)
    
    diferencia = dt_fin - dt_inicio
    return diferencia.total_seconds() / 3600


def calcular_duracion_segundos(inicio: time, fin: time) -> int:
    """
    Calcula la duración en segundos entre dos objetos time
    
    Args:
        inicio: Hora de inicio
        fin: Hora de fin
        
    Returns:
        Duración en segundos
    """
    horas = calcular_duracion_horas(inicio, fin)
    return int(horas * 3600)


# =============================================================================
# ANÁLISIS DE PUNTUALIDAD
# =============================================================================

def determinar_puntualidad(
    hora_real: Union[str, time],
    hora_asignada: Union[str, time],
    tipo: str = "entrada"
) -> str:
    """
    Determina si fue temprano, a tiempo o tarde
    
    Args:
        hora_real: Hora real registrada
        hora_asignada: Hora asignada/programada
        tipo: "entrada" o "salida"
        
    Returns:
        "Temprano", "A Tiempo", o "Tarde"
    """
    # Parsear las horas
    if isinstance(hora_real, str):
        hora_real = parse_time(hora_real)
    if isinstance(hora_asignada, str):
        hora_asignada = parse_time(hora_asignada)
    
    if not hora_real or not hora_asignada:
        return "N/A"
    
    # Calcular diferencia en minutos
    fecha_base = datetime(2000, 1, 1)
    dt_real = datetime.combine(fecha_base, hora_real)
    dt_asignada = datetime.combine(fecha_base, hora_asignada)
    
    diferencia_minutos = (dt_real - dt_asignada).total_seconds() / 60
    
    if tipo == "entrada":
        # Para entrada: temprano si llega más de 10 min antes
        if diferencia_minutos < -work_schedule.TOLERANCIA_ENTRADA_TEMPRANO:
            return "Temprano"
        # Tarde si llega más de 5 min después
        elif diferencia_minutos > work_schedule.TOLERANCIA_ENTRADA_TARDE:
            return "Tarde"
        else:
            return "A Tiempo"
    else:  # salida
        # Para salida: temprano si sale más de 5 min antes
        if diferencia_minutos < -work_schedule.TOLERANCIA_SALIDA_TEMPRANO:
            return "Temprano"
        # Tarde si sale más de 10 min después
        elif diferencia_minutos > work_schedule.TOLERANCIA_SALIDA_TARDE:
            return "Tarde"
        else:
            return "A Tiempo"


def calcular_minutos_diferencia(
    hora_real: Union[str, time],
    hora_asignada: Union[str, time]
) -> int:
    """
    Calcula la diferencia en minutos entre dos horas
    
    Args:
        hora_real: Hora real
        hora_asignada: Hora asignada
        
    Returns:
        Diferencia en minutos (positivo si llegó tarde, negativo si llegó temprano)
    """
    if isinstance(hora_real, str):
        hora_real = parse_time(hora_real)
    if isinstance(hora_asignada, str):
        hora_asignada = parse_time(hora_asignada)
    
    if not hora_real or not hora_asignada:
        return 0
    
    fecha_base = datetime(2000, 1, 1)
    dt_real = datetime.combine(fecha_base, hora_real)
    dt_asignada = datetime.combine(fecha_base, hora_asignada)
    
    return int((dt_real - dt_asignada).total_seconds() / 60)


# =============================================================================
# CÁLCULOS DE CALENDARIO
# =============================================================================

def calcular_semana_mes(fecha: datetime) -> int:
    """
    Calcula la semana del mes (1-5) para una fecha dada
    
    Args:
        fecha: Objeto datetime
        
    Returns:
        Número de semana del mes (1-5)
    """
    primer_dia = fecha.replace(day=1)
    dia_del_mes = fecha.day
    
    # Calcular cuántos días hay hasta el primer lunes
    dias_hasta_lunes = (7 - primer_dia.weekday()) % 7
    
    # Si el mes no comienza en lunes, ajustar
    if dias_hasta_lunes > 0:
        semana = ((dia_del_mes + primer_dia.weekday()) // 7) + 1
    else:
        semana = (dia_del_mes - 1) // 7 + 1
    
    return min(semana, 5)  # Máximo 5 semanas


def obtener_nombre_dia(fecha: Union[datetime, str]) -> str:
    """
    Obtiene el nombre del día en español para una fecha
    
    Args:
        fecha: Objeto datetime o string en formato YYYY-MM-DD
        
    Returns:
        Nombre del día en español (ej: "Lunes")
    """
    from core.config import calendar
    
    if isinstance(fecha, str):
        try:
            fecha = datetime.strptime(fecha, "%Y-%m-%d")
        except ValueError:
            return "N/A"
    
    # weekday() retorna 0=Lunes, 6=Domingo
    dia_numero = fecha.weekday()
    return calendar.DIAS_SEMANA[dia_numero]


# =============================================================================
# VALIDACIÓN
# =============================================================================

def validar_tiempo_formato(tiempo_str: str) -> bool:
    """
    Valida si un string está en formato HH:MM:SS válido
    
    Args:
        tiempo_str: String a validar
        
    Returns:
        True si es válido, False si no
    """
    if not tiempo_str or tiempo_str in ["N/A", "", "None"]:
        return False
    
    try:
        partes = tiempo_str.strip().split(':')
        if len(partes) != 3:
            return False
        
        horas, minutos, segundos = map(int, partes)
        
        # Validar rangos
        if not (0 <= horas <= 23 and 0 <= minutos <= 59 and 0 <= segundos <= 59):
            return False
        
        return True
    except (ValueError, AttributeError):
        return False
