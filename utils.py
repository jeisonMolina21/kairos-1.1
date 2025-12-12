# utils.py (COLOCAR EN LA RAÍZ DEL PROYECTO)
from core.config import work_schedule, punctuality

# Importar constantes desde configuración centralizada
DIAS_MINIMOS_ANALISIS = work_schedule.DIAS_MINIMOS_ANALISIS

def determinar_estado_general(puntuacion):
    """Determina el estado general basado en la puntuación corregida"""
    if puntuacion >= punctuality.UMBRAL_EXCELENTE:
        return "Excelente"
    elif puntuacion >= punctuality.UMBRAL_BUENO:
        return "Bueno"
    elif puntuacion >= punctuality.UMBRAL_REGULAR:
        return "Regular"
    else:
        return "A Mejorar"

def calcular_puntuacion_corregida(dias_temprano, dias_atiempo, dias_tarde, dias_trabajados):
    """
    FÓRMULA CORREGIDA: Premia puntualidad, castiga tardanza
    60% Temprano + 30% A Tiempo + 10% Tarde = 100% Puntuación
    """
    if dias_trabajados < DIAS_MINIMOS_ANALISIS:
        return 0

    dias_clasificados = dias_temprano + dias_atiempo + dias_tarde
    if dias_clasificados == 0:
        return 0

    # Pesos corregidos: Máximo premio a puntualidad, mínimo a tardanza
    puntuacion_bruta = (
        dias_temprano * punctuality.PESO_TEMPRANO + 
        dias_atiempo * punctuality.PESO_A_TIEMPO + 
        dias_tarde * punctuality.PESO_TARDE
    )
    
    # CORRECCIÓN: Máximo posible basado en días clasificados
    maximo_posible = dias_clasificados * punctuality.PESO_TEMPRANO
    
    if maximo_posible > 0:
        puntuacion_normalizada = (puntuacion_bruta / maximo_posible) * 100
        return round(puntuacion_normalizada, 1)
    return 0

def es_dia_festivo_simple(fecha):
    """
    Función simplificada para verificar días festivos
    """
    return False, ""