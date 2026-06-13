"""
Configuración centralizada para el proyecto Kairos v4
Sistema de Procesamiento de Asistencias
"""
import os
from dataclasses import dataclass
from typing import Dict, List
from pathlib import Path
from dotenv import load_dotenv

# Cargar variables de entorno desde .env
load_dotenv(Path(__file__).parent.parent / '.env')

# =============================================================================
# CONFIGURACIÓN DE RUTAS Y DIRECTORIOS
# =============================================================================

@dataclass
class PathConfig:
    """Configuración de rutas del proyecto"""
    # Directorio raíz del proyecto
    ROOT_DIR: Path = Path(__file__).parent.parent
    
    # Directorios principales
    OUTPUT_DIR: str = "outputs"
    TEMP_DIR: str = "temp"
    LOGS_DIR: str = "logs"
    DATOS_DIR: str = "datos"
    
    # Subdirectorios de outputs
    EXCEL_OUTPUT_DIR: str = os.path.join(OUTPUT_DIR, "temp")
    
    @classmethod
    def initialize_directories(cls):
        """Crea los directorios necesarios si no existen"""
        dirs = [
            cls.OUTPUT_DIR,
            cls.TEMP_DIR,
            cls.LOGS_DIR,
            cls.DATOS_DIR,
            cls.EXCEL_OUTPUT_DIR
        ]
        for directory in dirs:
            os.makedirs(directory, exist_ok=True)


# =============================================================================
# CONFIGURACIÓN DE API
# =============================================================================

@dataclass
class APIConfig:
    """Configuración de conexión a API"""
    # URLs de la API (primaria y secundaria para fallback)
    PRIMARY_URL: str = os.getenv("API_PRIMARY_URL", "https://api.rh360.unihorizonte.edu.co/api")
    SECONDARY_URL: str = os.getenv("API_SECONDARY_URL", "http://10.0.1.14:5000/api")
    
    # Credenciales
    EMAIL: str = os.getenv("API_EMAIL", "admin@admin.com")
    PASSWORD: str = os.getenv("API_PASSWORD", "Admin3811")
    
    # Configuración de conexión
    TIMEOUT: int = int(os.getenv("API_TIMEOUT", "60"))  # segundos
    MAX_RETRIES: int = 6
    RETRY_DELAY: int = 5  # segundos
    
    # Endpoints
    LOGIN_ENDPOINT: str = "/auth/login"
    STATISTICS_ENDPOINT: str = "/permissions/statistics"
    BULK_CREATE_ENDPOINT: str = "/attendance-reports/bulk-create"
    UPLOAD_EXCEL_ENDPOINT: str = "/attendance-reports/upload-excel"


# =============================================================================
# CONFIGURACIÓN DE HORARIOS Y UMBRALES
# =============================================================================

@dataclass
class WorkScheduleConfig:
    """Configuración de horarios y jornadas laborales"""
    # Horas de trabajo
    HORAS_JORNADA_COMPLETA: int = 8
    HORAS_SEMANALES_MINIMAS: int = 44
    HORAS_ALMUERZO_ESTANDAR: float = 1.0
    
    # Tolerancias de puntualidad (en minutos)
    TOLERANCIA_ENTRADA_TARDE: int = 5
    TOLERANCIA_ENTRADA_TEMPRANO: int = 10
    TOLERANCIA_SALIDA_TARDE: int = 10
    TOLERANCIA_SALIDA_TEMPRANO: int = 5
    
    # Horarios por defecto
    HORA_ENTRADA_DEFAULT: str = "08:00:00"
    HORA_SALIDA_DEFAULT: str = "17:00:00"
    
    # Días mínimos para análisis de puntualidad
    DIAS_MINIMOS_ANALISIS: int = 5
    
    # Hora inicio recargo nocturno (en formato decimal, ej: 19.0 para 19:00 o 7 PM)
    HORA_INICIO_RECARGO_NOCTURNO: float = 19.0


# =============================================================================
# CONFIGURACIÓN DE PUNTUALIDAD
# =============================================================================

@dataclass
class PunctualityConfig:
    """Configuración de criterios de puntualidad"""
    # Pesos para cálculo de puntuación
    PESO_TEMPRANO: int = 60
    PESO_A_TIEMPO: int = 30
    PESO_TARDE: int = 10
    
    # Umbrales de puntuación
    UMBRAL_EXCELENTE: int = 80
    UMBRAL_BUENO: int = 60
    UMBRAL_REGULAR: int = 40
    # Menor a 40 = "A Mejorar"
    
    # Porcentaje para detección de novedad orden público
    UMBRAL_ORDEN_PUBLICO: float = 0.85  # 85%


# =============================================================================
# CONFIGURACIÓN DE DEPARTAMENTOS
# =============================================================================

@dataclass
class DepartmentConfig:
    """Configuración de departamentos"""
    # Mapeo de nombres de archivo a códigos de departamento
    MAPEO_ARCHIVOS: Dict[str, str] = None
    
    # Mapeo de palabras clave a iniciales de departamento
    MAPEO_PALABRAS: Dict[str, List[str]] = None
    
    def __post_init__(self):
        if self.MAPEO_ARCHIVOS is None:
            self.MAPEO_ARCHIVOS = {
                "UNIHORIZONTE": "UH",
                "FINANCIERA": "FN",
                "TICS": "TC",
                "RECTORIA": "RC",
                "DOCENTES": "DC",
                "ASEO Y MANT": "AM",
                "REGISTRO Y CONTROL": "RYC",
                "MERCADEO": "MC"
            }
        
        if self.MAPEO_PALABRAS is None:
            self.MAPEO_PALABRAS = {
                'RH': ['RECURSOS HUMANOS', 'RRHH', 'HUMAN RESOURCES'],
                'TI': ['TECNOLOGÍA INFORMACIÓN', 'SISTEMAS', 'INFORMÁTICA', 'IT'],
                'FIN': ['FINANZAS', 'CONTABILIDAD', 'TESORERÍA'],
                'ADM': ['ADMINISTRACIÓN', 'ADMINISTRATIVO'],
                'VENT': ['VENTAS', 'COMERCIAL'],
                'MKT': ['MARKETING', 'MERCADEO'],
                'LOG': ['LOGÍSTICA', 'ALMACÉN'],
                'PROD': ['PRODUCCIÓN', 'OPERACIONES'],
                'CAL': ['CALIDAD'],
                'JUR': ['JURÍDICO', 'LEGAL'],
                'COM': ['COMUNICACIONES'],
                'DIR': ['DIRECCIÓN', 'GERENCIA'],
                'DOC': ['DOCENCIA', 'ACADÉMICO'],
                'BIEN': ['BIENESTAR'],
                'SEC': ['SECRETARÍA'],
                'MNT': ['MANTENIMIENTO']
            }


# =============================================================================
# CONFIGURACIÓN DE DÍAS Y CALENDARIO
# =============================================================================

@dataclass
class CalendarConfig:
    """Configuración de calendario y días"""
    DIAS_SEMANA: List[str] = None
    DIAS_SEMANA_API: Dict[str, str] = None
    DIAS_SEMANA_INGLES: List[str] = None
    
    def __post_init__(self):
        if self.DIAS_SEMANA is None:
            self.DIAS_SEMANA = [
                "Lunes", "Martes", "Miércoles", "Jueves", 
                "Viernes", "Sábado", "Domingo"
            ]
        
        if self.DIAS_SEMANA_API is None:
            self.DIAS_SEMANA_API = {
                'monday': 'Lunes',
                'tuesday': 'Martes',
                'wednesday': 'Miércoles',
                'thursday': 'Jueves',
                'friday': 'Viernes',
                'saturday': 'Sábado',
                'sunday': 'Domingo'
            }
        
        if self.DIAS_SEMANA_INGLES is None:
            self.DIAS_SEMANA_INGLES = [
                "Monday", "Tuesday", "Wednesday", "Thursday",
                "Friday", "Saturday", "Sunday"
            ]


# =============================================================================
# CONFIGURACIÓN DE ESTILOS EXCEL
# =============================================================================

@dataclass
class StyleConfig:
    """Configuración de colores y estilos para Excel"""
    # Colores principales
    AZUL_PRINCIPAL: str = "0054A6"
    VERDE_SUTIL: str = "C6E0B4"
    AMARILLO_SUTIL: str = "FFE699"
    ROJO_SUTIL: str = "F8CBAD"
    EMERGENCIA: str = "E9D502"
    GRIS_CLARO: str = "D9D9D9"
    BLANCO: str = "FFFFFF"
    
    # Configuración de formato
    FUENTE_PRINCIPAL: str = "Calibri"
    TAMAÑO_FUENTE_ENCABEZADO: int = 11
    TAMAÑO_FUENTE_NORMAL: int = 10


# =============================================================================
# CONFIGURACIÓN DE ARCHIVOS DE SALIDA
# =============================================================================

@dataclass
class OutputConfig:
    """Configuración de archivos de salida"""
    # Nombres de archivos
    DATOS_LIMPIOS_FILENAME: str = "datos_limpios.xlsx"
    REPORTE_COMPLETO_FILENAME: str = "Reporte_Asistencias_Completo.xlsx"
    STATISTICS_FILENAME: str = "statistics.json"
    RESUMEN_ASISTENCIAS_FILENAME: str = "resumen_asistencias.json"
    
    # Orden de hojas en el Excel final
    ORDEN_HOJAS_EXCEL: List[str] = None
    
    def __post_init__(self):
        if self.ORDEN_HOJAS_EXCEL is None:
            self.ORDEN_HOJAS_EXCEL = [
                "Diario",
                "Resumen Semanal",
                "Permisos",
                "Horarios Programados",
                "Explicación",
                "Registros Brutos"
            ]


# =============================================================================
# CONFIGURACIÓN DE LOGGING
# =============================================================================

@dataclass
class LogConfig:
    """Configuración de logging"""
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")
    LOG_MAX_BYTES: int = int(os.getenv("LOG_MAX_BYTES", 10 * 1024 * 1024))  # 10MB
    LOG_BACKUP_COUNT: int = int(os.getenv("LOG_BACKUP_COUNT", 5))
    LOG_FORMAT: str = '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
    LOG_DATE_FORMAT: str = '%Y-%m-%d %H:%M:%S'


# =============================================================================
# INSTANCIAS GLOBALES DE CONFIGURACIÓN
# =============================================================================

# Instanciar todas las configuraciones
paths = PathConfig()
api = APIConfig()
work_schedule = WorkScheduleConfig()
punctuality = PunctualityConfig()
departments = DepartmentConfig()
calendar = CalendarConfig()
styles = StyleConfig()
output = OutputConfig()
logging_config = LogConfig()


# =============================================================================
# FUNCIONES DE INICIALIZACIÓN
# =============================================================================

def initialize_app():
    """
    Inicializa la aplicación creando directorios necesarios
    """
    paths.initialize_directories()
    print("✅ Directorios inicializados")


def get_config_summary() -> str:
    """
    Retorna un resumen de la configuración actual
    """
    return f"""
🔧 CONFIGURACIÓN KAIROS V4
{'='*60}
📁 Directorios:
   - Outputs: {paths.OUTPUT_DIR}
   - Temp: {paths.TEMP_DIR}
   - Logs: {paths.LOGS_DIR}

🌐 API:
   - Primaria: {api.PRIMARY_URL}
   - Secundaria: {api.SECONDARY_URL}
   - Timeout: {api.TIMEOUT}s

⏰ Horarios:
   - Jornada completa: {work_schedule.HORAS_JORNADA_COMPLETA}h
   - Tolerancia entrada: ±{work_schedule.TOLERANCIA_ENTRADA_TARDE} min

📊 Puntualidad:
   - Pesos: {punctuality.PESO_TEMPRANO}% temprano, {punctuality.PESO_A_TIEMPO}% a tiempo, {punctuality.PESO_TARDE}% tarde
   - Umbral excelente: {punctuality.UMBRAL_EXCELENTE}%

📝 Logging:
   - Nivel: {logging_config.LOG_LEVEL}
   - Max size: {logging_config.LOG_MAX_BYTES / (1024*1024):.1f}MB
{'='*60}
"""


# Inicializar directorios al importar el módulo
initialize_app()


# =============================================================================
# BACKWARDS COMPATIBILITY (deprecar en versión futura)
# =============================================================================

class Config:
    """
    Clase de compatibilidad con versiones anteriores
    DEPRECATED: Usar las configuraciones específicas arriba
    """
    # Directorios
    OUTPUT_DIR = paths.OUTPUT_DIR
    EXCEL_OUTPUT_DIR = paths.EXCEL_OUTPUT_DIR
    TEMP_DIR = paths.TEMP_DIR
    
    # API
    API_BASE_URL = api.SECONDARY_URL.replace("/api", "")
    API_TIMEOUT = api.TIMEOUT
    API_MAX_RETRIES = api.MAX_RETRIES
    API_EMAIL = api.EMAIL
    API_PASSWORD = api.PASSWORD
    
    # Endpoints
    API_LOGIN = f"{API_BASE_URL}/api/auth/login"
    API_STATISTICS = f"{API_BASE_URL}/api/permissions/statistics"
    API_ATTENDANCE = f"{API_BASE_URL}/api/attendance-reports/bulk-create"
    
    # Horarios
    HORAS_JORNADA_COMPLETA = work_schedule.HORAS_JORNADA_COMPLETA
    HORAS_SEMANALES_MINIMAS = work_schedule.HORAS_SEMANALES_MINIMAS
    MINUTOS_TOLERANCIA_ENTRADA = work_schedule.TOLERANCIA_ENTRADA_TARDE
    MINUTOS_EARLY_ARRIVAL = work_schedule.TOLERANCIA_ENTRADA_TEMPRANO
    
    # Días
    DIAS_SEMANA = calendar.DIAS_SEMANA[:6]  # Solo hasta sábado
    DIAS_SEMANA_API = calendar.DIAS_SEMANA_API
    
    @classmethod
    def initialize_directories(cls):
        """DEPRECATED: Usar paths.initialize_directories()"""
        paths.initialize_directories()