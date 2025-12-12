"""
Sistema de logging profesional con rotación de archivos.
"""
import logging
import os
from logging.handlers import RotatingFileHandler
from typing import Optional


def setup_logger(
    name: str,
    log_dir: str = "logs",
    log_level: str = "INFO",
    max_bytes: int = 10 * 1024 * 1024,  # 10MB
    backup_count: int = 5
) -> logging.Logger:
    """
    Configura un logger con salida a archivo y consola.
    
    Args:
        name: Nombre del logger
        log_dir: Directorio donde guardar los archivos de log
        log_level: Nivel de logging (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        max_bytes: Tamaño máximo de cada archivo de log en bytes
        backup_count: Número de archivos de respaldo a mantener
        
    Returns:
        Logger configurado
        
    Example:
        >>> logger = setup_logger("kairos")
        >>> logger.info("Aplicación iniciada")
    """
    # Crear directorio de logs si no existe
    os.makedirs(log_dir, exist_ok=True)
    
    # Crear logger
    logger = logging.getLogger(name)
    
    # Si el logger ya tiene handlers, no agregamos más
    if logger.handlers:
        return logger
    
    # Configurar nivel
    numeric_level = getattr(logging, log_level.upper(), logging.INFO)
    logger.setLevel(numeric_level)
    
    # Formato detallado para archivo
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Formato simple para consola
    console_formatter = logging.Formatter(
        '%(levelname)s: %(message)s'
    )
    
    # Handler para archivo con rotación
    log_file = os.path.join(log_dir, f"{name}.log")
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(file_formatter)
    
    # Handler para consola
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(console_formatter)
    
    # Agregar handlers al logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Obtiene un logger existente o crea uno nuevo.
    
    Args:
        name: Nombre del logger. Si es None, usa "kairos"
        
    Returns:
        Logger configurado
    """
    logger_name = name or "kairos"
    
    # Si el logger ya existe, retornarlo
    if logger_name in logging.Logger.manager.loggerDict:
        return logging.getLogger(logger_name)
    
    # Si no existe, crear uno nuevo
    return setup_logger(logger_name)
