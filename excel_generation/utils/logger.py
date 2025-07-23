import logging
import sys
from pathlib import Path

def setup_logger(log_level: str = "INFO", log_file: str = None):
    """
    Configura el sistema de logging para la aplicación
    
    Args:
        log_level: Nivel de logging (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Archivo donde guardar los logs (opcional)
    """
    
    # Crear formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Configurar logger raíz
    logger = logging.getLogger()
    logger.setLevel(getattr(logging, log_level.upper()))
    
    # Handler para consola
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Handler para archivo (opcional)
    if log_file:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    # Configurar loggers específicos
    logging.getLogger("excel_generation").setLevel(logging.INFO)
    logging.getLogger("sqlalchemy.engine").setLevel(logging.WARNING)
    
    return logger

def get_logger(name: str) -> logging.Logger:
    """Obtiene un logger configurado para un módulo específico"""
    return logging.getLogger(name)