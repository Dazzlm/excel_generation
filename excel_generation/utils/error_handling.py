"""
Manejo centralizado de errores para el módulo de generación Excel
"""

class ExcelGenerationError(Exception):
    """Excepción base para errores de generación de Excel"""
    pass

class DatabaseError(ExcelGenerationError):
    """Excepción para errores de base de datos"""
    pass

class TemplateError(ExcelGenerationError):
    """Excepción para errores de templates"""
    pass

class ValidationError(ExcelGenerationError):
    """Excepción para errores de validación"""
    pass

def handle_database_error(func):
    """Decorador para manejar errores de base de datos"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            raise DatabaseError(f"Error en operación de base de datos: {str(e)}")
    return wrapper

def handle_excel_error(func):
    """Decorador para manejar errores de Excel"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            raise ExcelGenerationError(f"Error generando Excel: {str(e)}")
    return wrapper