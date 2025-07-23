import asyncio
from typing import Dict, List, Any, Optional
import logging
import os
from sqlalchemy import create_engine, text, MetaData, inspect
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession
from sqlalchemy.orm import sessionmaker

from .error_handling import DatabaseError

logger = logging.getLogger(__name__)

class DatabaseManager:
    """Manejador de conexiones y operaciones de base de datos"""
    
    def __init__(self):
        self.database_url = self._get_database_url()
        self.engine = create_async_engine(
            self.database_url,
            echo=False,
            pool_size=10,
            max_overflow=20
        )
        self.session_factory = sessionmaker(
            self.engine, 
            class_=AsyncSession, 
            expire_on_commit=False
        )
    
    def _get_database_url(self) -> str:
        """Construye la URL de conexión a la base de datos"""
        host = os.getenv("DB_HOST", "localhost")
        port = os.getenv("DB_PORT", "5432")
        database = os.getenv("DB_NAME", "postgres")
        username = os.getenv("DB_USER", "postgres")
        password = os.getenv("DB_PASSWORD", "password")
        
        return f"postgresql+asyncpg://{username}:{password}@{host}:{port}/{database}"
    
    async def get_table_columns(self, table_name: str) -> List[str]:
        """Obtiene las columnas de una tabla"""
        try:
            async with self.session_factory() as session:
                query = text("""
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name = :table_name
                    ORDER BY ordinal_position
                """)
                
                result = await session.execute(query, {"table_name": table_name})
                columns = [row[0] for row in result]
                
                if not columns:
                    raise DatabaseError(f"Tabla '{table_name}' no encontrada")
                
                return columns
                
        except Exception as e:
            logger.error(f"Error obteniendo columnas de tabla {table_name}: {str(e)}")
            raise DatabaseError(f"Error accediendo a tabla: {str(e)}")
    
    async def fetch_data(
        self, 
        table_name: str, 
        fields: List[str], 
        filters: Dict[str, Any]
    ) -> List[Dict]:
        """Obtiene datos de la tabla con filtros aplicados"""
        try:
            # Construir consulta SELECT
            fields_str = ", ".join(f'"{field}"' for field in fields)
            query_str = f'SELECT {fields_str} FROM "{table_name}"'
            
            # Agregar filtros WHERE
            where_conditions = []
            params = {}
            
            for key, value in filters.items():
                if value is not None:
                    where_conditions.append(f'"{key}" = :{key}')
                    params[key] = value
            
            if where_conditions:
                query_str += " WHERE " + " AND ".join(where_conditions)
            
            # Limitar resultados para prevenir sobrecarga
            query_str += " LIMIT 200000"
            
            logger.info(f"Ejecutando consulta: {query_str}")
            
            async with self.session_factory() as session:
                result = await session.execute(text(query_str), params)
                
                # Convertir a lista de diccionarios
                rows = result.fetchall()
                columns = result.keys()
                
                data = [dict(zip(columns, row)) for row in rows]
                
                logger.info(f"Obtenidas {len(data)} filas de {table_name}")
                return data
                
        except Exception as e:
            logger.error(f"Error obteniendo datos de {table_name}: {str(e)}")
            raise DatabaseError(f"Error en consulta: {str(e)}")
    
    async def test_connection(self) -> bool:
        """Prueba la conexión a la base de datos"""
        try:
            async with self.session_factory() as session:
                await session.execute(text("SELECT 1"))
            return True
        except Exception as e:
            logger.error(f"Error probando conexión: {str(e)}")
            return False
    
    async def close(self):
        """Cierra las conexiones de la base de datos"""
        await self.engine.dispose()