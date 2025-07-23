import csv
import tempfile
import logging
import os
from typing import Dict, List, Any
from ironxl import WorkBook
from sqlalchemy import text, create_engine
from sqlalchemy.ext.asyncio import AsyncEngine
from ..utils.db import DatabaseManager
from ..utils.error_handling import ValidationError, DatabaseError

logger = logging.getLogger(__name__)

class ExcelUploadService:
    """
    Servicio optimizado para cargar datos desde Excel a la base de datos usando IronXL
    y COPY/UPSERT masivo (siempre el método más rápido, sin pandas ni otras libs de procesamiento de Excel).
    """

    def __init__(self):
        self.db_manager = DatabaseManager()
        # Engine síncrono para COPY (psycopg2)
        self.sync_engine = create_engine(self.db_manager.engine.url.set(drivername="postgresql+psycopg2"))

    def _read_excel_file_ironxl(self, excel_data: bytes) -> List[Dict[str, Any]]:
        """Lee archivo Excel usando IronXL desde archivo temporal y convierte a lista de diccionarios"""
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
                tmp_file.write(excel_data)
                temp_path = tmp_file.name

            try:
                workbook = WorkBook.Load(temp_path)
                worksheet = workbook.DefaultWorkSheet

                data = []
                headers = []
                first_row = True

                for row_index in range(worksheet.RowCount):
                    row_data = {}
                    has_data = False

                    for col_index in range(worksheet.ColumnCount):
                        try:
                            cell = worksheet[f"{self._get_column_letter(col_index)}{row_index + 1}"]
                            cell_value = cell.Value

                            if cell_value is not None and str(cell_value).strip():
                                has_data = True
                                if first_row:
                                    headers.append(str(cell_value).strip())
                                else:
                                    if col_index < len(headers) and headers[col_index]:
                                        row_data[headers[col_index]] = str(cell_value).strip()
                            else:
                                if first_row:
                                    headers.append("")
                        except Exception:
                            continue

                    if first_row:
                        first_row = False
                        headers = [h for h in headers if h]
                        if not headers:
                            raise ValidationError("No se encontraron headers válidos en el archivo Excel")
                    elif has_data and row_data:
                        data.append(row_data)
                    elif not has_data and len(data) > 0:
                        break
            finally:
                try:
                    os.unlink(temp_path)
                except Exception:
                    pass

            if not data:
                raise ValidationError("El archivo Excel está vacío o no contiene datos válidos")

            logger.info(f"Excel leído con IronXL: {len(data)} filas, {len(headers)} columnas")
            logger.info(f"Headers encontrados: {headers}")

            return data

        except Exception as e:
            logger.error(f"Error leyendo Excel con IronXL: {str(e)}")
            raise ValidationError(f"Error leyendo archivo Excel: {str(e)}")

    def _get_column_letter(self, col_idx: int) -> str:
        result = ""
        col_idx += 1
        while col_idx > 0:
            col_idx -= 1
            result = chr(col_idx % 26 + ord('A')) + result
            col_idx //= 26
        return result

    async def upload_excel_to_table(
        self,
        excel_data: bytes,
        table_name: str,
        conflict_columns: List[str] = ["id"]
    ) -> dict:
        """
        Lee el Excel, convierte a CSV, usa COPY a tabla staging (real) y luego UPSERT masivo a la tabla de destino.
        Siempre usa el método más rápido (no hay modo normal/lento).
        """
        try:
            # 1. Leer Excel y convertir a lista de dicts
            rows = self._read_excel_file_ironxl(excel_data)
            if not rows:
                raise ValidationError("El archivo Excel está vacío")

            fieldnames = list(rows[0].keys())
            staging_table = f"staging_{table_name}_upsert"

            # 2. Escribir a CSV temporal
            with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode="w", encoding="utf-8", newline='') as csv_temp:
                writer = csv.DictWriter(csv_temp, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
                csv_path = csv_temp.name

            engine: AsyncEngine = self.db_manager.engine

            # 3. Crear tabla staging (real, no temporal)
            async with engine.begin() as conn:
                create_cols = ', '.join([f'"{col}" text' for col in fieldnames])
                await conn.execute(text(f'DROP TABLE IF EXISTS "{staging_table}"'))
                await conn.execute(text(f'CREATE TABLE "{staging_table}" ({create_cols})'))

            # 4. COPY a la tabla staging, con engine sync (psycopg2)
            with self.sync_engine.begin() as sync_conn:
                with open(csv_path, 'r', encoding="utf-8") as f:
                    sync_conn.connection.cursor().copy_expert(
                        f'COPY "{staging_table}" ({", ".join(fieldnames)}) FROM STDIN WITH (FORMAT CSV, HEADER TRUE, DELIMITER ",")',
                        f
                    )

            # 5. UPSERT masivo a la tabla real (con transformación de tipos)
            async with engine.begin() as conn:
                update_cols = [col for col in fieldnames if col not in conflict_columns]
                set_clause = ', '.join([f'"{col}" = EXCLUDED."{col}"' for col in update_cols])
                cast_cols = ', '.join([f'"{col}"::{await self._get_column_db_type(conn, table_name, col)}' for col in fieldnames])

                insert_sql = f'''
                INSERT INTO "{table_name}" ({", ".join([f'"{col}"' for col in fieldnames])})
                SELECT {cast_cols} FROM "{staging_table}"
                ON CONFLICT ({", ".join([f'"{col}"' for col in conflict_columns])}) DO UPDATE
                SET {set_clause}
                '''

                await conn.execute(text(insert_sql))

            # 6. DROP TABLE staging
            async with engine.begin() as conn:
                await conn.execute(text(f'DROP TABLE IF EXISTS "{staging_table}"'))

            # 7. Limpiar archivo temporal
            try:
                os.unlink(csv_path)
            except Exception:
                pass

            return {
                'success': True,
                'message': f'Archivo cargado y upsert masivo realizado con COPY ({len(rows)} filas)',
                'rows_processed': len(rows),
                'errors': None
            }
        except Exception as e:
            logger.error(f"Error en carga masiva y upsert: {str(e)}")
            return {
                'success': False,
                'message': f'Error en carga COPY+upsert: {str(e)}',
                'rows_processed': 0,
                'errors': [str(e)]
            }

    async def _get_column_db_type(self, conn, table_name: str, column: str) -> str:
        # Obtiene el tipo de dato de una columna en la tabla real
        query = text("""
            SELECT data_type
            FROM information_schema.columns
            WHERE table_name = :table_name
              AND column_name = :column
        """)
        result = await conn.execute(query, {"table_name": table_name, "column": column})
        row = result.first()
        if not row:
            return "text"
        pg_types = {
            "integer": "integer",
            "bigint": "bigint",
            "smallint": "smallint",
            "numeric": "numeric",
            "double precision": "double precision",
            "real": "real",
            "boolean": "boolean",
            "date": "date",
            "timestamp without time zone": "timestamp",
            "timestamp with time zone": "timestamptz",
            "character varying": "varchar",
            "character": "char",
            "text": "text",
        }
        return pg_types.get(row.data_type, "text")