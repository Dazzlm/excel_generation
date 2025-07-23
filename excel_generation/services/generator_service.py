import asyncio
from typing import Dict, List, Optional, Any
import logging
from pathlib import Path
import os
from ironxl import WorkBook
import tempfile
from ..utils.db import DatabaseManager
from ..utils.error_handling import ExcelGenerationError, DatabaseError
from sqlalchemy import text

logger = logging.getLogger(__name__)

class ExcelGeneratorService:
    """Servicio optimizado para generar archivos Excel usando IronXL"""

    def __init__(self):
        self.db_manager = DatabaseManager()
        self.templates_path = Path(__file__).parent.parent / "templates"

    async def generate_excel(
        self,
        table: str,
        fields: List[str],
        filters: Optional[Dict] = None,
        template: str = "default_template.xlsx"
    ) -> bytes:
        try:
            await self._validate_fields(table, fields)
            data = await self._fetch_data(table, fields, filters or {})
            excel_buffer = await self._create_excel_file(data, fields, template)
            logger.info(f"Excel generado exitosamente para tabla {table} con {len(data)} filas")
            return excel_buffer
        except Exception as e:
            logger.error(f"Error generando Excel: {str(e)}")
            raise ExcelGenerationError(f"Error generando Excel: {str(e)}")

    async def generate_full_table_excel(
        self,
        table: str,
        template: str = "default_template.xlsx"
    ) -> bytes:
        try:
            fields = await self.db_manager.get_table_columns(table)
            data = await self._fetch_data(table, fields, {})
            excel_buffer = await self._create_excel_file(data, fields, template)
            logger.info(f"Excel de tabla completa generado para {table} ({len(data)} filas)")
            return excel_buffer
        except Exception as e:
            logger.error(f"Error generando Excel de tabla completa: {str(e)}")
            raise ExcelGenerationError(f"Error generando Excel de tabla completa: {str(e)}")

    async def generate_full_database_excel(
        self,
        template: str = "default_template.xlsx"
    ) -> bytes:
        """Genera un Excel con toda la base de datos. Cada tabla va en una hoja distinta."""
        try:
            async with self.db_manager.session_factory() as session:
                result = await session.execute(
                    text("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'")
                )
                table_names = [row[0] for row in result.fetchall()]

            workbook = WorkBook.Create()
            first_sheet = True

            for table in table_names:
                fields = await self.db_manager.get_table_columns(table)
                data = await self._fetch_data(table, fields, {})
                if first_sheet:
                    worksheet = workbook.DefaultWorkSheet
                    worksheet.Name = table
                    first_sheet = False
                else:
                    worksheet = workbook.CreateWorkSheet(table)
                await self._write_data_to_worksheet(worksheet, data, fields, fast_mode=True)

            excel_bytes = await self._save_workbook_to_bytes(workbook)
            logger.info(f"Excel de toda la base generado con {len(table_names)} hojas")
            return excel_bytes
        except Exception as e:
            logger.error(f"Error generando Excel de toda la base: {str(e)}")
            raise ExcelGenerationError(f"Error generando Excel de toda la base: {str(e)}")

    async def _validate_fields(self, table: str, fields: List[str]):
        try:
            table_columns = await self.db_manager.get_table_columns(table)
            invalid_fields = [field for field in fields if field not in table_columns]
            if invalid_fields:
                raise ExcelGenerationError(
                    f"Campos inválidos para tabla {table}: {invalid_fields}"
                )
        except DatabaseError as e:
            raise ExcelGenerationError(f"Error validando campos: {str(e)}")

    async def _fetch_data(self, table: str, fields: List[str], filters: Dict) -> List[Dict]:
        try:
            return await self.db_manager.fetch_data(table, fields, filters)
        except DatabaseError as e:
            raise ExcelGenerationError(f"Error obteniendo datos: {str(e)}")

    async def _create_excel_file(
        self,
        data: List[Dict],
        fields: List[str],
        template: str,
        fast_mode: bool = True
    ) -> bytes:
        try:
            template_path = self.templates_path / template
            if template_path.exists():
                workbook = WorkBook.Load(str(template_path))
                worksheet = workbook.DefaultWorkSheet if workbook.WorkSheets.Count > 0 else workbook.CreateWorkSheet("Data")
                self._clear_data_preserve_format(worksheet, len(fields))
            else:
                workbook = WorkBook.Create()
                worksheet = workbook.DefaultWorkSheet if workbook.WorkSheets.Count > 0 else workbook.CreateWorkSheet("Data")
                worksheet.Name = "Data"
            await self._write_data_to_worksheet(worksheet, data, fields, fast_mode)
            excel_bytes = await self._save_workbook_to_bytes(workbook)
            return excel_bytes
        except Exception as e:
            logger.error(f"Error creando archivo Excel con IronXL: {str(e)}")
            raise ExcelGenerationError(f"Error creando Excel: {str(e)}")

    async def _save_workbook_to_bytes(self, workbook) -> bytes:
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
                temp_path = tmp_file.name
            workbook.SaveAs(temp_path)
            with open(temp_path, 'rb') as f:
                excel_bytes = f.read()
            try:
                os.unlink(temp_path)
            except Exception:
                pass
            logger.info(f"Workbook saved successfully, size: {len(excel_bytes)} bytes")
            return excel_bytes
        except Exception as e:
            logger.error(f"Error saving workbook to bytes: {str(e)}")
            raise ExcelGenerationError(f"Error saving Excel to bytes: {str(e)}")

    def _clear_data_preserve_format(self, worksheet, num_columns: int):
        try:
            for row_index in range(1, 2000):
                all_empty = True
                for col_index in range(num_columns):
                    cell_address = f"{self._get_column_letter(col_index)}{row_index + 1}"
                    try:
                        cell = worksheet[cell_address]
                        if cell.Value:
                            cell.Value = ""
                            all_empty = False
                    except Exception:
                        continue
                if all_empty and row_index > 15:
                    break
        except Exception as e:
            logger.warning(f"Error limpiando datos del template: {str(e)}")

    async def _write_data_to_worksheet(self, worksheet, data: List[Dict], fields: List[str], fast_mode: bool = True):
        try:
            # Escribir encabezados
            for col_idx, field in enumerate(fields):
                cell_address = f"{self._get_column_letter(col_idx)}1"
                cell = worksheet[cell_address]
                cell.Value = str(field)
                try:
                    cell.Style.Font.Bold = True
                    cell.Style.Font.Name = "Arial"
                    cell.Style.Font.Height = 11
                    cell.Style.Fill.BackgroundColor = "#D9EAD3"
                except AttributeError:
                    pass

            # Escritura rápida en batches grandes
            chunk_size = 10000 if fast_mode else 1000
            row_idx = 2
            for i in range(0, len(data), chunk_size):
                chunk = data[i:i + chunk_size]
                for row_data in chunk:
                    for col_idx, field in enumerate(fields):
                        cell_address = f"{self._get_column_letter(col_idx)}{row_idx}"
                        cell = worksheet[cell_address]
                        value = row_data.get(field, "")
                        formatted_value = self._format_cell_value(value)
                        cell.Value = formatted_value
                    row_idx += 1
                # No sleep, maximiza velocidad
                if i % (chunk_size * 5) == 0:
                    logger.info(f"Lote {i//chunk_size + 1}: {row_idx-2} filas escritas...")

            # Autosize columnas SOLO al final
            for col_idx in range(len(fields)):
                try:
                    worksheet.AutoSizeColumn(col_idx)
                except AttributeError:
                    try:
                        worksheet.SetColumnWidth(col_idx, 15)
                    except AttributeError:
                        pass

            logger.info(f"Datos escritos: {len(data)} filas, {len(fields)} columnas")
        except Exception as e:
            logger.error(f"Error escribiendo datos con IronXL: {str(e)}")
            raise ExcelGenerationError(f"Error escribiendo datos: {str(e)}")

    def _format_cell_value(self, value: Any) -> str:
        if value is None:
            return ""
        elif isinstance(value, (int, float)):
            return str(value)
        elif isinstance(value, bool):
            return "Sí" if value else "No"
        else:
            return str(value)

    def _get_column_letter(self, col_idx: int) -> str:
        result = ""
        col_idx += 1  # Excel columns are 1-based
        while col_idx > 0:
            col_idx -= 1
            result = chr(col_idx % 26 + ord('A')) + result
            col_idx //= 26
        return result