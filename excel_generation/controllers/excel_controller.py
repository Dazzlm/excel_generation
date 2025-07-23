from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Dict, List, Optional
import io
import logging

from ..services.generator_service import ExcelGeneratorService
from ..services.upload_service import ExcelUploadService
from ..utils.error_handling import ExcelGenerationError, ValidationError

router = APIRouter()
logger = logging.getLogger(__name__)

class ExcelRequest(BaseModel):
    table: str
    fields: List[str]
    filters: Optional[Dict] = {}
    template: Optional[str] = "default_template.xlsx"

class UploadResponse(BaseModel):
    success: bool
    message: str
    rows_processed: int
    errors: Optional[List[str]] = None

@router.post("/reports/excel")
async def generate_excel(request: ExcelRequest):
    try:
        logger.info(f"Generando Excel para tabla: {request.table}")
        service = ExcelGeneratorService()
        excel_buffer = await service.generate_excel(
            table=request.table,
            fields=request.fields,
            filters=request.filters,
            template=request.template
        )
        headers = {
            'Content-Disposition': f'attachment; filename="{request.table}_export.xlsx"',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        return StreamingResponse(
            io.BytesIO(excel_buffer),
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers=headers
        )
    except ExcelGenerationError as e:
        logger.error(f"Error de generación Excel: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error(f"Error interno: {str(e)}")
        raise HTTPException(status_code=500, detail="Error interno del servidor")

@router.get("/reports/excel/table/{table_name}")
async def download_whole_table(table_name: str, template: str = "default_template.xlsx"):
    """
    Descarga la tabla completa (todos los campos, todos los datos) en Excel.
    """
    try:
        service = ExcelGeneratorService()
        excel_buffer = await service.generate_full_table_excel(table_name, template)
        headers = {
            'Content-Disposition': f'attachment; filename="{table_name}_completa.xlsx"',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        return StreamingResponse(
            io.BytesIO(excel_buffer),
            headers=headers,
            media_type=headers['Content-Type']
        )
    except Exception as e:
        logger.error(f"Error descargando tabla completa: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

@router.get("/reports/excel/all")
async def download_entire_database(template: str = "default_template.xlsx"):
    """
    Descarga toda la base de datos, cada tabla en una hoja diferente del mismo Excel.
    """
    try:
        service = ExcelGeneratorService()
        excel_buffer = await service.generate_full_database_excel(template)
        headers = {
            'Content-Disposition': 'attachment; filename="toda_base_de_datos.xlsx"',
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        return StreamingResponse(
            io.BytesIO(excel_buffer),
            headers=headers,
            media_type=headers['Content-Type']
        )
    except Exception as e:
        logger.error(f"Error descargando toda la base de datos: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error: {str(e)}")

@router.post("/upload/excel", response_model=UploadResponse)
async def upload_excel(
    file: UploadFile = File(...),
    table: str = Form(...),
):
    """
    Carga un archivo Excel en la tabla especificada.
    Siempre usa el método más rápido (COPY + UPSERT masivo).
    """
    try:
        logger.info(f"Cargando Excel a tabla: {table}")
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise ValidationError("El archivo debe ser .xlsx o .xls")
        content = await file.read()
        upload_service = ExcelUploadService()
        # Siempre usa el modo más rápido, no batch insert normal
        result = await upload_service.upload_excel_to_table(
            excel_data=content,
            table_name=table,
            conflict_columns=["id"]
        )
        logger.info(f"Carga completada: {result['rows_processed']} filas procesadas")
        return UploadResponse(
            success=result['success'],
            message=result['message'],
            rows_processed=result['rows_processed'],
            errors=result.get('errors', [])
        )
    except ValidationError as e:
        logger.error(f"Error de validación: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error(f"Error interno en carga: {str(e)}")
        raise HTTPException(status_code=500, detail="Error interno del servidor")

@router.get("/tables/{table_name}/columns")
async def get_table_columns(table_name: str):
    try:
        upload_service = ExcelUploadService()
        columns = await upload_service.get_table_info(table_name)
        return {
            "table": table_name,
            "columns": columns
        }
    except Exception as e:
        logger.error(f"Error obteniendo columnas: {str(e)}")
        raise HTTPException(status_code=400, detail=str(e))

@router.get("/health")
async def health_check():
    return {"status": "healthy", "service": "Excel Generation API"}