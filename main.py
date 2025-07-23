import os
from dotenv import load_dotenv

load_dotenv()  # Ruta absoluta recomendada

from fastapi import FastAPI
from excel_generation.controllers.excel_controller import router as excel_router
from excel_generation.utils.logger import setup_logger

from excel_generation.config.ironxl_config import configurar_ironxl
configurar_ironxl()


app = FastAPI(
    title="Excel Generation API",
    description="API para generar archivos Excel desde PostgreSQL",
    version="1.0.0"
)

# Configurar logging
setup_logger()

# Registrar rutas
app.include_router(excel_router, prefix="/api/v1")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)