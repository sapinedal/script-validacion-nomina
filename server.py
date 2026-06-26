import os
import uuid
import json
import logging
import psycopg2
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from src.config_loader import ConfigLoader
from src.extractor import PayrollExtractor
from src.comparator import PayrollComparator
from src.exporter import ExcelReportExporter

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN GLOBAL
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# Instancia de FastAPI
app = FastAPI(
    title="Validador de Nómina (Trazalo ERP)",
    description="API y Frontend para auditar archivos Excel contra PostgreSQL",
    version="1.0.0"
)

# Configurar motor de plantillas (Jinja2)
templates = Jinja2Templates(directory="templates")

# Crear carpeta temporal para los archivos subidos si no existe
TEMP_DIR = "temp_uploads"
os.makedirs(TEMP_DIR, exist_ok=True)


# ─────────────────────────────────────────────────────────────────────────────
# DEPENDENCIAS E INSTANCIAS (Singleton-like para el ciclo de vida)
# ─────────────────────────────────────────────────────────────────────────────

def get_config() -> ConfigLoader:
    # Asume que tienes un settings.json en la misma ruta o ruta por defecto
    try:
        return ConfigLoader()
    except Exception as e:
        logger.error(f"No se pudo cargar la configuración: {e}")
        raise HTTPException(status_code=500, detail="Error de configuración del servidor")

def get_db_connection(config: ConfigLoader):
    try:
        conn = psycopg2.connect(
            host=config.db_host,
            port=config.db_port,
            database=config.db_name,
            user=config.db_username,
            password=config.db_password
        )
        return conn
    except Exception as e:
        logger.error(f"Error conectando a PostgreSQL: {e}")
        raise Exception(str(e))


# ─────────────────────────────────────────────────────────────────────────────
# RUTAS FRONTEND (Vistas HTML)
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def serve_frontend(request: Request):
    """
    Renderiza la interfaz principal (index.html) donde el usuario puede arrastrar
    el archivo de novedades de Excel.
    """
    return templates.TemplateResponse(request=request, name="index.html")

@app.get("/resumen", response_class=HTMLResponse)
async def serve_resumen(request: Request):
    """
    Renderiza la vista de resumen (resumen.html)
    """
    return templates.TemplateResponse(request=request, name="resumen.html")

# ─────────────────────────────────────────────────────────────────────────────
# RUTAS API (Backend)
# ─────────────────────────────────────────────────────────────────────────────

@app.get("/api/health")
async def health_check():
    """
    Endpoint para que el frontend verifique si hay conexión a la Base de Datos.
    """
    try:
        config = get_config()
        
        # Opcional: si la BD está deshabilitada en config
        if not config.db_enabled:
            return {"status": "OK", "detail": "Base de datos deshabilitada por configuración"}

        # Intentar conectar
        conn = get_db_connection(config)
        
        # Validar periodo
        extractor = PayrollExtractor(conn, config)
        periodo_actual = extractor.periodo
        
        conn.close()
        return {"status": "OK", "periodo_actual": periodo_actual}
        
    except Exception as e:
        return {"status": "ERROR", "detail": str(e)}


@app.post("/api/validar")
async def validate_payroll(file_excel: UploadFile = File(...)):
    """
    Endpoint principal que recibe el Excel, ejecuta la extracción, comparación
    y exporta el archivo de auditoría, devolviéndolo como descarga.
    """
    
    # 1. Validar extensión del archivo
    if not (file_excel.filename.endswith(".xlsx") or file_excel.filename.endswith(".xls")):
        raise HTTPException(status_code=400, detail="El archivo debe ser un Excel (.xlsx, .xls)")

    # 2. Guardar archivo temporalmente
    file_id = str(uuid.uuid4())[:8]
    temp_file_path = os.path.join(TEMP_DIR, f"upload_{file_id}_{file_excel.filename}")
    
    try:
        with open(temp_file_path, "wb") as buffer:
            buffer.write(await file_excel.read())
            
        logger.info(f"Archivo recibido: {temp_file_path}")

        # 3. Cargar Configuración
        config = get_config()

        # 4. Leer Excel Operativo
        df_operativo = pd.read_excel(temp_file_path, sheet_name=config.excel_sheet_name)
        if df_operativo.empty:
            raise HTTPException(status_code=400, detail="El archivo Excel está vacío o no se encontró la hoja correcta.")
            
        # Opcional: Limpiar N/A o nulos en las columnas clave
        if config.excel_cedula_col in df_operativo.columns:
            df_operativo = df_operativo.dropna(subset=[config.excel_cedula_col])

        # 5. Extraer Datos del ERP (PostgreSQL)
        if not config.db_enabled:
            raise HTTPException(status_code=500, detail="La validación de BD está deshabilitada en settings.json")

        conn = get_db_connection(config)
        extractor = PayrollExtractor(conn, config)
        
        logger.info("Extrayendo novedades del ERP...")
        df_erp = extractor.extract_novedades()
        conn.close()

        if df_erp.empty:
            raise HTTPException(status_code=404, detail="No se encontraron novedades en el ERP para el periodo actual.")

        # 6. Comparar Operativo vs ERP
        logger.info("Comparando datos...")
        comparator = PayrollComparator(df_operativo, df_erp, config)
        df_resultado = comparator.compare()

        # 7. Exportar Reporte de Auditoría
        output_filename = f"Auditoria_Nomina_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_filepath = os.path.join(TEMP_DIR, output_filename)
        
        logger.info("Generando reporte de diferencias...")
        exporter = ExcelReportExporter(df_resultado, output_filepath)
        exporter.export()

        # 8. Devolver archivo como descarga
        return FileResponse(
            path=output_filepath, 
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            background=None # No borrar de inmediato para evitar errores en FileResponse, se debe limpiar asíncronamente en prod
        )

    except Exception as e:
        logger.error(f"Error procesando el archivo: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Limpieza del archivo de subida (opcional, podrías limpiar también el de salida con un background task)
        if os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except Exception as e:
                logger.warning(f"No se pudo eliminar el archivo temporal {temp_file_path}: {e}")
