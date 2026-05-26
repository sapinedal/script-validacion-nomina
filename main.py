import os
import sys
import traceback
import logging

from src.config_loader import ConfigLoader
from src.logger import setup_logger
from src.extractor import PayrollExtractor
from src.comparator import PayrollComparator
from src.exporter import ExcelReportExporter

def run_payroll_reconciliation():
    """
    Función orquestadora principal que ejecuta el proceso completo de conciliación
    de novedades de nómina entre el reporte operativo Excel y el archivo plano CSV.
    """
    # 1. Cargar Configuración
    config_loader = ConfigLoader()
    
    # 2. Configurar Logging
    logger = setup_logger(config_loader.log_file, config_loader.log_level)
    
    logger.info("======================================================================")
    logger.info("   INICIANDO PROCESAMIENTO AUTOMATIZADO DE CONCILIACIÓN DE NÓMINA")
    logger.info("======================================================================")
    logger.info(f"Archivo Plano (CSV): {config_loader.csv_filename}")
    logger.info(f"Reporte Operativo (Excel): {config_loader.excel_filename}")
    logger.info(f"Tolerancia de comparación matemática: {config_loader.comparison_tolerance} horas")
    
    if config_loader.db_enabled:
        logger.info(f"Conexión a Base de Datos de Producción ACTIVA en host: {config_loader.db_host}")
        if config_loader.db_filter_by_jefe_id is not None:
            logger.info(f"Filtro de subordinados activo para jefe_id = {config_loader.db_filter_by_jefe_id} (Carlo)")
    else:
        logger.info("Conexión a Base de Datos DESHABILITADA en configuración. Modo local global activo.")
        
    try:
        # 3. Inicializar Extractor con integración de Base de Datos
        db_config = {
            "enabled": config_loader.db_enabled,
            "host": config_loader.db_host,
            "port": config_loader.db_port,
            "database_name": config_loader.db_name,
            "username": config_loader.db_username,
            "password": config_loader.db_password,
            "filter_by_jefe_id": config_loader.db_filter_by_jefe_id
        }

        extractor = PayrollExtractor(
            csv_path=config_loader.csv_filename,
            excel_path=config_loader.excel_filename,
            novelty_mapping=config_loader.novelty_mapping,
            db_config=db_config,
            logger=logger,
            csv_eventuales_path=config_loader.csv_eventuales_filename
        )
        
        # 4. Construir base de datos de nombres (desde BD o Excel)
        logger.info("Construyendo mapeo maestro de empleados...")
        extractor.build_employee_database()
        
        # 5. Cargar archivos limpios (el filtrado de equipo se aplica aquí automáticamente)
        df_csv = extractor.load_csv()
        df_excel = extractor.load_excel()
        
        # 6. Inicializar Comparator y Ejecutar Cruces
        comparator = PayrollComparator(
            df_csv=df_csv,
            df_excel=df_excel,
            novelty_mapping=config_loader.novelty_mapping,
            employee_db=extractor.employee_db,
            tolerance=config_loader.comparison_tolerance,
            logger=logger
        )
        
        comp_results = comparator.run_comparison()
        
        # 7. Inicializar Exporter y Crear Reporte Excel Corporativo
        exporter = ExcelReportExporter(
            output_path=config_loader.output_filename,
            logger=logger
        )
        
        exporter.generate_report(comp_results)
        
        # 8. Mostrar Resumen Ejecutivo en Consola
        metrics = comp_results["metrics"]
        
        logger.info("======================================================================")
        logger.info("   PROCESAMIENTO FINALIZADO CON ÉXITO - RESUMEN EJECUTIVO EN CONSOLA")
        logger.info("======================================================================")
        
        # Generar etiquetas informativas según el modo
        modo_cruce = "GLOBAL - TODA LA COMPAÑÍA"
        if config_loader.db_enabled and hasattr(extractor, 'team_filter') and extractor.team_filter:
            jefe_display = getattr(extractor, 'jefe_name', f"JEFE ID {config_loader.db_filter_by_jefe_id}")
            modo_cruce = f"SECTORIAL - EQUIPO DE {jefe_display} ({len(extractor.team_filter)} Empleados)"

        print("\n\n" + "="*80)
        print("   REPORTE DE AUDITORÍA Y CONCILIACIÓN DE NÓMINA (MAYO 2026)")
        print("="*80)
        print(f" -> Modo de Conciliación:  {modo_cruce}")
        print(f" -> Archivo Reporte Excel: {config_loader.excel_filename}")
        print(f" -> Archivo Plano CSV:     {config_loader.csv_filename}")
        print(f" -> Archivo de Salida:     {config_loader.output_filename}")
        print("-"*80)
        print(f"  [+] REGISTROS CARGADOS (FILTRADOS POR EL MODO):")
        print(f"      - Excel Operativo:    {metrics['total_records_excel']} registros ({metrics['total_hours_excel']:.2f} horas)")
        print(f"      - Plano CSV Nómina:   {metrics['total_records_csv']} registros ({metrics['total_hours_csv']:.2f} horas)")
        print("-"*80)
        print(f"  [+] RESULTADO DE LA CONCILIACIÓN (Cruces agrupados por Día/Novedad):")
        print(f"      - Coincidencias Perfectas:   {metrics['count_matches']} registros ({metrics['sum_hours_matches']:.2f} horas)")
        print(f"      - Diferencias en Horas:      {metrics['count_diffs']} registros")
        print(f"      - FALTANTES EN PLANO (Excel): {metrics['count_missing']} registros ({metrics['sum_hours_missing']:.2f} horas)")
        print(f"      - SOBRANTES EN PLANO (CSV):   {metrics['count_extra']} registros ({metrics['sum_hours_extra']:.2f} horas)")
        print(f"      - Registros Duplicados:      {metrics['count_duplicates']} celdas")
        print("="*80)
        print("   Las hojas generadas son:")
        print("   - [RESUMEN]           : Dashboard corporativo y KPI de control financiero.")
        print("   - [COINCIDENCIAS]     : Novedades que concuerdan 100% en ambos archivos.")
        print("   - [DIFERENCIAS]       : Novedades con discrepancias en el conteo de horas.")
        print("   - [FALTANTES_EN_PLANO]: Novedades operativas no enviadas a nómina (Riesgo!).")
        print("   - [SOBRANTES_EN_PLANO]: Novedades en nómina no reportadas por la operación.")
        print("   - [DUPLICADOS]        : Registros repetidos en los archivos originales.")
        print("="*80 + "\n")
        
        logger.info(f"El reporte de auditoría ha sido exportado exitosamente a '{config_loader.output_filename}'.")
        logger.info("======================================================================")

    except Exception as e:
        logger.error("Se produjo un error fatal durante la conciliación de nómina:")
        logger.error(str(e))
        logger.error(traceback.format_exc())
        print("\n\n" + "!"*80)
        print("   ERROR FATAL EN LA EJECUCIÓN DEL SCRIPT")
        print("!"*80)
        print(f"Detalle: {e}")
        print("Por favor revise el archivo de log en 'logs/conciliacion_nomina.log' para más información.")
        print("!"*80 + "\n")
        sys.exit(1)

if __name__ == '__main__':
    run_payroll_reconciliation()
