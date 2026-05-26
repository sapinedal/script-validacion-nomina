import pandas as pd
import numpy as np
import logging

class PayrollComparator:
    """
    Motor de comparación matemática de novedades de nómina.
    Agrupa registros duplicados, realiza cruces outer-join y clasifica
    los resultados en coincidencias, diferencias, faltantes y sobrantes.
    También computa métricas de resumen para el dashboard.
    """
    def __init__(self, df_csv: pd.DataFrame, df_excel: pd.DataFrame, 
                 novelty_mapping: dict, employee_db: dict, 
                 tolerance: float = 0.01, logger: logging.Logger = None):
        self.df_csv = df_csv.copy()
        self.df_excel = df_excel.copy()
        self.novelty_mapping = novelty_mapping
        self.employee_db = employee_db
        self.tolerance = tolerance
        self.logger = logger or logging.getLogger(__name__)

        # Invertir el mapeo de novedades para fácil visualización (Código -> Nombre)
        self.code_to_name = {str(v): k for k, v in novelty_mapping.items()}

    def get_novelty_name(self, code: str) -> str:
        """
        Retorna el nombre descriptivo de la novedad a partir de su código numérico.
        """
        return self.code_to_name.get(str(code), f"NOVEDAD OTROS (CÓDIGO {code})")

    def run_comparison(self) -> dict:
        """
        Ejecuta la comparación completa de ambos datasets.
        Retorna un diccionario con DataFrames clasificados y métricas globales.
        """
        self.logger.info("Iniciando comparación y auditoría de datos...")

        # 1. Identificar y aislar registros duplicados en los archivos originales
        excel_dupe_mask = self.df_excel.duplicated(subset=['IDENTIFICACION', 'CODIGO_NOVEDAD', 'FECHA'], keep=False)
        df_excel_dupes = self.df_excel[excel_dupe_mask].copy()
        if len(df_excel_dupes) > 0:
            df_excel_dupes['ORIGEN'] = 'EXCEL REPORTE'
            df_excel_dupes['DETALLE_DUPLICADO'] = df_excel_dupes.apply(
                lambda r: f"Fila Excel {int(r['ORIGINAL_ROW'])}: ID {r['IDENTIFICACION']} tiene otra entrada el {r['FECHA']}", axis=1
            )
            self.logger.warning(f"Se detectaron {len(df_excel_dupes)} filas duplicadas en el reporte Excel.")
        else:
            df_excel_dupes = pd.DataFrame(columns=['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'VALOR', 'FECHA', 'ORIGINAL_ROW', 'ORIGEN', 'DETALLE_DUPLICADO'])

        csv_dupe_mask = self.df_csv.duplicated(subset=['IDENTIFICACION', 'CODIGO_NOVEDAD', 'FECHA'], keep=False)
        df_csv_dupes = self.df_csv[csv_dupe_mask].copy()
        if len(df_csv_dupes) > 0:
            df_csv_dupes['ORIGEN'] = 'CSV PLANO'
            df_csv_dupes['DETALLE_DUPLICADO'] = "Entrada redundante en archivo plano CSV."
            self.logger.warning(f"Se detectaron {len(df_csv_dupes)} filas duplicadas en el archivo CSV de nómina.")
        else:
            df_csv_dupes = pd.DataFrame(columns=['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'VALOR', 'FECHA', 'ORIGEN', 'DETALLE_DUPLICADO'])

        # Concatenar todos los duplicados para el reporte
        cols_dupes_report = ['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'FECHA', 'VALOR', 'ORIGEN', 'DETALLE_DUPLICADO']
        if not df_excel_dupes.empty or not df_csv_dupes.empty:
            df_all_dupes = pd.concat([df_excel_dupes, df_csv_dupes], ignore_index=True)[cols_dupes_report]
            # Mapear nombre de novedad
            df_all_dupes['TIPO_NOVEDAD'] = df_all_dupes['CODIGO_NOVEDAD'].apply(self.get_novelty_name)
            # Reorganizar columnas
            df_all_dupes = df_all_dupes[['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'TIPO_NOVEDAD', 'FECHA', 'VALOR', 'ORIGEN', 'DETALLE_DUPLICADO']]
        else:
            df_all_dupes = pd.DataFrame(columns=['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'TIPO_NOVEDAD', 'FECHA', 'VALOR', 'ORIGEN', 'DETALLE_DUPLICADO'])

        # 2. Agrupar duplicados antes de la comparación (suma de horas)
        self.logger.info("Agrupando registros por Identificación, Novedad y Fecha...")
        df_excel_grouped = self.df_excel.groupby(['IDENTIFICACION', 'CODIGO_NOVEDAD', 'FECHA'], as_index=False)['VALOR'].sum()
        df_excel_grouped.rename(columns={'VALOR': 'VALOR_EXCEL'}, inplace=True)

        df_csv_grouped = self.df_csv.groupby(['IDENTIFICACION', 'CODIGO_NOVEDAD', 'FECHA'], as_index=False)['VALOR'].sum()
        df_csv_grouped.rename(columns={'VALOR': 'VALOR_CSV'}, inplace=True)

        # 3. Realizar Outer Join
        df_merged = pd.merge(
            df_excel_grouped, 
            df_csv_grouped, 
            on=['IDENTIFICACION', 'CODIGO_NOVEDAD', 'FECHA'], 
            how='outer'
        )

        # Rellenar valores nulos con 0.0 para poder realizar cálculos
        df_merged['VALOR_EXCEL_FILL'] = df_merged['VALOR_EXCEL'].fillna(0.0)
        df_merged['VALOR_CSV_FILL'] = df_merged['VALOR_CSV'].fillna(0.0)
        df_merged['DIFERENCIA'] = df_merged['VALOR_EXCEL_FILL'] - df_merged['VALOR_CSV_FILL']
        df_merged['DIFERENCIA_ABS'] = df_merged['DIFERENCIA'].abs()

        # Enriquecer nombres de empleados
        df_merged['NOMBRE'] = df_merged['IDENTIFICACION'].apply(
            lambda x: self.employee_db.get(x, "EMPLEADO NO ENCONTRADO EN EXCEL TH")
        )
        
        # Mapear nombres de novedades
        df_merged['TIPO_NOVEDAD'] = df_merged['CODIGO_NOVEDAD'].apply(self.get_novelty_name)

        # 4. Clasificación Estricta de Categorías
        
        # A. COINCIDENCIAS: Presentes en ambos, diferencia menor o igual a tolerancia
        matches_mask = (df_merged['VALOR_EXCEL'].notnull()) & \
                       (df_merged['VALOR_CSV'].notnull()) & \
                       (df_merged['DIFERENCIA_ABS'] <= self.tolerance)
        df_matches = df_merged[matches_mask].copy()

        # B. DIFERENCIAS EN VALOR: Presentes en ambos, diferencia mayor a tolerancia
        diff_mask = (df_merged['VALOR_EXCEL'].notnull()) & \
                    (df_merged['VALOR_CSV'].notnull()) & \
                    (df_merged['DIFERENCIA_ABS'] > self.tolerance)
        df_diffs = df_merged[diff_mask].copy()

        # C. FALTANTES EN PLANO (Excel operativo tiene horas, pero CSV no tiene nada)
        missing_mask = (df_merged['VALOR_EXCEL'].notnull() & (df_merged['VALOR_EXCEL'] > 0.0)) & \
                       (df_merged['VALOR_CSV'].isnull() | (df_merged['VALOR_CSV_FILL'] == 0.0))
        df_missing = df_merged[missing_mask].copy()

        # D. SOBRANTES EN PLANO (CSV tiene horas, pero Excel operativo no tiene nada)
        extra_mask = (df_merged['VALOR_CSV'].notnull() & (df_merged['VALOR_CSV'] > 0.0)) & \
                     (df_merged['VALOR_EXCEL'].isnull() | (df_merged['VALOR_EXCEL_FILL'] == 0.0))
        df_extra = df_merged[extra_mask].copy()

        # Limpiar columnas temporales de cálculo en cada DataFrame final
        cols_final_det = ['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'TIPO_NOVEDAD', 'FECHA', 'VALOR_EXCEL', 'VALOR_CSV', 'DIFERENCIA']
        
        df_matches_clean = df_matches[cols_final_det].copy()
        df_diffs_clean = df_diffs[cols_final_det].copy()
        df_missing_clean = df_missing[['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'TIPO_NOVEDAD', 'FECHA', 'VALOR_EXCEL']].copy()
        df_extra_clean = df_extra[['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'TIPO_NOVEDAD', 'FECHA', 'VALOR_CSV']].copy()

        # 5. Generación de Métricas Globales para el RESUMEN
        metrics = {
            "total_records_excel": len(self.df_excel),
            "total_records_csv": len(self.df_csv),
            "total_hours_excel": self.df_excel['VALOR'].sum(),
            "total_hours_csv": self.df_csv['VALOR'].sum(),
            
            "count_matches": len(df_matches_clean),
            "sum_hours_matches": df_matches_clean['VALOR_EXCEL'].sum(),
            
            "count_diffs": len(df_diffs_clean),
            "sum_hours_excel_diffs": df_diffs_clean['VALOR_EXCEL'].sum(),
            "sum_hours_csv_diffs": df_diffs_clean['VALOR_CSV'].sum(),
            
            "count_missing": len(df_missing_clean),
            "sum_hours_missing": df_missing_clean['VALOR_EXCEL'].sum(),
            
            "count_extra": len(df_extra_clean),
            "sum_hours_extra": df_extra_clean['VALOR_CSV'].sum(),
            
            "count_duplicates": len(df_all_dupes),
        }

        # 6. Tabla Resumida por Tipo de Novedad
        # Recopila estadísticas agregadas por tipo de novedad para dar visibilidad al usuario
        novelty_summary_list = []
        all_codes = sorted(list(set(self.df_excel['CODIGO_NOVEDAD'].unique()).union(set(self.df_csv['CODIGO_NOVEDAD'].unique()))))
        
        for code in all_codes:
            code_str = str(code)
            name = self.get_novelty_name(code_str)
            
            sub_excel = self.df_excel[self.df_excel['CODIGO_NOVEDAD'] == code_str]
            sub_csv = self.df_csv[self.df_csv['CODIGO_NOVEDAD'] == code_str]
            
            sub_matches = df_matches_clean[df_matches_clean['CODIGO_NOVEDAD'] == code_str]
            sub_diffs = df_diffs_clean[df_diffs_clean['CODIGO_NOVEDAD'] == code_str]
            sub_missing = df_missing_clean[df_missing_clean['CODIGO_NOVEDAD'] == code_str]
            sub_extra = df_extra_clean[df_extra_clean['CODIGO_NOVEDAD'] == code_str]
            
            val_excel = sub_excel['VALOR'].sum()
            val_csv = sub_csv['VALOR'].sum()
            
            novelty_summary_list.append({
                'CÓDIGO': code_str,
                'DESCRIPCIÓN': name,
                'HORAS EXCEL (REPORTE)': val_excel,
                'HORAS CSV (PLANO)': val_csv,
                'DIFERENCIA (HORAS)': val_excel - val_csv,
                'COINCIDENCIAS (CANT)': len(sub_matches),
                'DIFERENCIAS (CANT)': len(sub_diffs),
                'FALTANTES EN CSV (CANT)': len(sub_missing),
                'SOBRANTES EN CSV (CANT)': len(sub_extra)
            })
            
        df_novelty_summary = pd.DataFrame(novelty_summary_list)

        self.logger.info(f"Comparación ejecutada. Coincidencias: {len(df_matches_clean)}, Diferencias: {len(df_diffs_clean)}, Faltantes en plano: {len(df_missing_clean)}, Sobrantes en plano: {len(df_extra_clean)}")

        return {
            "coincidencias": df_matches_clean,
            "diferencias": df_diffs_clean,
            "faltantes_en_plano": df_missing_clean,
            "sobrantes_en_plano": df_extra_clean,
            "duplicados": df_all_dupes,
            "resumen_novedades": df_novelty_summary,
            "metrics": metrics
        }
