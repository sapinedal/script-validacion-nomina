import os
import re
import pandas as pd
import numpy as np
import openpyxl
import logging

class PayrollExtractor:
    """
    Clase encargada de leer, limpiar y normalizar los datos de nómina,
    y realizar cruces y filtrados utilizando la base de datos PostgreSQL de producción.
    """
    def __init__(self, csv_path: str, excel_path: str, novelty_mapping: dict, 
                 db_config: dict = None, logger: logging.Logger = None, csv_eventuales_path: str = ""):
        self.csv_path = csv_path
        self.csv_eventuales_path = csv_eventuales_path
        self.excel_path = excel_path
        self.novelty_mapping = novelty_mapping
        self.db_config = db_config or {"enabled": False}
        self.logger = logger or logging.getLogger(__name__)
        
        self.employee_db = {}     # Base de datos ID -> Nombre construida dinámicamente
        self.team_filter = None   # Conjunto de IDs filtrados si se activa filtro por jefe

    def clean_id(self, raw_id) -> str:
        """
        Limpia un ID de empleado. Remueve espacios, guiones y convierte floats (12345.0) a string limpio (12345).
        """
        if pd.isnull(raw_id):
            return ""
        
        val_str = str(raw_id).strip()
        if val_str.endswith(".0"):
            val_str = val_str[:-2]
            
        # Remover cualquier caracter no alfanumérico
        val_str = re.sub(r"\s+", "", val_str)
        return val_str

    def parse_date(self, raw_date) -> str:
        """
        Parsea una fecha de tipo mixto y la retorna en formato estándar DD/MM/YYYY.
        """
        if pd.isnull(raw_date):
            return ""
            
        if hasattr(raw_date, "strftime"):
            return raw_date.strftime("%d/%m/%Y")
            
        val_str = str(raw_date).strip()
        if not val_str or val_str.lower() == "nan":
            return ""
            
        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y"]:
            try:
                dt = pd.to_datetime(val_str, format=fmt, errors='raise')
                return dt.strftime("%d/%m/%Y")
            except:
                continue
                
        try:
            dt = pd.to_datetime(val_str, dayfirst=True)
            return dt.strftime("%d/%m/%Y")
        except Exception as e:
            self.logger.warning(f"No se pudo normalizar la fecha '{raw_date}': {e}")
            return val_str

    def parse_value(self, raw_val) -> float:
        """
        Parsea un valor decimal que puede venir con coma o punto, o ser un tipo numérico directo.
        """
        if pd.isnull(raw_val):
            return 0.0
            
        if isinstance(raw_val, (int, float)):
            return float(raw_val)
            
        val_str = str(raw_val).strip()
        if not val_str or val_str.lower() == "nan":
            return 0.0
            
        try:
            val_str = val_str.replace(",", ".")
            return float(val_str)
        except Exception as e:
            self.logger.warning(f"No se pudo convertir el valor '{raw_val}' a float: {e}. Se asume 0.0")
            return 0.0

    def build_employee_database(self):
        """
        Construye la base de datos de nombres. Prioriza la conexión PostgreSQL
        para extraer los nombres reales en producción. Si falla o está deshabilitada,
        usa el algoritmo de escaneo de hojas de Excel locales.
        """
        # A. INTENTAR CONSULTAR BASE DE DATOS POSTGRESQL DE PRODUCCIÓN
        if self.db_config.get("enabled", False):
            self.logger.info("Conectando a base de datos PostgreSQL de producción para descargar maestros...")
            try:
                import psycopg2
                conn = psycopg2.connect(
                    host=self.db_config.get("host", "localhost"),
                    database=self.db_config.get("database_name", ""),
                    user=self.db_config.get("username", ""),
                    password=self.db_config.get("password", ""),
                    port=self.db_config.get("port", 5432)
                )
                cur = conn.cursor()
                
                # 1. Cargar base de datos maestra de nombres
                cur.execute("SELECT documento, name FROM users WHERE documento IS NOT NULL;")
                db_users = cur.fetchall()
                for row in db_users:
                    clean_doc = self.clean_id(row[0])
                    if clean_doc:
                        self.employee_db[clean_doc] = str(row[1]).strip()
                        
                self.logger.info(f"Base de datos de empleados cargada desde Producción: {len(self.employee_db)} empleados.")
                
                # 2. Si se requiere filtrado por Jefe de Área
                jefe_id = self.db_config.get("filter_by_jefe_id")
                if jefe_id is not None:
                    cur.execute("SELECT name FROM users WHERE id = %s;", (jefe_id,))
                    jefe_row = cur.fetchone()
                    self.jefe_name = str(jefe_row[0]).strip().upper() if jefe_row else f"JEFE ID {jefe_id}"
                    
                    self.logger.info(f"Filtrando novedades por subordinados de jefe_id = {jefe_id} ({self.jefe_name})...")
                    cur.execute("SELECT documento FROM users WHERE jefe_id = %s AND documento IS NOT NULL;", (jefe_id,))
                    team_rows = cur.fetchall()
                    self.team_filter = {self.clean_id(r[0]) for r in team_rows}
                    self.logger.info(f"Filtro cargado desde la BD: se auditarán únicamente {len(self.team_filter)} empleados asignados a {self.jefe_name}.")
                    
                cur.close()
                conn.close()
                return  # Éxito absoluto, saltar carga por hojas
            except Exception as e:
                self.logger.error(f"Fallo la carga desde la base de datos de producción: {e}. Se procederá con fallback local.")

        # B. FALLBACK LOCAL (ESCANEAR EXCEL LOCAL)
        self.logger.info("Construyendo base de datos de empleados localmente desde el Excel...")
        if not os.path.exists(self.excel_path):
            self.logger.error(f"Archivo Excel no encontrado para base de datos de empleados: {self.excel_path}")
            return

        try:
            xls = pd.ExcelFile(self.excel_path)
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name, header=None)
                num_rows = min(15, len(df))
                id_col_idx = None
                name_col_idx = None
                
                for r_idx in range(num_rows):
                    row_vals = [str(x).strip().upper() for x in df.iloc[r_idx].tolist()]
                    for c_idx, val in enumerate(row_vals):
                        if "IDENTIFICACI" in val or "CEDULA" in val or "DOCUMENTO" in val:
                            id_col_idx = c_idx
                        if "NOMBRE" in val or "APELLIDO" in val:
                            name_col_idx = c_idx
                            
                    if id_col_idx is not None and name_col_idx is not None:
                        for data_r_idx in range(r_idx + 1, len(df)):
                            raw_id = df.iloc[data_r_idx, id_col_idx]
                            raw_name = df.iloc[data_r_idx, name_col_idx]
                            
                            clean_id = self.clean_id(raw_id)
                            if clean_id and pd.notnull(raw_name):
                                clean_name = str(raw_name).strip()
                                if clean_name and clean_name.lower() != "nan":
                                    self.employee_db[clean_id] = clean_name
                        break
            self.logger.info(f"Base de datos de empleados local construida con {len(self.employee_db)} registros únicos.")
        except Exception as e:
            self.logger.error(f"Error al construir base de datos de empleados local: {e}", exc_info=True)

    def load_csv(self) -> pd.DataFrame:
        """
        Carga y normaliza el archivo plano CSV de nómina global.
        Si hay un filtro de equipo cargado de la BD, filtra el set de datos para mayor precisión.
        """
        if not os.path.exists(self.csv_path):
            raise FileNotFoundError(f"Archivo plano CSV no encontrado en: {self.csv_path}")

        self.logger.info(f"Cargando archivo plano CSV: {self.csv_path}")
        try:
            # Escanear el archivo Excel (si existe) para ver si las horas extras tienen fechas específicas reportadas
            self.has_overtime_dates = False
            try:
                if os.path.exists(self.excel_path):
                    sheet_name, header_row = self.detect_novelty_sheet()
                    if sheet_name:
                        xl_df = pd.read_excel(self.excel_path, sheet_name=sheet_name, skiprows=header_row)
                        
                        # Buscar columnas de fechas
                        start_date_col = None
                        for col in xl_df.columns:
                            col_upper = str(col).upper().strip()
                            if "FECHA" in col_upper or "FEHCA" in col_upper:
                                if "INICIO" in col_upper or "DESDE" in col_upper or col_upper in ["FECHA", "FEHCA"]:
                                    start_date_col = col
                                    break
                        if not start_date_col:
                            for col in xl_df.columns:
                                col_upper = str(col).upper().strip()
                                if "FECHA" in col_upper or "FEHCA" in col_upper:
                                    start_date_col = col
                                    break
                        
                        mapeo_upper = {k.upper(): v for k, v in self.novelty_mapping.items()}
                        overtime_cols = [col for col in xl_df.columns if col.upper().strip() in mapeo_upper and mapeo_upper[col.upper().strip()] in ["0", "1", "4", "5", "6", "7", "8"]]
                        
                        if overtime_cols and start_date_col:
                            for col in overtime_cols:
                                non_empty = xl_df[xl_df[col].apply(self.parse_value) > 0.0]
                                if not non_empty.empty:
                                    # Si alguna fila con horas extras tiene fecha no nula ni vacía
                                    valid_dates = non_empty[start_date_col].dropna().apply(lambda x: str(x).strip())
                                    valid_dates = valid_dates[valid_dates != ""]
                                    if not valid_dates.empty:
                                        self.has_overtime_dates = True
                                        break
            except Exception as e:
                self.logger.warning(f"No se pudo escanear el Excel para detectar fechas de horas extras: {e}")
                
            self.logger.info(f"¿Planilla Excel contiene fechas específicas para horas extras/recargos?: {self.has_overtime_dates}")

            # Cargar archivo de novedades eventuales si está configurado y existe
            df_eventuales = None
            if self.csv_eventuales_path and os.path.exists(self.csv_eventuales_path):
                self.logger.info(f"Cargando archivo de novedades eventuales plano: {self.csv_eventuales_path}")
                try:
                    df_ev = pd.read_csv(self.csv_eventuales_path, sep=';', header=None, dtype=str)
                    if df_ev.shape[1] >= 4:
                        df_ev = df_ev.iloc[:, :4]
                        df_ev.columns = ['IDENTIFICACION', 'CODIGO_NOVEDAD', 'VALOR_RAW', 'FECHA_RAW']
                        df_eventuales = df_ev
                        self.logger.info(f"Archivo de novedades eventuales cargado con {len(df_ev)} registros.")
                except Exception as e:
                    self.logger.warning(f"No se pudo cargar el archivo de novedades eventuales plano: {e}")

            df = pd.read_csv(self.csv_path, sep=';', header=None, dtype=str)
            if df.shape[1] < 4:
                raise ValueError(f"El archivo CSV debe tener al menos 4 columnas. Encontradas: {df.shape[1]}")
                
            df = df.iloc[:, :4]
            df.columns = ['IDENTIFICACION', 'CODIGO_NOVEDAD', 'VALOR_RAW', 'FECHA_RAW']
            
            # Concatenar con novedades eventuales si se cargaron
            if df_eventuales is not None:
                df = pd.concat([df, df_eventuales], ignore_index=True)
                self.logger.info(f"Total de registros CSV consolidados (Horas Extras + Eventuales): {len(df)}")
            
            df['IDENTIFICACION'] = df['IDENTIFICACION'].apply(self.clean_id)
            
            # --- FILTRADO POR EQUIPO ---
            if self.team_filter is not None:
                orig_count = len(df)
                df = df[df['IDENTIFICACION'].isin(self.team_filter)].copy()
                jefe_display = getattr(self, 'jefe_name', "Jefe de Área")
                self.logger.info(f"Filtro aplicado al CSV: Reducido de {orig_count} a {len(df)} registros correspondientes al equipo de {jefe_display}.")
            else:
                self.logger.info("Comparación global activa: no se realiza filtrado de equipo.")
            
            df['CODIGO_NOVEDAD'] = df['CODIGO_NOVEDAD'].apply(lambda x: str(x).strip() if pd.notnull(x) else "")
            df['VALOR'] = df['VALOR_RAW'].apply(self.parse_value)
            df['FECHA'] = df['FECHA_RAW'].apply(self.parse_date)
            
            # Determinar el mes principal del archivo (el más frecuente)
            self.main_month_year = "05/2026"  # Fallback por defecto
            my_counts = {}
            for vd in df['FECHA'].dropna():
                parts = vd.split('/')
                if len(parts) == 3:
                    my = f"{parts[1]}/{parts[2]}"
                    my_counts[my] = my_counts.get(my, 0) + 1
            if my_counts:
                self.main_month_year = max(my_counts, key=my_counts.get)
            self.logger.info(f"Mes de nómina principal detectado en CSV: {self.main_month_year}")
            
            # Estandarizar fechas de horas extras/recargos (códigos 0,1,4,5,6,7,8) al primer día del mes principal solo si la planilla NO tiene fechas específicas
            def adjust_csv_date(r):
                c = r['CODIGO_NOVEDAD']
                d = r['FECHA']
                if not self.has_overtime_dates:
                    if c in ["0", "1", "4", "5", "6", "7", "8"]:
                        return f"01/{self.main_month_year}"
                return d
            df['FECHA'] = df.apply(adjust_csv_date, axis=1)
            
            df = df[df['IDENTIFICACION'] != ""]
            
            # Enriquecer con nombres de base de datos
            df['NOMBRE'] = df['IDENTIFICACION'].apply(lambda x: self.employee_db.get(x, "EMPLEADO NO ENCONTRADO EN EXCEL"))
            
            df_clean = df[['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'VALOR', 'FECHA']].copy()
            self.df_csv = df_clean
            return df_clean
        except Exception as e:
            self.logger.error(f"Error cargando archivo CSV: {e}", exc_info=True)
            raise

    def detect_novelty_sheet(self) -> tuple:
        """
        Analiza las hojas del Excel y retorna automáticamente (nombre_hoja, indice_fila_cabecera).
        """
        if not os.path.exists(self.excel_path):
            raise FileNotFoundError(f"Archivo Excel no encontrado en: {self.excel_path}")

        try:
            xls = pd.ExcelFile(self.excel_path)
            ordered_sheets = ['TH'] + [s for s in xls.sheet_names if s != 'TH']
            
            for sheet_name in ordered_sheets:
                df_sample = pd.read_excel(self.excel_path, sheet_name=sheet_name, nrows=20, header=None)
                for r_idx in range(len(df_sample)):
                    row_vals = [str(x).strip().upper() for x in df_sample.iloc[r_idx].tolist()]
                    
                    has_id = any("IDENTIFICACI" in val or "CEDULA" in val for val in row_vals)
                    has_novelty_cols = any(
                        "HORAS EXTRAS DIURNAS" in val or "RECARGO NOCTURNO" in val or "HORAS EXTRAS NOCTURNAS" in val
                        for val in row_vals
                    )
                    
                    if has_id and has_novelty_cols:
                        self.logger.info(f"Hoja de novedades detectada automáticamente: '{sheet_name}' en fila {r_idx + 1} (cabecera)")
                        return sheet_name, r_idx
                        
            self.logger.warning("No se detectó un patrón claro en ninguna hoja. Usando primera hoja por defecto.")
            return xls.sheet_names[0], 0
        except Exception as e:
            self.logger.error(f"Error detectando hoja de novedades automáticamente: {e}", exc_info=True)
            raise

    def load_excel(self) -> pd.DataFrame:
        """
        Carga el reporte de Excel, normaliza datos, transforma de formato ancho
        a largo, y aplica filtrado de subordinados si está habilitada la conexión a base de datos.
        """
        sheet_name, header_row = self.detect_novelty_sheet()
        self.logger.info(f"Cargando Excel desde hoja '{sheet_name}', fila cabecera: {header_row}")
        
        try:
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name, skiprows=header_row)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # Buscar columna de ID
            id_col = None
            for col in df.columns:
                if "IDENTIFICACI" in col or "CEDULA" in col:
                    id_col = col
                    break
                    
            if id_col is None:
                raise ValueError("No se encontró la columna de IDENTIFICACIÓN en la hoja de Excel.")
                
            # Buscar columnas de fecha
            start_date_col = None
            end_date_col = None
            for col in df.columns:
                if "FEHCA INICIO" in col or "FECHA INICIO" in col:
                    start_date_col = col
                if "FEHCA FINAL" in col or "FECHA FINAL" in col:
                    end_date_col = col
                    
            if start_date_col is None:
                self.logger.warning("No se encontró FECHA INICIO en el Excel. Se buscarán alternativas.")
                for col in df.columns:
                    if "FECHA" in col or "FEHCA" in col:
                        start_date_col = col
                        break

            # Buscar columna de Nombres
            name_col = None
            for col in df.columns:
                if "NOMBRE" in col or "APELLIDO" in col:
                    name_col = col
                    break

            # Estandarizar horas extras/recargos a la fecha principal del periodo de nómina
            main_my = getattr(self, 'main_month_year', "05/2026")
            overtime_date = f"01/{main_my}"

            mapeo_upper = {k.upper(): v for k, v in self.novelty_mapping.items()}
            excel_records = []
            
            for idx, row in df.iterrows():
                raw_id = row[id_col]
                clean_id = self.clean_id(raw_id)
                if not clean_id or clean_id == "nan":
                    continue
                
                # --- FILTRADO DE SUBORDINADOS EN EXCEL ---
                if self.team_filter is not None and clean_id not in self.team_filter:
                    continue  # Saltar fila si no pertenece al equipo de Carlo
                
                emp_name = str(row[name_col]).strip() if name_col and pd.notnull(row[name_col]) else ""
                if not emp_name:
                    emp_name = self.employee_db.get(clean_id, "EMPLEADO SIN NOMBRE")
                
                raw_date = row.get(start_date_col) if start_date_col else None
                if pd.isnull(raw_date) and end_date_col:
                    raw_date = row.get(end_date_col)
                
                # --- EXTRACCIÓN INTELIGENTE DE FECHAS PARA EVITAR OMITIR REGISTROS ---
                detected_dates = []
                if not pd.isnull(raw_date):
                    clean_date = self.parse_date(raw_date)
                    if clean_date:
                        detected_dates.append(clean_date)
                
                # Si no hay fecha en columnas, buscar en OBSERVACIONES y OTRO NO ESPEC *
                if not detected_dates:
                    obs_text = str(row.get("OBSERVACIONES", "")) if "OBSERVACIONES" in row else ""
                    other_text = str(row.get("OTRO NO ESPEC *", "")) if "OTRO NO ESPEC *" in row else ""
                    combined_text = f"{obs_text} {other_text}"
                    
                    # Buscar fechas con formato DD/MM/YYYY o DD-MM-YYYY (ej. 3/05/2026, 24-05-2026)
                    matches = re.findall(r'\b\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}\b', combined_text)
                    for m in matches:
                        parsed_m = self.parse_date(m)
                        if parsed_m and parsed_m not in detected_dates:
                            detected_dates.append(parsed_m)

                for col_name, code in mapeo_upper.items():
                    if col_name in df.columns:
                        raw_val = row[col_name]
                        val_float = self.parse_value(raw_val)
                        
                        if val_float > 0.0:
                            # Si es novedad de horas extras o recargos (códigos 0,1,4,5,6,7,8), se acumula mensualmente solo si la planilla NO tiene fechas específicas
                            if str(code) in ["0", "1", "4", "5", "6", "7", "8"] and not getattr(self, 'has_overtime_dates', False):
                                excel_records.append({
                                    'IDENTIFICACION': clean_id,
                                    'NOMBRE': emp_name,
                                    'CODIGO_NOVEDAD': str(code),
                                    'VALOR': val_float,
                                    'FECHA': overtime_date,
                                    'ORIGINAL_ROW': idx + header_row + 2
                                })
                            else:
                                # Si detectamos fechas en observaciones/columnas, distribuimos o asignamos (para eventualidades)
                                if detected_dates:
                                    split_val = val_float / len(detected_dates)
                                    for d in detected_dates:
                                        excel_records.append({
                                            'IDENTIFICACION': clean_id,
                                            'NOMBRE': emp_name,
                                            'CODIGO_NOVEDAD': str(code),
                                            'VALOR': split_val,
                                            'FECHA': d,
                                            'ORIGINAL_ROW': idx + header_row + 2
                                        })
                                else:
                                    # Fallback 1: Buscar en el CSV la fecha de este tipo de novedad para este empleado
                                    clean_date = None
                                    if hasattr(self, 'df_csv') and self.df_csv is not None:
                                        matching_csv = self.df_csv[
                                            (self.df_csv['IDENTIFICACION'] == clean_id) & 
                                            (self.df_csv['CODIGO_NOVEDAD'] == str(code))
                                        ]
                                        if not matching_csv.empty:
                                            csv_dates = matching_csv['FECHA'].dropna().unique()
                                            if len(csv_dates) > 0:
                                                clean_date = csv_dates[0]
                                                self.logger.info(
                                                    f"Fila {idx + header_row + 2}: Sin fecha en Excel. "
                                                    f"Asociada a la fecha {clean_date} del CSV para ID {clean_id} y novedad {code}."
                                                )
                                        
                                        # Fallback 2: Buscar cualquier fecha de este empleado en el CSV
                                        if clean_date is None:
                                            matching_emp = self.df_csv[self.df_csv['IDENTIFICACION'] == clean_id]
                                            if not matching_emp.empty:
                                                emp_dates = matching_emp['FECHA'].dropna().unique()
                                                if len(emp_dates) > 0:
                                                    clean_date = emp_dates[0]
                                                    self.logger.info(
                                                        f"Fila {idx + header_row + 2}: Sin fecha en Excel ni coincidencia de novedad. "
                                                        f"Asociada a la fecha general {clean_date} del empleado en el CSV."
                                                    )
                                    
                                    # Fallback 3: Usar la fecha del primer día del mes
                                    if clean_date is None:
                                        clean_date = "01/05/2026"
                                        self.logger.warning(
                                            f"Fila {idx + header_row + 2}: No se pudo determinar fecha. "
                                            f"Usando fecha por defecto {clean_date} para ID {clean_id}."
                                        )
                                    
                                    excel_records.append({
                                        'IDENTIFICACION': clean_id,
                                        'NOMBRE': emp_name,
                                        'CODIGO_NOVEDAD': str(code),
                                        'VALOR': val_float,
                                        'FECHA': clean_date,
                                        'ORIGINAL_ROW': idx + header_row + 2
                                    })
                            
            df_long = pd.DataFrame(excel_records)
            if df_long.empty:
                self.logger.warning("No se extrajeron registros del Excel para este filtro de jefe.")
                return pd.DataFrame(columns=['IDENTIFICACION', 'NOMBRE', 'CODIGO_NOVEDAD', 'VALOR', 'FECHA', 'ORIGINAL_ROW'])
                
            self.logger.info(f"Excel procesado y filtrado correctamente. Registros válidos cargados: {len(df_long)}")
            return df_long
        except Exception as e:
            self.logger.error(f"Error procesando reporte Excel: {e}", exc_info=True)
            raise
