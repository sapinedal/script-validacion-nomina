import os
import json
import logging

class ConfigLoader:
    """
    Clase responsable de cargar y validar la configuración del sistema desde config/settings.json.
    Proporciona valores por defecto seguros en caso de fallos y maneja la configuración de la BD.
    """
    DEFAULT_SETTINGS_PATH = os.path.join("config", "settings.json")
    
    DEFAULT_CONFIG = {
        "csv_filename": "plano_nomina(2026-05) (1).csv",
        "csv_eventuales_filename": "",
        "excel_filename": "FO-GH-005 REPORTE DE NOVEDADES NÓMINA CARLO mayo 2026.xlsx",
        "output_filename": "diferencias_nomina_mayo_2026.xlsx",
        "novelty_mapping": {
            "HORAS EXTRAS DIURNAS": "0",
            "HORAS EXTRAS NOCTURNAS": "1",
            "HORAS EXTRAS DIURNAS FESTIVAS": "4",
            "HORAS EXTRAS NOCTURNAS FESTIVAS": "5",
            "RECARGO NOCTURNO": "6",
            "RECARGO FESTIVO": "7",
            "RECARGO FESTIVO NOCTURNO": "8",
            "RENUNCIA": "10",
            "TERMINACION DE CONTRATO": "11",
            "INCAPACIDAD": "12",
            "AUSENCIA NO JUSTIFICADA": "13",
            "LICENCIA MATER": "14",
            "LICENCIA PATERNIDAD": "14",
            "PERMISO NO REM,": "15",
            "PERMISO REMUNER": "15",
            "VACACIONES DISFRUTADAS": "16",
            "VACACIONES EN DINERO": "16",
            "CITA MÉDICA": "17",
            "CALAMIDAD": "18",
            "CUMPLEAÑOS": "19",
            "REGALO DE BODAS": "20",
            "ANIVERSARIO": "21",
            "DIA BRIGADISTA": "22",
            "REGALO DE GRADO HIJOS": "23",
            "REGALO DE GRADO": "24",
            "RODAMIENTO CARRO": "25",
            "RODAMIENTO MOTO": "25",
            "DISPONIBILIDAD": "26"
        },
        "comparison_tolerance": 0.01,
        "log_file": os.path.join("logs", "conciliacion_nomina.log"),
        "log_level": "INFO",
        "database": {
            "enabled": False,
            "host": "localhost",
            "port": 5432,
            "database_name": "",
            "username": "",
            "password": "",
            "filter_by_jefe_id": None
        }
    }

    def __init__(self, config_path=None):
        self.config_path = config_path or self.DEFAULT_SETTINGS_PATH
        self.config = self.load_config()

    def load_config(self):
        """
        Intenta leer el archivo JSON de configuración. Si falla, genera valores por defecto.
        """
        if not os.path.exists(self.config_path):
            logging.warning(
                f"Archivo de configuración no encontrado en '{self.config_path}'. Usando valores predeterminados."
            )
            config_dir = os.path.dirname(self.config_path)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir, exist_ok=True)
            try:
                with open(self.config_path, 'w', encoding='utf-8') as f:
                    json.dump(self.DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
            except Exception as e:
                logging.error(f"No se pudo guardar la configuración por defecto en {self.config_path}: {e}")
            return self.DEFAULT_CONFIG

        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                
            # Validar y rellenar claves faltantes con valores por defecto
            merged_config = {}
            for key, val in self.DEFAULT_CONFIG.items():
                if key in loaded:
                    if isinstance(val, dict) and isinstance(loaded[key], dict):
                        merged_config[key] = {**val, **loaded[key]}
                    else:
                        merged_config[key] = loaded[key]
                else:
                    merged_config[key] = val
                    
            return merged_config
        except Exception as e:
            logging.error(f"Error al leer '{self.config_path}': {e}. Usando configuración de fallback.")
            return self.DEFAULT_CONFIG

    @property
    def csv_filename(self) -> str:
        return self.config.get("csv_filename", self.DEFAULT_CONFIG["csv_filename"])

    @property
    def csv_eventuales_filename(self) -> str:
        return self.config.get("csv_eventuales_filename", "")

    @property
    def excel_filename(self) -> str:
        return self.config.get("excel_filename", self.DEFAULT_CONFIG["excel_filename"])

    @property
    def output_filename(self) -> str:
        return self.config.get("output_filename", self.DEFAULT_CONFIG["output_filename"])

    @property
    def novelty_mapping(self) -> dict:
        return self.config.get("novelty_mapping", self.DEFAULT_CONFIG["novelty_mapping"])

    @property
    def comparison_tolerance(self) -> float:
        return float(self.config.get("comparison_tolerance", self.DEFAULT_CONFIG["comparison_tolerance"]))

    @property
    def log_file(self) -> str:
        return self.config.get("log_file", self.DEFAULT_CONFIG["log_file"])

    @property
    def log_level(self) -> str:
        return self.config.get("log_level", self.DEFAULT_CONFIG["log_level"])

    # --- PROPIEDADES DE BASE DE DATOS ---
    
    @property
    def db_enabled(self) -> bool:
        return bool(self.config.get("database", {}).get("enabled", False))

    @property
    def db_host(self) -> str:
        return str(self.config.get("database", {}).get("host", "localhost"))

    @property
    def db_port(self) -> int:
        return int(self.config.get("database", {}).get("port", 5432))

    @property
    def db_name(self) -> str:
        return str(self.config.get("database", {}).get("database_name", ""))

    @property
    def db_username(self) -> str:
        return str(self.config.get("database", {}).get("username", ""))

    @property
    def db_password(self) -> str:
        return str(self.config.get("database", {}).get("password", ""))

    @property
    def db_filter_by_jefe_id(self):
        val = self.config.get("database", {}).get("filter_by_jefe_id", None)
        return int(val) if val is not None else None
