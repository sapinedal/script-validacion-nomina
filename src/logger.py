import os
import logging
from logging.handlers import RotatingFileHandler

def setup_logger(log_file: str, log_level_str: str = "INFO") -> logging.Logger:
    """
    Configura y retorna el logger del sistema.
    Escribe tanto en consola como en un archivo de logs rotativo para producción.
    """
    # Normalizar el nivel de log
    level_map = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR,
        "CRITICAL": logging.CRITICAL
    }
    level = level_map.get(log_level_str.upper(), logging.INFO)

    # Crear directorio del archivo de logs si no existe
    log_dir = os.path.dirname(log_file)
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)

    logger = logging.getLogger("ConciliacionPayroll")
    logger.setLevel(level)

    # Evitar duplicar handlers en importaciones repetidas
    if not logger.handlers:
        # Formato profesional para consola y archivo
        formatter = logging.Formatter(
            '[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d]: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        # 1. Console Handler (Salida en vivo)
        console_handler = logging.StreamHandler()
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        # 2. File Handler (Historial persistente, rotativo)
        try:
            # 5MB por archivo, máximo 3 respaldos
            file_handler = RotatingFileHandler(
                log_file, 
                maxBytes=5*1024*1024, 
                backupCount=3, 
                encoding='utf-8'
            )
            file_handler.setLevel(level)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except Exception as e:
            # Si por permisos no se puede crear el log, imprimir un aviso en consola
            print(f"Advertencia: No se pudo iniciar el archivo de log rotativo '{log_file}': {e}")

    return logger
