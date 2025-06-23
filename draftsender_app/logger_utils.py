import os
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from draftsender_app.ui_utils import get_data_path

_logger_global = None

def configurar_logger(usuario: str = "usuario") -> logging.Logger:
    """
    Configura el sistema de logging de la aplicación.

    Crea un logger con formato personalizado que guarda logs en un archivo `.log`
    dentro de la carpeta `data/logs`, nombrado según el usuario o timestamp.

    Args:
        usuario (str): Nombre del usuario que genera los logs (usado en el nombre del archivo y dentro del log).

    Returns:
        logging.Logger: Objeto logger configurado para usar en toda la aplicación.
    """
    global _logger_global
    if _logger_global:
        return _logger_global

    usuario_normalizado = usuario.strip().replace(" ", "_")
    data_dir = get_data_path()
    log_dir = os.path.join(data_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)

    fecha = datetime.now().strftime("%Y%m%d")
    log_filename = f"{usuario_normalizado}_{fecha}.log"
    log_path = os.path.join(log_dir, log_filename)

    logger = logging.getLogger("DraftSender")
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        formatter = logging.Formatter(
            fmt=f"%(asctime)s - {usuario_normalizado} - %(module)s - %(levelname)s - [L%(lineno)d] - %(message)s",
            datefmt="%Y-%m-%d %H:%M"
        )

        file_handler = RotatingFileHandler(log_path, maxBytes=5*1024*1024, backupCount=3, encoding="utf-8")
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.DEBUG)

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        stream_handler.setLevel(logging.INFO)

        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)

    _logger_global = logger
    return logger