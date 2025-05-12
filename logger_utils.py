import os  # Para interactuar con el sistema de archivos y crear directorios
import logging  # Para trabajar con el sistema de logging en Python

# Crear el directorio 'logs' si no existe
os.makedirs("logs", exist_ok=True)
_log_file_path = "logs/app.log"  # Ruta del archivo de log

# Evitar duplicar handlers si ya se configuró
if not logging.getLogger().hasHandlers():
    # Formato común para los logs
    log_format = logging.Formatter(
        fmt="%(asctime)s - %(name)s - %(levelname)s - %(funcName)s [Línea %(lineno)d] - %(message)s",
        datefmt="%Y-%m-%d %H:%M"
    )

    # Handler para archivo de log
    file_handler = logging.FileHandler(_log_file_path, mode="a")  # Crea un archivo de log en modo append
    file_handler.setFormatter(log_format)  # Aplica el formato al archivo de log

    # Handler para la consola
    console_handler = logging.StreamHandler()  # Crea un handler para mostrar logs en consola
    console_handler.setFormatter(log_format)  # Aplica el formato a la salida en consola

    # Configurar logger raíz
    root_logger = logging.getLogger()  # Obtiene el logger raíz
    root_logger.setLevel(logging.INFO)  # Establece el nivel de log (INFO en este caso)
    root_logger.addHandler(file_handler)  # Añade el handler del archivo al logger
    root_logger.addHandler(console_handler)  # Añade el handler de la consola al logger

def configurar_logger(nombre_modulo: str) -> logging.Logger:
    """
    Configura un logger para el módulo especificado.

    Args:
        nombre_modulo (str): El nombre del módulo para el cual se configura el logger.

    Returns:
        logging.Logger: El logger configurado para el módulo.
    """
    return logging.getLogger(nombre_modulo)  # Devuelve el logger configurado para el módulo