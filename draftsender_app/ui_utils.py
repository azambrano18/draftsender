import logging
import os
import sys
from tkinter import messagebox

logger = logging.getLogger("DraftSender")

def mostrar_error(titulo: str, mensaje: str, exc: Exception = None):
    """
    Muestra una ventana emergente de error con un mensaje personalizado.

    Args:
        titulo (str): Título de la ventana.
        mensaje (str): Contenido del mensaje de error.
    """
    logger.error(f"{titulo}: {mensaje}", exc_info=exc)
    messagebox.showerror(titulo, mensaje)

def mostrar_advertencia(titulo: str, mensaje: str):
    """
    Muestra una advertencia al usuario en forma de ventana emergente.

    Args:
        titulo (str): Título de la advertencia.
        mensaje (str): Contenido del mensaje.
    """
    logger.warning(f"{titulo}: {mensaje}")
    messagebox.showwarning(titulo, mensaje)

def mostrar_info(titulo: str, mensaje: str):
    """
    Muestra un mensaje informativo al usuario mediante una ventana emergente.

    Args:
        titulo (str): Título del mensaje.
        mensaje (str): Contenido del mensaje.
    """
    logger.info(f"{titulo}: {mensaje}")
    messagebox.showinfo(titulo, mensaje)

def get_data_path():
    """
    Obtiene la ruta del directorio 'data' desde donde se ejecuta la aplicación.

    Si está empaquetada con PyInstaller, ajusta la ruta correctamente.

    Returns:
        str: Ruta absoluta al directorio `data`.
    """
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_path, "data")
    os.makedirs(data_dir, exist_ok=True)

    return data_dir