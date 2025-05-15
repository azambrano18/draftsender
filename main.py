import os
import sys
import atexit
from tkinter import Tk
from logger_utils import configurar_logger
from version import __version__
from gui import DraftSenderApp  # Importas la clase, no la función construir_gui

logger = configurar_logger("main")

def cargar_icono_ventana(root: Tk) -> None:
    """
    Carga el ícono de la ventana si está disponible.
    """
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        icon_path = os.path.join(base_path, "config", "icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
        else:
            logger.warning(f"Ícono no encontrado en: {icon_path}")
    except Exception as e:
        logger.warning(f"No se pudo cargar el ícono: {e}")

def cerrar_aplicacion() -> None:
    logger.info("Aplicación finalizada.")

def main():
    logger.info(f"Iniciando aplicación - Versión {__version__}")
    root = Tk()
    cargar_icono_ventana(root)
    DraftSenderApp(root)  # ← Usamos la clase
    atexit.register(cerrar_aplicacion)
    root.mainloop()

if __name__ == "__main__":
    main()