import os  # Para manejar rutas de archivos
import sys  # Para acceder a argumentos y entorno del sistema
import atexit  # Para registrar funciones que se ejecuten al cerrar el programa
from tkinter import Tk  # Tkinter para crear la interfaz gráfica
from gui import construir_gui  # Función que construye la GUI principal
from logger_utils import configurar_logger  # Configuración del sistema de logging
from version import __version__  # Versión actual del programa, centralizada en version.py

# Configura el logger para el módulo main
logger = configurar_logger("main")

def cargar_icono_ventana(root):
    """
    Intenta cargar y aplicar el ícono de la ventana principal desde el directorio 'config'.
    Args:
        root (Tk): La ventana principal de la interfaz gráfica.
    """
    try:
        # Obtiene la ruta base donde está el ícono, compatible con PyInstaller
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        icon_path = os.path.join(base_path, "config", "icono.ico")
        root.iconbitmap(icon_path)  # Aplica el ícono a la ventana
        logger.info("Ícono cargado correctamente")
    except Exception as e:
        logger.warning(f"No se pudo cargar el icono: {e}")

def main():
    """
    Punto de entrada de la aplicación.
    - Inicializa el logger.
    - Crea la ventana principal (Tkinter).
    - Carga el ícono.
    - Construye la interfaz gráfica llamando a `construir_gui`.
    - Ejecuta el loop principal de la GUI.
    - Registra eventos y errores importantes.
    """
    try:
        # Log de inicio con la versión del sistema
        logger.info(f"Iniciando DraftSender GUI - Versión {__version__}")

        # Crea la ventana principal
        root = Tk()
        root.title("DraftSender - Automatización de Borradores y Envíos")
        root.geometry("480x440")

        # Carga el ícono de la aplicación
        cargar_icono_ventana(root)

        # Construye toda la interfaz gráfica (botones, menús, campos, etc.)
        construir_gui(root)

        # Inicia el bucle principal de eventos (ventana visible)
        root.mainloop()

        # Log de cierre normal
        logger.info("Aplicación finalizada correctamente")
    except Exception:
        # Log de cualquier error inesperado en tiempo de ejecución
        logger.exception("Ocurrió un error crítico en main.py")

# Registra una acción al cerrar la app: dejar constancia en el log
atexit.register(lambda: logger.info("Aplicación cerrada correctamente."))

# Ejecuta la función main si este archivo se ejecuta directamente
if __name__ == "__main__":
    main()