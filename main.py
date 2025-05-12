import os
import sys
from tkinter import Tk  # Importación de Tkinter para la interfaz gráfica
from gui import construir_gui  # Función para construir la interfaz gráfica
from logger_utils import configurar_logger  # Función para configurar el logger

# Configuración del logger con el nombre "main"
logger = configurar_logger("main")


def main():
    """
    Función principal que arranca la aplicación gráfica (GUI), configura la ventana,
    carga el ícono y gestiona la ejecución del programa.

    1. Configura el logger para registrar eventos importantes.
    2. Inicializa la ventana principal de la interfaz gráfica de usuario.
    3. Carga el ícono de la ventana desde la ruta configurada.
    4. Llama a la función `construir_gui()` para construir la interfaz con Tkinter.
    5. Maneja excepciones y errores críticos durante la ejecución.
    """
    try:
        # Inicia el log para indicar que la aplicación ha comenzado
        logger.info("Iniciando DraftSender GUI")

        # Inicializa la ventana principal de la GUI
        root = Tk()
        root.title("DraftSender - Automatización de Borradores y Envíos")  # Título de la ventana
        root.geometry("480x450")  # Tamaño de la ventana

        # Establecer ícono de la ventana
        try:
            # Obtiene la ruta base y define la ruta del ícono
            base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
            icon_path = os.path.join(base_path, "config", "icono.ico")
            root.iconbitmap(icon_path)  # Establece el ícono de la ventana
            logger.info("Ícono cargado correctamente")
        except Exception as e:
            # Si no se puede cargar el ícono, se registra una advertencia
            logger.warning(f"No se pudo cargar el icono: {e}")

        # Llama a la función para construir la interfaz de usuario
        construir_gui(root)

        # Ejecuta el ciclo principal de la GUI
        root.mainloop()

        # Al finalizar la aplicación, registra que se cerró correctamente
        logger.info("Aplicación finalizada correctamente")

    except Exception:
        # Registra cualquier excepción que ocurra durante la ejecución
        logger.exception("Ocurrió un error crítico en main.py")


# Punto de entrada principal para ejecutar la aplicación
if __name__ == "__main__":
    main()