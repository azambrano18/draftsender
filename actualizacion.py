import os  # Para interactuar con el sistema de archivos
import sys  # Para acceder a información sobre el sistema y terminar el programa
import json  # Para manejar la respuesta JSON de la API de GitHub
import urllib.request  # Para realizar solicitudes HTTP
import subprocess  # Para ejecutar la nueva versión del programa
from tkinter import messagebox  # Para mostrar cuadros de mensaje en la interfaz gráfica

from logger_utils import configurar_logger  # Para configurar el logger
logger = configurar_logger("actualizacion")  # Instancia del logger para este módulo

# Versión actual del programa
__version__ = "1.0.0" # cambiar la version por cada empaquetamiento

def verificar_actualizacion(root, barra_progreso, porcentaje_var, frame_progreso, forzar=False):
    """
    Verifica si hay una nueva versión disponible del programa en GitHub. Si la hay, permite al usuario descargarla e instalarla.
    Args:
        root (tk.Tk): La ventana principal de la aplicación.
        barra_progreso (tk.Progressbar): Barra de progreso para mostrar el avance de la descarga.
        porcentaje_var (tk.StringVar): Variable que contiene el porcentaje de progreso.
        frame_progreso (tk.Frame): Marco que contiene la barra de progreso.
        forzar (bool): Si se establece como True, forzará la descarga de la actualización incluso si ya está en la última versión.
    Returns:
        None
    """
    url_api = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"  # URL de la API de GitHub
    try:
        logger.info("Verificando actualización...")
        with urllib.request.urlopen(url_api) as response:
            data = json.loads(response.read())  # Obtiene la última versión desde la API
            ultima_version = data["tag_name"].lstrip("v")  # Elimina el prefijo 'v' de la versión
            assets = data["assets"]  # Obtiene los archivos de la nueva versión
            logger.info(f"Última versión disponible: {ultima_version}")

            # Inicializa la barra de progreso
            barra_progreso["value"] = 0
            porcentaje_var.set("0%")
            barra_progreso.pack(side="left", padx=(0, 10))
            frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
            root.update_idletasks()

        if forzar or ultima_version != obtener_version_actual():  # Si hay una nueva versión o se fuerza la actualización
            logger.info("Nueva versión detectada. Solicitando confirmación al usuario.")
            if messagebox.askyesno("Actualización disponible", f"Hay una nueva versión ({ultima_version}). ¿Deseas descargarla ahora?"):
                exe_dir = os.path.dirname(sys.executable)  # Directorio donde se ejecuta el programa
                archivos = {
                    "main.exe": "CreadorBorradores_Nuevo.exe",
                    "txt_1.exe": "txt_1_nuevo.exe",
                    "timer_sent.exe": "timer_sent_nuevo.exe"
                }

                # Filtra los archivos que se deben descargar
                descargas = [a for a in assets if a["name"] in archivos]
                if not descargas:
                    logger.warning("No se encontraron archivos para descargar.")
                    messagebox.showwarning("No hay archivos", "No se encontraron archivos para actualizar.")
                    return

                avance = 100 // len(descargas)
                base = 0

                # Descarga cada archivo de la nueva versión
                for asset in descargas:
                    nombre = asset["name"]
                    url = asset["browser_download_url"]
                    destino = os.path.join(exe_dir, archivos[nombre])
                    logger.info(f"Descargando {nombre} desde {url} a {destino}")
                    try:
                        urllib.request.urlretrieve(
                            url,
                            destino,
                            reporthook=crear_hook(base, avance, barra_progreso, porcentaje_var, root)  # Muestra el progreso de la descarga
                        )
                    except Exception as e:
                        logger.error(f"Error al descargar el archivo {nombre}: {e}")
                        messagebox.showerror("Error de descarga", f"No se pudo descargar el archivo {nombre}. Intenta nuevamente.")
                        return
                    base += avance

                # Actualiza la barra de progreso y muestra el mensaje de éxito
                barra_progreso["value"] = 100
                porcentaje_var.set("100%")
                root.update_idletasks()
                frame_progreso.pack_forget()

                logger.info("Descarga completada. Lanzando nueva versión.")
                messagebox.showinfo("Actualización", "Se lanzará la nueva versión ahora.")
                subprocess.Popen([os.path.join(exe_dir, "CreadorBorradores_Nuevo.exe")])  # Lanza la nueva versión
                sys.exit()  # Termina el programa actual
        else:
            logger.info("Ya tienes la última versión.")
    except Exception as e:
        logger.exception(f"Fallo en la verificación de actualización desde {url_api}")
        messagebox.showerror("Error", f"No se pudo verificar actualización:\n{e}")

def obtener_version_actual():
    """
    Obtiene la versión actual del programa desde el archivo `main.py`.
    Returns:
        str: La versión actual del programa.
    """
    try:
        from main import __version__  # Intenta obtener la versión desde el archivo main.py
        return __version__
    except Exception as e:
        logger.warning("No se pudo obtener la versión desde main.py")
        return "0.0.0"  # Si no se puede obtener, retorna una versión predeterminada

def crear_hook(base, avance, barra_progreso, porcentaje_var, root):
    """
    Crea un hook para la descarga de archivos que actualiza la barra de progreso.
    Args:
        base (int): La base del progreso total.
        avance (int): El porcentaje de avance por archivo.
        barra_progreso (tk.Progressbar): Barra de progreso para mostrar el avance de la descarga.
        porcentaje_var (tk.StringVar): Variable de texto para mostrar el porcentaje.
        root (tk.Tk): Ventana principal de la aplicación.
    Returns:
        function: Una función de hook que se usa durante la descarga.
    """
    def hook(count, block_size, total_size):
        if total_size > 0:
            porcentaje = int((count * block_size * 100) / total_size)
            total = min(100, base + int(porcentaje * avance / 100))  # Calcula el progreso total
            barra_progreso["value"] = total
            porcentaje_var.set(f"{total}%")
            root.update_idletasks()
    return hook