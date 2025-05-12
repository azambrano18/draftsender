import os  # Para interactuar con el sistema de archivos
import sys  # Para terminar el programa o acceder a su ubicación
import json  # Para decodificar la respuesta JSON desde la API de GitHub
import urllib.request  # Para descargar archivos desde internet
import subprocess  # Para ejecutar la nueva versión del programa
import time  # asegúrate que está importado al inicio
from tkinter import messagebox  # Para mostrar mensajes en la interfaz gráfica

from logger_utils import configurar_logger  # Logger personalizado
from version import __version__  # Versión actual centralizada del programa

# Instancia del logger para este módulo
logger = configurar_logger("actualizacion")

def verificar_actualizacion(root, barra_progreso, porcentaje_var, frame_progreso, forzar=False):
    """
    Verifica si hay una nueva versión disponible del programa en GitHub.
    Si la hay, permite al usuario descargarla y ejecutarla.
    Args:
        root (tk.Tk): Ventana principal de la GUI.
        barra_progreso (tk.Progressbar): Barra de progreso de descarga.
        porcentaje_var (tk.StringVar): Variable para mostrar el porcentaje.
        frame_progreso (tk.Frame): Contenedor del progreso.
        forzar (bool): Si True, descarga la versión incluso si ya está actualizada.
    """
    url_api = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"

    try:
        logger.info("Verificando actualización...")

        # Solicita los datos de la última release desde GitHub
        with urllib.request.urlopen(url_api) as response:
            data = json.loads(response.read())
            ultima_version = data["tag_name"].lstrip("v")  # Quita la 'v' inicial
            assets = data["assets"]
            logger.info(f"Última versión disponible: {ultima_version}")

            # Inicializa progreso visual
            barra_progreso["value"] = 0
            porcentaje_var.set("0%")
            barra_progreso.pack(side="left", padx=(0, 10))
            frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
            root.update_idletasks()

        # Verifica si es necesario actualizar
        if forzar or ultima_version != obtener_version_actual():
            logger.info("Nueva versión detectada. Solicitando confirmación al usuario.")

            if messagebox.askyesno("Actualización disponible", f"Hay una nueva versión ({ultima_version}). ¿Deseas descargarla ahora?"):
                exe_dir = os.path.dirname(sys.executable)

                # Archivos esperados en la release
                archivos = {
                    "main.exe": "CreadorBorradores_Nuevo.exe",
                    "txt_1.exe": "txt_1_nuevo.exe",
                    "timer_sent.exe": "timer_sent_nuevo.exe"
                }

                # Filtra los que están presentes en GitHub
                descargas = [a for a in assets if a["name"] in archivos]
                if not descargas:
                    logger.warning("No se encontraron archivos para descargar.")
                    messagebox.showwarning("No hay archivos", "No se encontraron archivos para actualizar.")
                    return

                avance = 100 // len(descargas)
                base = 0

                # Descarga cada archivo
                for asset in descargas:
                    nombre = asset["name"]
                    url = asset["browser_download_url"]
                    destino = os.path.join(exe_dir, archivos[nombre])
                    logger.info(f"Descargando {nombre} desde {url} a {destino}")
                    try:
                        urllib.request.urlretrieve(
                            url,
                            destino,
                            reporthook=crear_hook(base, avance, barra_progreso, porcentaje_var, root, status_var)
                        )
                    except Exception as e:
                        logger.error(f"Error al descargar el archivo {nombre}: {e}")
                        messagebox.showerror("Error de descarga", f"No se pudo descargar el archivo {nombre}. Intenta nuevamente.")
                        return
                    base += avance

                # Finaliza la descarga y lanza nueva versión
                barra_progreso["value"] = 100
                porcentaje_var.set("100%")
                status_var.set("Actualización descargada y aplicada.")
                root.update_idletasks()

                # Oculta el mensaje luego de 5 segundos
                def ocultar_mensaje():
                    try:
                        status_var.set("")
                    except:
                        pass

                root.after(5000, ocultar_mensaje)

                frame_progreso.pack_forget()
                logger.info("Descarga completada. Lanzando nueva versión.")
                messagebox.showinfo("Actualización", "Se lanzará la nueva versión ahora.")
                subprocess.Popen([os.path.join(exe_dir, "CreadorBorradores_Nuevo.exe")])
                sys.exit()
        else:
            logger.info("Ya tienes la última versión.")
    except Exception as e:
        logger.exception(f"Fallo en la verificación de actualización desde {url_api}")
        messagebox.showerror("Error", f"No se pudo verificar actualización:\n{e}")

def obtener_version_actual():
    """
    Retorna la versión actual del programa usando el valor definido en version.py.
    Returns:
        str: Cadena con la versión actual, ej. "1.0.1"
    """
    return __version__


def crear_hook(base, avance, barra_progreso, porcentaje_var, root, status_var):
    inicio = time.time()

    def hook(count, block_size, total_size):
        if total_size > 0:
            porcentaje = int((count * block_size * 100) / total_size)
            total = min(100, base + int(porcentaje * avance / 100))
            barra_progreso["value"] = total
            porcentaje_var.set(f"{total}%")

            tiempo = time.time() - inicio
            velocidad_kb = (count * block_size) / 1024 / tiempo if tiempo > 0 else 0
            status_var.set(f"Descargando... {velocidad_kb:.1f} KB/s")

            root.update_idletasks()
    return hook