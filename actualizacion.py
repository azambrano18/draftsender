import os
import sys
import json
import urllib.request
import subprocess
import time
import tkinter as tk
from tkinter import messagebox

from logger_utils import configurar_logger
from version import __version__

logger = configurar_logger("actualizacion")

def verificar_actualizacion(root, barra_progreso, porcentaje_var, frame_progreso, status_label, status_var, forzar=False):
    """
    Verifica si hay una nueva versión disponible del programa en GitHub.
    Descarga draftsender.exe si es necesario y lo ejecuta como draftsender_<versión>.exe
    """
    url_api = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"

    try:
        logger.info("Verificando actualización...")
        status_var.set("Verificando actualizaciones...")
        status_label.pack(side="bottom", pady=(0, 5))
        root.update_idletasks()

        with urllib.request.urlopen(url_api) as response:
            data = json.loads(response.read())
            ultima_version = data["tag_name"].lstrip("v")
            assets = data["assets"]
            logger.info(f"Última versión disponible: {ultima_version}")

            barra_progreso["value"] = 0
            porcentaje_var.set("0%")
            frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
            root.update_idletasks()

        if forzar or ultima_version != obtener_version_actual():
            logger.info("Nueva versión detectada. Solicitando confirmación al usuario.")

            if messagebox.askyesno("Actualización disponible", f"Hay una nueva versión ({ultima_version}). ¿Deseas descargarla ahora?"):
                exe_dir = os.path.dirname(sys.executable)

                archivos = {
                    "draftsender.exe": f"draftsender_{ultima_version}.exe"
                }

                descargas = [a for a in assets if a["name"] in archivos]
                if not descargas:
                    logger.warning("No se encontraron archivos para descargar.")
                    messagebox.showwarning("No hay archivos", "No se encontraron archivos para actualizar.")
                    frame_progreso.pack_forget()
                    status_label.pack_forget()
                    return

                avance = 100 // len(descargas)
                base = 0

                for asset in descargas:
                    nombre = asset["name"]
                    url = asset["browser_download_url"]
                    destino = os.path.join(exe_dir, archivos[nombre])

                    logger.info(f"Descargando {nombre} desde {url} a {destino}")
                    status_var.set(f"Descargando {nombre}...")
                    root.update_idletasks()

                    try:
                        urllib.request.urlretrieve(
                            url,
                            destino,
                            reporthook=crear_hook(base, avance, barra_progreso, porcentaje_var, root, status_var)
                        )
                    except Exception as e:
                        logger.error(f"Error al descargar el archivo {nombre}: {e}")
                        messagebox.showerror("Error de descarga", f"No se pudo descargar el archivo {nombre}. Intenta nuevamente.")
                        frame_progreso.pack_forget()
                        status_label.pack_forget()
                        return

                    base += avance

                barra_progreso["value"] = 100
                porcentaje_var.set("100%")
                status_var.set("Actualización descargada y aplicada.")
                root.update_idletasks()

                def ocultar_mensaje():
                    try:
                        status_var.set("")
                        frame_progreso.pack_forget()
                        status_label.pack_forget()
                    except Exception as e:
                        logger.warning(f"Error al ocultar elementos: {e}")

                root.after(5000, ocultar_mensaje)

                nuevo_path = os.path.join(exe_dir, f"draftsender_{ultima_version}.exe")
                logger.info(f"Ejecución de la nueva versión: {nuevo_path}")
                messagebox.showinfo("Actualización", "Se lanzará la nueva versión ahora.")
                subprocess.Popen([nuevo_path])
                sys.exit()
        else:
            logger.info("Ya tienes la última versión.")
            status_var.set("Ya tienes la última versión instalada.")
            root.after(5000, lambda: status_var.set(""))
            frame_progreso.pack_forget()
            status_label.pack_forget()

    except Exception as e:
        logger.exception(f"Fallo en la verificación de actualización desde {url_api}")
        messagebox.showerror("Error", f"No se pudo verificar actualización:\n{e}")
        status_var.set("Error al verificar actualización")
        root.after(5000, lambda: status_var.set(""))
        frame_progreso.pack_forget()
        status_label.pack_forget()

def obtener_version_actual():
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