import os  # Manejo de rutas de archivos
import sys  # Acceso a funciones del sistema
import json  # Para procesar respuestas JSON
import urllib.request  # Para hacer solicitudes HTTP
import subprocess  # Para ejecutar procesos externos
import time  # Medir tiempo y calcular velocidad
from tkinter import messagebox  # Cuadros de diálogo para el usuario

from logger_utils import configurar_logger  # Función personalizada para logging
from version import __version__  # Importa la versión actual del sistema

logger = configurar_logger("actualizacion")  # Función personalizada para logging

def verificar_actualizacion(root, barra_progreso, porcentaje_var, frame_progreso, status_label, status_var, forzar=False):  # Verifica si hay una nueva versión y la descarga si es necesario
    """
    Verifica si hay una nueva versión disponible del programa en GitHub.
    Descarga draftsender.exe si es necesario y lo ejecuta como draftsender_<versión>.exe
    """
    url_api = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"  # URL del API de GitHub para obtener la última versión

    try:
        logger.info("Verificando actualización...")
        status_var.set("Verificando actualizaciones...")
        status_label.pack(side="bottom", pady=(0, 5))
        root.update_idletasks()

        with urllib.request.urlopen(url_api) as response:  # Realiza la solicitud HTTP y procesa la respuesta JSON
            data = json.loads(response.read())
            ultima_version = data["tag_name"].lstrip("v")
            assets = data["assets"]
            logger.info(f"Última versión disponible: {ultima_version}")

            barra_progreso["value"] = 0
            porcentaje_var.set("0%")
            frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
            root.update_idletasks()

        if forzar or ultima_version != obtener_version_actual():  # Compara la versión actual con la última disponible
            logger.info("Nueva versión detectada. Solicitando confirmación al usuario.")

            if messagebox.askyesno("Actualización disponible", f"Hay una nueva versión ({ultima_version}). ¿Deseas descargarla ahora?"):  # Cuadros de diálogo para el usuario
                exe_dir = os.path.dirname(sys.executable)  # Directorio donde está ubicado el ejecutable actual

                archivos = {
                    "draftsender.exe": f"draftsender_{ultima_version}.exe"
                }

                descargas = [a for a in assets if a["name"] in archivos]  # Filtra los archivos a descargar desde los assets
                if not descargas:
                    logger.warning("No se encontraron archivos para descargar.")
                    messagebox.showwarning("No hay archivos", "No se encontraron archivos para actualizar.")  # Cuadros de diálogo para el usuario
                    frame_progreso.pack_forget()
                    status_label.pack_forget()
                    return

                avance = 100 // len(descargas)
                base = 0

                for asset in descargas:  # Itera sobre cada archivo que se debe descargar
                    nombre = asset["name"]
                    url = asset["browser_download_url"]
                    destino = os.path.join(exe_dir, archivos[nombre])

                    logger.info(f"Descargando {nombre} desde {url} a {destino}")
                    root.update_idletasks()

                    try:
                        urllib.request.urlretrieve(
                            url,
                            destino,
                            reporthook=crear_hook(base, avance, barra_progreso, porcentaje_var, root, status_var)
                        )
                    except Exception as e:
                        logger.error(f"Error al descargar el archivo {nombre}: {e}")
                        messagebox.showerror("Error de descarga", f"No se pudo descargar el archivo {nombre}. Intenta nuevamente.")  # Cuadros de diálogo para el usuario
                        frame_progreso.pack_forget()
                        status_label.pack_forget()
                        return

                    base += avance

                barra_progreso["value"] = 100  # Marca el progreso al 100%
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

                root.after(5000, ocultar_mensaje)  # Oculta elementos después de 5 segundos

                nuevo_path = os.path.join(exe_dir, f"draftsender_{ultima_version}.exe")
                logger.info(f"Ejecución de la nueva versión: {nuevo_path}")
                messagebox.showinfo("Actualización", "Se lanzará la nueva versión ahora.")  # Cuadros de diálogo para el usuario
                subprocess.Popen([nuevo_path])  # Lanza el nuevo ejecutable descargado
                sys.exit()
        else:
            logger.info("Ya tienes la última versión.")
            status_var.set("Ya tienes la última versión instalada.")
            root.after(5000, lambda: status_var.set(""))  # Oculta elementos después de 5 segundos
            frame_progreso.pack_forget()
            status_label.pack_forget()

    except Exception as e:
        logger.exception(f"Fallo en la verificación de actualización desde {url_api}")
        messagebox.showerror("Error", f"No se pudo verificar actualización:\n{e}")  # Cuadros de diálogo para el usuario
        status_var.set("Error al verificar actualización")
        root.after(5000, lambda: status_var.set(""))  # Oculta elementos después de 5 segundos
        frame_progreso.pack_forget()
        status_label.pack_forget()

def obtener_version_actual():  # Retorna la versión actual del sistema
    return __version__  # Importa la versión actual del sistema

def crear_hook(base, avance, barra_progreso, porcentaje_var, root, status_var):  # Crea una función para actualizar la barra de progreso durante la descarga
    inicio = time.time()

    def hook(count, block_size, total_size):  # Función hook que actualiza la GUI mientras se descarga el archivo
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