import sys
import json
import urllib.request
import ssl
import certifi
import time
import logging
import urllib.request
import shutil
import os
import subprocess
from draftsender_app import logger_utils as logger
from tkinter import messagebox

logger = logging.getLogger("DraftSender")

VERSION_FILE = os.path.join("data", "version.txt")
URL_API = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"

def descargar_actualizacion(url: str, nueva_version: str):
    """
    Descarga la nueva versión como 'DraftSender vX.Y.Z.exe',
    actualiza version.txt, lanza el nuevo exe y cierra la app actual.
    """
    try:
        carpeta = os.path.dirname(sys.argv[0])
        nuevo_nombre = f"DraftSender v{nueva_version}.exe"
        ruta_nuevo = os.path.join(carpeta, nuevo_nombre)

        # Descargar el nuevo .exe
        with urllib.request.urlopen(url) as response, open(ruta_nuevo, 'wb') as out_file:
            shutil.copyfileobj(response, out_file)

        if os.path.getsize(ruta_nuevo) < 5 * 1024 * 1024:
            raise ValueError("El ejecutable descargado parece incompleto.")

        logger.info(f"Nuevo ejecutable descargado: {ruta_nuevo}")

        # Crear o actualizar version.txt
        os.makedirs("data", exist_ok=True)
        with open(os.path.join("data", "version.txt"), "w", encoding="utf-8") as f:
            f.write(nueva_version)

        messagebox.showinfo(
            "Actualización completada",
            f"La nueva versión {nueva_version} fue descargada.\nLa app se cerrará para continuar con la actualización."
        )

        subprocess.Popen([ruta_nuevo], shell=True)
        sys.exit(0)

    except Exception as e:
        logger.error(f"Error en actualización: {e}")
        messagebox.showerror("Error", f"No se pudo completar la actualización:\n{e}")

def obtener_version_actual():
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, "w") as f:
            f.write("v0.0.0")
        return "v0.0.0"
    with open(VERSION_FILE, "r") as f:
        return f.read().strip()

def actualizar_version_local(nueva_version: str):
    with open(VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(nueva_version)


def crear_hook(base, avance, barra_progreso, porcentaje_var, root, status_var):
    """
    Hook para mostrar progreso de descarga con urllib.request.urlretrieve
    """
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

def verificar_actualizacion(root, barra_progreso, porcentaje_var, frame_progreso, status_label, status_var, forzar=False):
    try:
        logger.info("===== INICIO DE PROCESO DE ACTUALIZACIÓN =====")
        status_label.pack(side="bottom", pady=(0, 5))
        root.update_idletasks()
        root.geometry("")

        context = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(URL_API, context=context) as response:
            data = json.loads(response.read())

        ultima_version = data["tag_name"].strip()  # Ej: 'v1.0.3'
        assets = data["assets"]
        logger.info(f"Última versión disponible: {ultima_version}")

        version_local = obtener_version_actual()
        if not forzar and version_local == ultima_version:
            logger.info("Ya tienes la última versión.")
            status_var.set("Ya tienes la última versión instalada.")
            root.after(5000, lambda: status_var.set(""))
            frame_progreso.pack_forget()
            status_label.pack_forget()
            root.geometry("")
            logger.info("===== FIN DE PROCESO DE ACTUALIZACIÓN (sin cambios) =====")
            return

        if messagebox.askyesno("Actualización disponible", f"Hay una nueva versión ({ultima_version}). ¿Deseas descargarla ahora?"):
            ejecutable_actual = sys.executable
            if not ejecutable_actual.lower().endswith(".exe") or "python" in os.path.basename(ejecutable_actual).lower():
                logger.warning("Modo desarrollo detectado. No se reemplazará el ejecutable.")
                messagebox.showinfo("Modo desarrollo", "Estás ejecutando en modo desarrollo. No se realizará actualización automática.")
                frame_progreso.pack_forget()
                status_label.pack_forget()
                root.geometry("")
                logger.info("===== FIN DE PROCESO DE ACTUALIZACIÓN (modo desarrollo) =====")
                return

            # Buscar el asset cuyo nombre coincida con el formato 'DraftSender vX.Y.Z.exe'
            nombre_esperado = f"draftsender v{ultima_version.strip().lower()}.exe"
            asset_match = next((a for a in assets if a["name"].strip().lower() == nombre_esperado), None)

            if not asset_match:
                logger.warning(f"No se encontró el asset '{nombre_esperado}' en el release.")
                messagebox.showwarning("No disponible", f"No se encontró el archivo '{nombre_esperado}' para descargar.")
                return

            actualizar_version_local(ultima_version)
            descargar_actualizacion(url=asset_match["browser_download_url"], nueva_version=ultima_version)

    except Exception as e:
        logger.exception(f"Error al verificar actualización desde {URL_API}")
        messagebox.showerror("Error", f"No se pudo verificar actualización:\n{e}")
        status_var.set("Error al verificar actualización")
        root.after(5000, lambda: status_var.set(""))
        frame_progreso.pack_forget()
        status_label.pack_forget()
        root.geometry("")
        logger.info("===== FIN DE PROCESO DE ACTUALIZACIÓN (error inesperado) =====")

def verificar_version_disponible():
    """
    Compara la versión local con la versión más reciente publicada en GitHub.
    Retorna True si hay una versión nueva disponible.
    """
    try:
        context = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(URL_API, context=context) as response:
            data = json.loads(response.read())
            ultima_version = data["tag_name"].lstrip("v")

        version_local = obtener_version_actual()
        return version_local != ultima_version

    except Exception as e:
        logger.warning(f"No se pudo verificar la versión disponible: {e}")
        return False
