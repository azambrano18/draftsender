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

def descargar_actualizacion(url: str):
    """
    Descarga el nuevo .exe desde la URL proporcionada y lanza un script por lotes
    que espera que el ejecutable actual se cierre antes de reemplazarlo.
    """
    try:
        nombre_exe = os.path.basename(sys.argv[0])
        carpeta = os.path.dirname(sys.argv[0])
        ruta_nueva = os.path.join(carpeta, f"{nombre_exe}_update.exe")
        ruta_bat = os.path.join(carpeta, "updater.bat")

        # Descargar el nuevo ejecutable
        with urllib.request.urlopen(url) as response, open(ruta_nueva, 'wb') as out_file:
            shutil.copyfileobj(response, out_file)
        logger.info(f"Descargado {ruta_nueva}")

        # Crear .bat que espera cierre del exe, reemplaza y relanza
        with open(ruta_bat, "w") as f:
            f.write(f"""@echo off
echo Esperando que {nombre_exe} se cierre...
:espera
tasklist | findstr /i "{nombre_exe}" >nul
if not errorlevel 1 (
    timeout /t 2 >nul
    goto espera
)
del "{nombre_exe}" /f
rename "{nombre_exe}_update.exe" "{nombre_exe}"
start "" "{nombre_exe}"
del "%~f0"
""")

        logger.info(f"Script de actualización creado: {ruta_bat}")
        subprocess.Popen(['cmd', '/c', 'start', '', ruta_bat], shell=True)
        messagebox.showinfo("Actualización", "Se descargó la nueva versión.\nLa app se cerrará para completar la actualización.")
        sys.exit(0)

    except Exception as e:
        logger.error(f"Error al descargar la actualización: {e}")
        messagebox.showerror("Error", f"No se pudo descargar la actualización:\n{e}")

def obtener_version_actual():
    """
    Obtiene la versión actual desde version.txt. Si no existe, lo crea con v0.0.0.
    """
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, "w") as f:
            f.write("v0.0.0")
        return "0.0.0"
    with open(VERSION_FILE, "r") as f:
        return f.read().strip().lstrip("v")

def actualizar_version_local(nueva_version: str):
    """
    Actualiza el archivo version.txt con la versión descargada.
    """
    with open(VERSION_FILE, "w") as f:
        f.write(f"v{nueva_version}")

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

        nombre_exe_actual = os.path.basename(sys.argv[0])
        if data.get("name") != nombre_exe_actual:
            logger.warning(f"El release no coincide con el nombre esperado ({nombre_exe_actual}). Abortando.")
            messagebox.showwarning("Versión incompatible", f"No se encontró un release que coincida con '{nombre_exe_actual}'.")
            return

        ultima_version = data["tag_name"].lstrip("v")
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

            # Buscar asset que coincida exactamente con el nombre del ejecutable actual
            asset_match = next((a for a in assets if a["name"] == nombre_exe_actual), None)
            if not asset_match:
                logger.warning("No se encontró un archivo .exe que coincida exactamente con el ejecutable actual.")
                messagebox.showwarning("No disponible", f"No se encontró el archivo '{nombre_exe_actual}' para descargar.")
                return

            actualizar_version_local(ultima_version)
            descargar_actualizacion(url=asset_match["browser_download_url"])

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
