import sys
import json
import ssl
import certifi
import time
import logging
import shutil
import os
import subprocess
import urllib.request
import psutil
from tkinter import messagebox
from draftsender_app.version import obtener_version_local
from draftsender_app import logger_utils as logger

logger = logging.getLogger("DraftSender")

VERSION_FILE = os.path.join("data", "version.txt")
URL_API = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"
EXE_NAME_PREFIX = "DraftSender_v"

def cerrar_procesos_relacionados(nombre_exe):
    current_pid = os.getpid()
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['pid'] != current_pid and nombre_exe.lower() in proc.info['name'].lower():
            try:
                proc.terminate()
                proc.wait(timeout=5)
            except Exception:
                pass

def descargar_actualizacion(url: str, nueva_version: str, nombre_exe: str):
    try:
        carpeta = os.path.dirname(sys.argv[0])
        ruta_nuevo = os.path.join(carpeta, nombre_exe)

        with urllib.request.urlopen(url) as response, open(ruta_nuevo, 'wb') as out_file:
            shutil.copyfileobj(response, out_file)

        if os.path.getsize(ruta_nuevo) < 5 * 1024 * 1024:
            raise ValueError("El ejecutable descargado parece incompleto.")

        logger.info(f"Nuevo ejecutable descargado: {ruta_nuevo}")
        return ruta_nuevo
    except Exception as e:
        logger.error(f"Error al descargar ejecutable: {e}")
        raise

def reemplazar_y_reiniciar(ruta_nuevo_exe, nueva_version):
    try:
        exe_actual = sys.executable

        # Cerrar instancias activas
        cerrar_procesos_relacionados(os.path.basename(exe_actual))
        time.sleep(1)

        # Crear carpeta /data/versiones si no existe
        carpeta_versiones = os.path.join("data", "versiones")
        os.makedirs(carpeta_versiones, exist_ok=True)

        # Mover el ejecutable actual a /data/versiones/
        nombre_actual = os.path.basename(exe_actual)
        destino = os.path.join(carpeta_versiones, nombre_actual)
        try:
            shutil.move(exe_actual, destino)
        except Exception as e:
            logger.warning(f"No se pudo mover el ejecutable actual: {e}")

        # Guardar nueva versión en data/version.txt
        os.makedirs(os.path.dirname(VERSION_FILE), exist_ok=True)
        with open(VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(nueva_version)

        # Mover el nuevo ejecutable a la ruta del actual
        shutil.move(ruta_nuevo_exe, exe_actual)

        # Reiniciar aplicación
        subprocess.Popen([exe_actual], close_fds=True)
        sys.exit(0)

    except Exception as e:
        logger.error(f"Error durante reemplazo y reinicio: {e}")
        messagebox.showerror("Actualización", f"Ocurrió un error al actualizar la aplicación:\n{e}")

def ejecutar_actualizacion(forzar=False):
    try:
        logger.info("===== INICIO DE ACTUALIZACIÓN =====")

        context = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(URL_API, context=context) as response:
            data = json.loads(response.read())

        ultima_version = data["tag_name"].strip()  # Ej: "v1.0.0"
        assets = data["assets"]
        version_local = obtener_version_local()

        if not forzar and version_local == ultima_version:
            logger.info("Ya estás en la última versión.")
            messagebox.showinfo("Sin cambios", f"Ya tienes la versión {version_local}")
            return

        ejecutable_actual = sys.executable
        if not ejecutable_actual.lower().endswith(".exe") or "python" in os.path.basename(ejecutable_actual).lower():
            logger.warning("Modo desarrollo detectado.")
            messagebox.showinfo("Modo desarrollo", "No se puede actualizar automáticamente en modo desarrollo.")
            return

        # Corrección: remueve el prefijo "v" para buscar bien el ejecutable
        nombre_esperado = f"{EXE_NAME_PREFIX}{ultima_version.lstrip('v')}.exe"

        asset_match = next(
            (a for a in assets if a["name"].lower() == nombre_esperado.lower()),
            None
        )

        if not asset_match:
            logger.warning(f"No se encontró el asset '{nombre_esperado}' en GitHub.")
            messagebox.showwarning("No disponible", f"No se encontró el archivo '{nombre_esperado}' para descargar.")
            return

        ruta_descargada = descargar_actualizacion(asset_match["browser_download_url"], ultima_version, nombre_esperado)
        reemplazar_y_reiniciar(ruta_descargada, ultima_version)

    except Exception as e:
        logger.exception(f"Error general durante la actualización")
        messagebox.showerror("Error", f"No se pudo verificar actualización:\n{e}")

def hay_nueva_version_disponible() -> bool:
    """
    Consulta GitHub Releases y compara con la versión local.
    Devuelve True si hay una nueva versión disponible.
    """
    try:
        context = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(URL_API, context=context) as response:
            data = json.loads(response.read())
            ultima_version = data["tag_name"].strip()
        return obtener_version_local() != ultima_version
    except Exception as e:
        logger.warning(f"No se pudo verificar la versión disponible: {e}")
        return False

def crear_carpeta_versiones():
    ruta = os.path.join("data", "versiones")
    os.makedirs(ruta, exist_ok=True)

def mover_exe_antiguo(exe_actual):
    crear_carpeta_versiones()
    nombre = os.path.basename(exe_actual)
    destino = os.path.join("data", "versiones", nombre)
    shutil.move(exe_actual, destino)
