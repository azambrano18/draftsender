import json
import logging
import os
import shutil
import ssl
import subprocess
import sys
import urllib.request
from tkinter import messagebox
from typing import Dict, List, Optional

import certifi

from draftsender_app.version import (
    es_version_mayor,
    normalizar_version,
    obtener_base_app,
    obtener_data_path,
    obtener_version_local,
)


logger = logging.getLogger("DraftSender")

URL_API_RELEASE_LATEST = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"
EXE_NAME_PREFIX = "DraftSender_v"
MIN_EXE_SIZE_BYTES = 5 * 1024 * 1024
HTTP_TIMEOUT_SECONDS = 20


def esta_en_modo_desarrollo() -> bool:
    """
    Detecta si la app está corriendo desde Python/PyCharm y no como ejecutable final.

    El actualizador automático solo debe ejecutarse desde el .exe empaquetado.
    """
    ejecutable = os.path.basename(sys.executable).lower()

    if not getattr(sys, "frozen", False):
        return True

    if "python" in ejecutable:
        return True

    return False


def crear_contexto_ssl() -> ssl.SSLContext:
    """
    Crea un contexto SSL usando certificados de certifi.
    """
    return ssl.create_default_context(cafile=certifi.where())


def consultar_release_latest() -> Dict:
    """
    Consulta la última release publicada en GitHub.

    Returns:
        Dict con la respuesta JSON de GitHub.

    Raises:
        RuntimeError: si GitHub responde con datos incompletos.
    """
    request = urllib.request.Request(
        URL_API_RELEASE_LATEST,
        headers={
            "User-Agent": "DraftSender-Updater",
            "Accept": "application/vnd.github+json",
        },
    )

    context = crear_contexto_ssl()

    with urllib.request.urlopen(request, context=context, timeout=HTTP_TIMEOUT_SECONDS) as response:
        data = json.loads(response.read().decode("utf-8"))

    if not data.get("tag_name"):
        raise RuntimeError("La respuesta de GitHub no contiene tag_name.")

    if "assets" not in data:
        raise RuntimeError("La respuesta de GitHub no contiene assets.")

    return data


def construir_nombre_exe(version: str) -> str:
    """
    Construye el nombre esperado del ejecutable según la versión.

    Ejemplo:
        v1.2.0 -> DraftSender_v1.2.0.exe
    """
    version_normalizada = normalizar_version(version)
    return f"{EXE_NAME_PREFIX}{version_normalizada.lstrip('v')}.exe"


def buscar_asset_ejecutable(assets: List[Dict], nombre_esperado: str) -> Optional[Dict]:
    """
    Busca dentro de los assets de la release el ejecutable esperado.
    """
    for asset in assets:
        nombre_asset = str(asset.get("name", "")).strip()

        if nombre_asset.lower() == nombre_esperado.lower():
            return asset

    return None


def obtener_ruta_updates() -> str:
    """
    Devuelve la carpeta donde se descargan temporalmente las actualizaciones.
    """
    ruta = os.path.join(obtener_data_path(), "updates")
    os.makedirs(ruta, exist_ok=True)
    return ruta


def limpiar_actualizaciones_temporales() -> None:
    """
    Limpia ejecutables temporales de actualizaciones anteriores.
    """
    ruta_updates = obtener_ruta_updates()

    for nombre in os.listdir(ruta_updates):
        if nombre.lower().endswith(".exe"):
            ruta = os.path.join(ruta_updates, nombre)
            try:
                os.remove(ruta)
            except OSError:
                logger.warning("No se pudo eliminar actualización temporal: %s", ruta)


def descargar_actualizacion(url: str, nombre_exe: str) -> str:
    """
    Descarga el nuevo ejecutable en data/updates.

    Args:
        url: browser_download_url del asset.
        nombre_exe: nombre local del ejecutable.

    Returns:
        Ruta absoluta del ejecutable descargado.
    """
    limpiar_actualizaciones_temporales()

    ruta_destino = os.path.join(obtener_ruta_updates(), nombre_exe)
    ruta_parcial = f"{ruta_destino}.download"

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "DraftSender-Updater",
            "Accept": "application/octet-stream",
        },
    )

    context = crear_contexto_ssl()

    try:
        logger.info("Descargando actualización desde GitHub: %s", url)

        with urllib.request.urlopen(request, context=context, timeout=HTTP_TIMEOUT_SECONDS) as response:
            with open(ruta_parcial, "wb") as archivo:
                shutil.copyfileobj(response, archivo)

        if not os.path.exists(ruta_parcial):
            raise RuntimeError("No se creó el archivo descargado.")

        tamaño = os.path.getsize(ruta_parcial)

        if tamaño < MIN_EXE_SIZE_BYTES:
            raise RuntimeError(
                f"El ejecutable descargado parece incompleto. Tamaño: {tamaño} bytes."
            )

        if os.path.exists(ruta_destino):
            os.remove(ruta_destino)

        os.replace(ruta_parcial, ruta_destino)

        logger.info("Actualización descargada correctamente: %s", ruta_destino)
        return ruta_destino

    except Exception:
        try:
            if os.path.exists(ruta_parcial):
                os.remove(ruta_parcial)
        except OSError:
            pass

        logger.exception("Error al descargar actualización.")
        raise


def crear_script_bat_actualizador(
    ruta_nuevo_exe: str,
    ruta_exe_actual: str,
    nueva_version: str,
) -> str:
    """
    Crea un script .bat temporal para reemplazar el ejecutable actual.

    Motivo:
        En Windows, un .exe no puede reemplazarse a sí mismo mientras está abierto.
        Por eso se crea un proceso externo que espera el cierre de la app,
        mueve el .exe viejo a data/versiones y pone el nuevo en su lugar.
    """
    data_path = obtener_data_path()
    carpeta_versiones = os.path.join(data_path, "versiones")
    os.makedirs(carpeta_versiones, exist_ok=True)

    ruta_bat = os.path.join(data_path, "actualizar_draftsender.bat")
    ruta_log = os.path.join(data_path, "actualizacion.log")

    pid_actual = os.getpid()
    nombre_actual = os.path.basename(ruta_exe_actual)
    ruta_backup = os.path.join(carpeta_versiones, f"{nombre_actual}.old")

    contenido = f"""@echo off
setlocal enabledelayedexpansion

set "PID_ACTUAL={pid_actual}"
set "NUEVO_EXE={ruta_nuevo_exe}"
set "EXE_ACTUAL={ruta_exe_actual}"
set "BACKUP_EXE={ruta_backup}"
set "DATA_PATH={data_path}"
set "VERSION={normalizar_version(nueva_version)}"
set "LOG_FILE={ruta_log}"

echo ======================================== >> "%LOG_FILE%"
echo Iniciando actualizacion DraftSender >> "%LOG_FILE%"
echo PID actual: %PID_ACTUAL% >> "%LOG_FILE%"
echo Nuevo EXE: %NUEVO_EXE% >> "%LOG_FILE%"
echo EXE actual: %EXE_ACTUAL% >> "%LOG_FILE%"

set /a INTENTOS=0

:ESPERAR_PROCESO
tasklist /FI "PID eq %PID_ACTUAL%" | find "%PID_ACTUAL%" >nul
if not errorlevel 1 (
    set /a INTENTOS+=1
    echo Esperando cierre de DraftSender. Intento !INTENTOS! >> "%LOG_FILE%"
    timeout /t 1 /nobreak >nul

    if !INTENTOS! GEQ 60 (
        echo ERROR: El proceso no cerro a tiempo. >> "%LOG_FILE%"
        exit /b 1
    )

    goto ESPERAR_PROCESO
)

timeout /t 2 /nobreak >nul

if not exist "%NUEVO_EXE%" (
    echo ERROR: No existe el nuevo ejecutable. >> "%LOG_FILE%"
    exit /b 2
)

if exist "%BACKUP_EXE%" (
    del /F /Q "%BACKUP_EXE%" >> "%LOG_FILE%" 2>&1
)

if exist "%EXE_ACTUAL%" (
    move /Y "%EXE_ACTUAL%" "%BACKUP_EXE%" >> "%LOG_FILE%" 2>&1
    if errorlevel 1 (
        echo ERROR: No se pudo mover el ejecutable anterior. >> "%LOG_FILE%"
        exit /b 3
    )
)

move /Y "%NUEVO_EXE%" "%EXE_ACTUAL%" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    echo ERROR: No se pudo instalar el nuevo ejecutable. >> "%LOG_FILE%"

    if exist "%BACKUP_EXE%" (
        move /Y "%BACKUP_EXE%" "%EXE_ACTUAL%" >> "%LOG_FILE%" 2>&1
    )

    exit /b 4
)

if not exist "%DATA_PATH%" (
    mkdir "%DATA_PATH%"
)

echo %VERSION%> "%DATA_PATH%\\version.txt"

echo Actualizacion completada correctamente. >> "%LOG_FILE%"
start "" "%EXE_ACTUAL%"

timeout /t 2 /nobreak >nul
del "%~f0"
"""

    with open(ruta_bat, "w", encoding="utf-8") as archivo:
        archivo.write(contenido)

    return ruta_bat


def reemplazar_y_reiniciar(ruta_nuevo_exe: str, nueva_version: str) -> None:
    """
    Lanza el script externo de actualización y cierra la app actual.
    """
    ruta_exe_actual = os.path.abspath(sys.executable)

    if not os.path.exists(ruta_nuevo_exe):
        raise FileNotFoundError(f"No existe el nuevo ejecutable: {ruta_nuevo_exe}")

    if not os.path.exists(ruta_exe_actual):
        raise FileNotFoundError(f"No existe el ejecutable actual: {ruta_exe_actual}")

    ruta_bat = crear_script_bat_actualizador(
        ruta_nuevo_exe=ruta_nuevo_exe,
        ruta_exe_actual=ruta_exe_actual,
        nueva_version=nueva_version,
    )

    logger.info("Lanzando actualizador externo: %s", ruta_bat)

    creationflags = 0

    if os.name == "nt":
        creationflags = subprocess.CREATE_NEW_CONSOLE

    subprocess.Popen(
        ["cmd.exe", "/c", ruta_bat],
        cwd=obtener_base_app(),
        close_fds=True,
        creationflags=creationflags,
    )

    sys.exit(0)


def hay_nueva_version_disponible() -> bool:
    """
    Consulta GitHub Releases y devuelve True si hay una versión superior disponible.
    """
    try:
        data = consultar_release_latest()
        ultima_version = normalizar_version(data["tag_name"])
        version_local = obtener_version_local()

        return es_version_mayor(ultima_version, version_local)

    except Exception as error:
        logger.warning("No se pudo verificar si hay nueva versión: %s", error)
        return False


def ejecutar_actualizacion(forzar: bool = False) -> None:
    """
    Verifica, descarga e instala una actualización desde GitHub Releases.

    Args:
        forzar:
            Si es True, permite descargar la release aunque la versión local
            no sea menor. No permite actualizar desde PyCharm/desarrollo.
    """
    try:
        logger.info("===== INICIO DE VERIFICACIÓN DE ACTUALIZACIÓN =====")

        if esta_en_modo_desarrollo():
            logger.warning("Actualización cancelada: modo desarrollo detectado.")
            messagebox.showinfo(
                "Modo desarrollo",
                "La actualización automática solo funciona desde el ejecutable instalado, no desde PyCharm.",
            )
            return

        data = consultar_release_latest()

        ultima_version = normalizar_version(data["tag_name"])
        version_local = obtener_version_local()
        assets = data.get("assets", [])

        logger.info("Versión local: %s | Última versión: %s", version_local, ultima_version)

        if not forzar and not es_version_mayor(ultima_version, version_local):
            messagebox.showinfo(
                "Sin actualizaciones",
                f"Ya tienes instalada la versión {version_local}.",
            )
            return

        nombre_esperado = construir_nombre_exe(ultima_version)
        asset = buscar_asset_ejecutable(assets, nombre_esperado)

        if not asset:
            logger.warning("No se encontró asset esperado: %s", nombre_esperado)
            messagebox.showwarning(
                "Actualización no disponible",
                "Se encontró una nueva release, pero no está disponible el ejecutable esperado:\n\n"
                f"{nombre_esperado}\n\n"
                "Revisa el nombre del archivo subido a GitHub Releases.",
            )
            return

        url_descarga = asset.get("browser_download_url")

        if not url_descarga:
            raise RuntimeError("El asset no contiene browser_download_url.")

        confirmar = messagebox.askyesno(
            "Actualización disponible",
            f"Hay una nueva versión disponible.\n\n"
            f"Versión actual: {version_local}\n"
            f"Nueva versión: {ultima_version}\n\n"
            "¿Quieres descargarla e instalarla ahora?",
        )

        if not confirmar:
            logger.info("El usuario canceló la actualización.")
            return

        ruta_descargada = descargar_actualizacion(
            url=url_descarga,
            nombre_exe=nombre_esperado,
        )

        messagebox.showinfo(
            "Actualización descargada",
            "La actualización se descargó correctamente.\n\n"
            "DraftSender se cerrará y volverá a abrirse automáticamente.",
        )

        reemplazar_y_reiniciar(
            ruta_nuevo_exe=ruta_descargada,
            nueva_version=ultima_version,
        )

    except Exception as error:
        logger.exception("Error general durante la actualización.")
        messagebox.showerror(
            "Error de actualización",
            f"No se pudo completar la actualización:\n\n{error}",
        )
