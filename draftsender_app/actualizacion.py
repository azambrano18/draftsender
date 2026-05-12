import logging
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict

import requests

from draftsender_app.version import (
    obtener_version_local,
    normalizar_version,
    es_version_mayor,
)


GITHUB_API_LATEST = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"

INSTALL_DIR = Path(r"C:\DraftSender_app")
DATA_DIR = INSTALL_DIR / "data"
LOGS_DIR = DATA_DIR / "logs"

APP_EXE_PREFIX = "DraftSender"
SHORTCUT_NAME = "DraftSender.lnk"

LOGS_DIR.mkdir(parents=True, exist_ok=True)

updater_logger = logging.getLogger("DraftSenderUpdater")
updater_logger.setLevel(logging.INFO)

if not updater_logger.handlers:
    file_handler = logging.FileHandler(LOGS_DIR / "updater.log", encoding="utf-8")
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - %(message)s"
    )
    file_handler.setFormatter(formatter)
    updater_logger.addHandler(file_handler)


def obtener_directorio_app() -> Path:
    """
    actualizacion.py > obtener_directorio_app
    Devuelve la carpeta donde está ejecutándose la app.
    En producción debe ser C:\\DraftSender_app.
    """
    import sys

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return INSTALL_DIR


def obtener_ejecutable_actual() -> Path:
    """
    actualizacion.py > obtener_ejecutable_actual
    Devuelve la ruta del ejecutable actual.
    """
    import sys

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()

    return Path(__file__).resolve()


def obtener_pid_actual() -> int:
    """
    actualizacion.py > obtener_pid_actual
    Devuelve el PID del proceso actual.
    """
    return os.getpid()


def obtener_release_latest() -> Dict[str, Any]:
    """
    actualizacion.py > obtener_release_latest
    Consulta el último release publicado en GitHub.
    """
    updater_logger.info("Consultando release latest en GitHub...")

    response = requests.get(
        GITHUB_API_LATEST,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "DraftSender-App-Updater",
        },
        timeout=30,
    )

    response.raise_for_status()

    release = response.json()

    if not release.get("tag_name"):
        raise RuntimeError("El release latest no tiene tag_name.")

    if not release.get("assets"):
        raise RuntimeError("El release latest no tiene assets adjuntos.")

    updater_logger.info("Release latest detectado: %s", release.get("tag_name"))

    return release


def seleccionar_asset_exe(release: Dict[str, Any]) -> Dict[str, Any]:
    """
    actualizacion.py > seleccionar_asset_exe
    Selecciona el ejecutable correcto del release.

    Acepta:
    - DraftSender_v1.5.exe
    - DraftSender_v1.5.5.exe

    Rechaza:
    - setup_installer.exe
    - instalador.exe
    """
    assets = release.get("assets", [])
    candidatos = []

    for asset in assets:
        nombre = str(asset.get("name", "")).strip()
        nombre_lower = nombre.lower()

        if not nombre_lower.endswith(".exe"):
            continue

        if not nombre_lower.startswith(APP_EXE_PREFIX.lower()):
            continue

        if "setup" in nombre_lower:
            continue

        if "installer" in nombre_lower:
            continue

        if "instalador" in nombre_lower:
            continue

        candidatos.append(asset)

    if not candidatos:
        encontrados = ", ".join(str(asset.get("name", "")) for asset in assets)

        raise RuntimeError(
            "No se encontró un ejecutable válido de DraftSender en el último release.\n\n"
            "El asset debe comenzar con 'DraftSender' y terminar en '.exe'.\n\n"
            f"Assets encontrados:\n{encontrados or 'ninguno'}"
        )

    asset = candidatos[0]

    updater_logger.info("Asset seleccionado: %s", asset.get("name"))

    return asset


def obtener_info_actualizacion() -> Dict[str, Any]:
    """
    actualizacion.py > obtener_info_actualizacion
    Compara la versión actual embebida en el ejecutable contra el último release.
    """
    version_actual = obtener_version_local()

    release = obtener_release_latest()
    version_remota = normalizar_version(release.get("tag_name", ""))

    asset = seleccionar_asset_exe(release)

    disponible = es_version_mayor(version_remota, version_actual)

    updater_logger.info(
        "Comparación de versiones. Actual=%s, Remota=%s, Disponible=%s",
        version_actual,
        version_remota,
        disponible,
    )

    return {
        "disponible": disponible,
        "version_actual": version_actual,
        "version_remota": version_remota,
        "release": release,
        "asset": asset,
        "asset_name": asset.get("name"),
        "download_url": asset.get("browser_download_url"),
    }


def hay_nueva_version_disponible() -> bool:
    """
    actualizacion.py > hay_nueva_version_disponible
    Retorna True si existe una versión remota mayor que la actual.
    """
    try:
        info = obtener_info_actualizacion()
        return bool(info.get("disponible"))

    except Exception as e:
        updater_logger.exception("Error verificando actualización: %s", e)
        return False


def descargar_asset(asset: Dict[str, Any]) -> Path:
    """
    actualizacion.py > descargar_asset
    Descarga el ejecutable nuevo como archivo temporal .download.
    """
    nombre = str(asset.get("name", "")).strip()
    url = asset.get("browser_download_url")

    if not nombre:
        raise RuntimeError("El asset de actualización no tiene nombre.")

    if not url:
        raise RuntimeError("El asset de actualización no tiene URL de descarga.")

    directorio_app = obtener_directorio_app()
    directorio_app.mkdir(parents=True, exist_ok=True)

    destino_temporal = directorio_app / f"{nombre}.download"

    if destino_temporal.exists():
        destino_temporal.unlink()

    updater_logger.info("Descargando actualización desde: %s", url)
    updater_logger.info("Destino temporal: %s", destino_temporal)

    with requests.get(url, stream=True, timeout=180) as response:
        response.raise_for_status()

        with destino_temporal.open("wb") as archivo:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    archivo.write(chunk)

    if not destino_temporal.exists():
        raise RuntimeError("No se creó el archivo temporal de actualización.")

    size = destino_temporal.stat().st_size

    updater_logger.info("Descarga finalizada. Tamaño: %s bytes", size)

    if size < 1024 * 1024:
        raise RuntimeError(
            f"La descarga parece inválida o incompleta. Tamaño recibido: {size} bytes."
        )

    return destino_temporal


def crear_acceso_directo_escritorio(ruta_exe: Path) -> None:
    """
    actualizacion.py > crear_acceso_directo_escritorio
    Crea o actualiza el acceso directo del escritorio.

    Importante:
    Esta función puede ejecutarse desde un hilo secundario, por eso debe
    inicializar COM con pythoncom.CoInitialize().
    """
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()

        try:
            escritorio = Path(os.path.join(os.environ["USERPROFILE"], "Desktop"))
            acceso_directo = escritorio / SHORTCUT_NAME

            if acceso_directo.exists():
                acceso_directo.unlink()

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(str(acceso_directo))

            shortcut.TargetPath = str(ruta_exe)
            shortcut.WorkingDirectory = str(ruta_exe.parent)
            shortcut.IconLocation = str(ruta_exe)
            shortcut.Description = "DraftSender"
            shortcut.Save()

            updater_logger.info("Acceso directo actualizado: %s", acceso_directo)

        finally:
            pythoncom.CoUninitialize()

    except Exception as e:
        updater_logger.exception("No se pudo actualizar acceso directo: %s", e)


def crear_script_reemplazo(
    exe_temporal: Path,
    exe_final: Path,
    exe_actual: Path,
    pid_actual: int,
) -> Path:
    """
    actualizacion.py > crear_script_reemplazo
    Crea un .bat que espera a que cierre DraftSender, instala la nueva versión
    con su nombre correcto y abre el nuevo ejecutable.
    """
    bat_path = DATA_DIR / "actualizar_draftsender.bat"
    log_path = LOGS_DIR / "updater_bat.log"

    contenido = f"""@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo =============================== >> "{log_path}"
echo Iniciando actualizacion DraftSender >> "{log_path}"
echo PID actual: {pid_actual} >> "{log_path}"
echo EXE temporal: "{exe_temporal}" >> "{log_path}"
echo EXE final: "{exe_final}" >> "{log_path}"
echo EXE actual: "{exe_actual}" >> "{log_path}"

echo Esperando cierre de DraftSender...
echo Esperando cierre de DraftSender... >> "{log_path}"

set /a INTENTOS=0

:waitloop
tasklist /FI "PID eq {pid_actual}" | find "{pid_actual}" > nul
if not errorlevel 1 (
    set /a INTENTOS+=1
    echo Esperando cierre. Intento !INTENTOS! >> "{log_path}"
    timeout /t 1 /nobreak > nul

    if !INTENTOS! GEQ 60 (
        echo ERROR: El proceso no cerro despues de 60 segundos. >> "{log_path}"
        exit /b 10
    )

    goto waitloop
)

echo Proceso anterior cerrado. >> "{log_path}"

if not exist "{exe_temporal}" (
    echo ERROR: No existe el archivo temporal descargado: "{exe_temporal}" >> "{log_path}"
    exit /b 1
)

echo Copiando nueva version... >> "{log_path}"
copy /Y "{exe_temporal}" "{exe_final}" >> "{log_path}" 2>&1

if errorlevel 1 (
    echo ERROR: Fallo copy hacia exe final. >> "{log_path}"
    exit /b 2
)

if not exist "{exe_final}" (
    echo ERROR: No se pudo crear el exe final: "{exe_final}" >> "{log_path}"
    exit /b 3
)

echo Eliminando temporal... >> "{log_path}"
del /F /Q "{exe_temporal}" >> "{log_path}" 2>&1

if exist "{exe_actual}" (
    if /I not "{exe_actual}"=="{exe_final}" (
        echo Eliminando version anterior: "{exe_actual}" >> "{log_path}"
        del /F /Q "{exe_actual}" >> "{log_path}" 2>&1
    )
)

echo Abriendo nueva version: "{exe_final}" >> "{log_path}"
start "" "{exe_final}"

echo Actualizacion completada. >> "{log_path}"

timeout /t 2 /nobreak > nul
del "%~f0" > nul 2>&1
"""

    bat_path.write_text(contenido, encoding="utf-8")

    updater_logger.info("Script de reemplazo creado: %s", bat_path)

    return bat_path


def preparar_actualizacion(info: Dict[str, Any]) -> Path:
    """
    actualizacion.py > preparar_actualizacion
    Descarga la nueva versión y prepara el script de reemplazo.
    """
    asset = info.get("asset")

    if not asset:
        raise RuntimeError("No se recibió información del asset de actualización.")

    exe_temporal = descargar_asset(asset)

    nombre_final = str(asset.get("name", "")).strip()

    if not nombre_final:
        raise RuntimeError("El asset no tiene nombre final válido.")

    exe_final = obtener_directorio_app() / nombre_final
    exe_actual = obtener_ejecutable_actual()
    pid_actual = obtener_pid_actual()

    updater_logger.info("EXE actual: %s", exe_actual)
    updater_logger.info("EXE final: %s", exe_final)

    crear_acceso_directo_escritorio(exe_final)

    return crear_script_reemplazo(
        exe_temporal=exe_temporal,
        exe_final=exe_final,
        exe_actual=exe_actual,
        pid_actual=pid_actual,
    )


def ejecutar_script_actualizacion(bat_path: Path) -> None:
    """
    actualizacion.py > ejecutar_script_actualizacion
    Ejecuta el .bat de actualización.
    """
    updater_logger.info("Ejecutando script de actualización: %s", bat_path)

    subprocess.Popen(
        ["cmd.exe", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )


def ejecutar_actualizacion(forzar: bool = False) -> bool:
    """
    actualizacion.py > ejecutar_actualizacion
    Ejecuta el flujo de actualización.

    Retorna True si preparó y lanzó el actualizador.
    """
    info = obtener_info_actualizacion()

    if not info.get("disponible"):
        updater_logger.info("No hay actualización disponible.")
        return False

    bat_path = preparar_actualizacion(info)
    ejecutar_script_actualizacion(bat_path)

    return True