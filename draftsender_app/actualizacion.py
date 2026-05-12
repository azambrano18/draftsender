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
APP_EXE_PREFIX = "DraftSender"
SHORTCUT_NAME = "DraftSender.lnk"


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


def obtener_release_latest() -> Dict[str, Any]:
    """
    actualizacion.py > obtener_release_latest
    Consulta el último release publicado en GitHub.
    """
    response = requests.get(
        GITHUB_API_LATEST,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "DraftSender-App-Updater",
        },
        timeout=25,
    )

    response.raise_for_status()

    release = response.json()

    if not release.get("tag_name"):
        raise RuntimeError("El release latest no tiene tag_name.")

    if not release.get("assets"):
        raise RuntimeError("El release latest no tiene assets adjuntos.")

    return release


def seleccionar_asset_exe(release: Dict[str, Any]) -> Dict[str, Any]:
    """
    actualizacion.py > seleccionar_asset_exe
    Selecciona el ejecutable correcto del release.

    Acepta:
    - DraftSender_v1.5.exe
    - DraftSender_v1.5.1.exe

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

    return candidatos[0]


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
    Función usada por gui.py.
    Retorna True si existe una versión remota mayor que la actual.
    """
    try:
        info = obtener_info_actualizacion()
        return bool(info.get("disponible"))

    except Exception:
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

    with requests.get(url, stream=True, timeout=180) as response:
        response.raise_for_status()

        with destino_temporal.open("wb") as archivo:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    archivo.write(chunk)

    if not destino_temporal.exists() or destino_temporal.stat().st_size == 0:
        raise RuntimeError("La descarga de actualización quedó vacía o incompleta.")

    return destino_temporal


def crear_acceso_directo_escritorio(ruta_exe: Path) -> None:
    """
    actualizacion.py > crear_acceso_directo_escritorio
    Crea o actualiza el acceso directo del escritorio.
    """
    try:
        import win32com.client

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

    except Exception:
        # No bloqueamos la actualización si falla el acceso directo.
        pass


def crear_script_reemplazo(exe_temporal: Path, exe_final: Path, exe_actual: Path) -> Path:
    """
    actualizacion.py > crear_script_reemplazo
    Crea un .bat que reemplaza el ejecutable actual cuando la app se cierre.
    """
    bat_path = Path(tempfile.gettempdir()) / "draftsender_update.bat"

    contenido = f"""@echo off
chcp 65001 > nul
echo Actualizando DraftSender...
timeout /t 2 /nobreak > nul

copy /Y "{exe_temporal}" "{exe_final}" > nul

if exist "{exe_final}" (
    del "{exe_temporal}" > nul 2>&1
)

if exist "{exe_actual}" (
    if /I not "{exe_actual}"=="{exe_final}" (
        del "{exe_actual}" > nul 2>&1
    )
)

start "" "{exe_final}"

del "%~f0" > nul 2>&1
"""

    bat_path.write_text(contenido, encoding="utf-8")

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
    exe_final = obtener_directorio_app() / nombre_final
    exe_actual = obtener_ejecutable_actual()

    crear_acceso_directo_escritorio(exe_final)

    return crear_script_reemplazo(
        exe_temporal=exe_temporal,
        exe_final=exe_final,
        exe_actual=exe_actual,
    )


def ejecutar_script_actualizacion(bat_path: Path) -> None:
    """
    actualizacion.py > ejecutar_script_actualizacion
    Ejecuta el .bat de actualización.
    """
    subprocess.Popen(
        ["cmd.exe", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )


def ejecutar_actualizacion(forzar: bool = False) -> bool:
    """
    actualizacion.py > ejecutar_actualizacion
    Función usada por gui.py.

    Si forzar=True:
        permite ejecutar desde el menú Actualizar.

    Retorna True si preparó la actualización.
    """
    info = obtener_info_actualizacion()

    if not info.get("disponible") and not forzar:
        return False

    if not info.get("disponible") and forzar:
        return False

    bat_path = preparar_actualizacion(info)
    ejecutar_script_actualizacion(bat_path)

    return True