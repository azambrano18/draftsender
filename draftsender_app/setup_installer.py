import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from tkinter import Tk, messagebox

import requests


INSTALL_DIR = Path(r"C:\DraftSender_app")
DATA_DIR = INSTALL_DIR / "data"
LOGS_DIR = DATA_DIR / "logs"
ENV_FILE = DATA_DIR / ".env"

GITHUB_API = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"

APP_EXE_PREFIX = "DraftSender"
SHORTCUT_NAME = "DraftSender.lnk"

ENV_TEMPLATE = """# DraftSender configuration
DATABASE_URL=postgresql://tracking_db_e8b1_user:xtjrT1S65w61IC3HfyR5txLYjeXmSkQJ@dpg-d1a8hjh5pdvs73aj7fd0-a.oregon-postgres.render.com:5432/tracking_db_e8b1
"""


def crear_directorios() -> None:
    """
    setup_installer.py > crear_directorios
    Crea la estructura base de instalación.
    """
    INSTALL_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)


def crear_env_si_no_existe() -> None:
    """
    setup_installer.py > crear_env_si_no_existe
    Crea el archivo .env local si no existe.
    """
    if ENV_FILE.exists():
        return

    ENV_FILE.write_text(ENV_TEMPLATE, encoding="utf-8")


def obtener_release_latest() -> dict:
    """
    setup_installer.py > obtener_release_latest
    Consulta el último release publicado en GitHub.
    """
    response = requests.get(
        GITHUB_API,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "DraftSender-Installer",
        },
        timeout=30,
    )

    response.raise_for_status()
    release = response.json()

    if not release.get("assets"):
        raise RuntimeError(
            "El último release no tiene archivos adjuntos.\n\n"
            "Debes subir el ejecutable de la app como asset, por ejemplo:\n"
            "DraftSender_v1.5.exe"
        )

    return release


def seleccionar_asset_exe(release: dict) -> dict:
    """
    setup_installer.py > seleccionar_asset_exe
    Selecciona el .exe correcto dentro del último release.
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
        nombres = ", ".join(str(asset.get("name", "")) for asset in assets)

        raise RuntimeError(
            "No se encontró un ejecutable válido de DraftSender en el último release.\n\n"
            "El archivo debe comenzar con 'DraftSender' y terminar en '.exe'.\n\n"
            "Ejemplo válido:\n"
            "DraftSender_v1.5.exe\n\n"
            f"Assets encontrados:\n{nombres or 'ninguno'}"
        )

    return candidatos[0]


def descargar_exe(asset: dict) -> Path:
    """
    setup_installer.py > descargar_exe
    Descarga el .exe del último release a C:\\DraftSender_app.
    """
    nombre = str(asset.get("name", "")).strip()
    url = asset.get("browser_download_url")

    if not nombre:
        raise RuntimeError("El asset seleccionado no tiene nombre.")

    if not url:
        raise RuntimeError("El asset seleccionado no tiene URL de descarga.")

    destino_temporal = INSTALL_DIR / f"{nombre}.download"

    if destino_temporal.exists():
        destino_temporal.unlink()

    with requests.get(url, stream=True, timeout=180) as response:
        response.raise_for_status()

        with destino_temporal.open("wb") as archivo:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    archivo.write(chunk)

    if not destino_temporal.exists() or destino_temporal.stat().st_size == 0:
        raise RuntimeError("La descarga del ejecutable quedó vacía o incompleta.")

    return destino_temporal


def cerrar_draftsender_si_esta_abierto() -> None:
    """
    setup_installer.py > cerrar_draftsender_si_esta_abierto
    Intenta cerrar DraftSender antes de reemplazar la versión.
    """
    try:
        import psutil
    except Exception:
        return

    pid_actual = os.getpid()

    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            pid = proc.info.get("pid")
            nombre = (proc.info.get("name") or "").lower()
            exe = (proc.info.get("exe") or "").lower()

            if pid == pid_actual:
                continue

            es_draftsender = "draftsender" in nombre or "draftsender" in exe
            es_instalador = (
                "setup" in nombre
                or "installer" in nombre
                or "instalador" in nombre
            )

            if es_draftsender and not es_instalador:
                proc.terminate()
                proc.wait(timeout=8)

        except Exception:
            continue


def eliminar_versiones_anteriores(exe_temporal: Path) -> None:
    """
    setup_installer.py > eliminar_versiones_anteriores
    Elimina versiones antiguas de DraftSender en C:\\DraftSender_app.
    """
    for archivo in INSTALL_DIR.glob("*.exe"):
        try:
            nombre = archivo.name.lower()

            if not nombre.startswith("draftsender"):
                continue

            if "setup" in nombre or "installer" in nombre or "instalador" in nombre:
                continue

            if archivo.resolve() != exe_temporal.resolve():
                archivo.unlink()

        except PermissionError:
            raise RuntimeError(
                "No se pudo eliminar la versión anterior porque DraftSender está abierto.\n\n"
                "Cierra DraftSender y vuelve a ejecutar el instalador."
            )

        except Exception:
            continue


def instalar_exe_descargado(exe_temporal: Path) -> Path:
    """
    setup_installer.py > instalar_exe_descargado
    Convierte el archivo .download en el .exe final.
    """
    nombre_final = exe_temporal.name.replace(".download", "")
    destino_final = INSTALL_DIR / nombre_final

    if destino_final.exists():
        try:
            destino_final.unlink()
        except PermissionError:
            raise RuntimeError(
                "No se pudo reemplazar el ejecutable actual porque está en uso.\n\n"
                "Cierra DraftSender y vuelve a ejecutar el instalador."
            )

    exe_temporal.replace(destino_final)

    return destino_final


def crear_acceso_directo(ruta_exe: Path) -> None:
    """
    setup_installer.py > crear_acceso_directo
    Crea acceso directo en el escritorio.
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
        shortcut.WorkingDirectory = str(INSTALL_DIR)
        shortcut.IconLocation = str(ruta_exe)
        shortcut.Description = "DraftSender"
        shortcut.Save()

    except Exception as e:
        raise RuntimeError(
            "DraftSender fue instalado, pero no se pudo crear el acceso directo.\n\n"
            f"Detalle técnico:\n{e}"
        )


def instalar_o_actualizar() -> Path:
    """
    setup_installer.py > instalar_o_actualizar
    Flujo completo del instalador.
    """
    crear_directorios()
    crear_env_si_no_existe()

    release = obtener_release_latest()
    asset = seleccionar_asset_exe(release)

    exe_temporal = descargar_exe(asset)

    cerrar_draftsender_si_esta_abierto()
    eliminar_versiones_anteriores(exe_temporal)

    ruta_instalada = instalar_exe_descargado(exe_temporal)

    crear_acceso_directo(ruta_instalada)

    return ruta_instalada


def abrir_app(ruta_exe: Path) -> None:
    """
    setup_installer.py > abrir_app
    Abre DraftSender después de instalar.
    """
    subprocess.Popen(
        [str(ruta_exe)],
        cwd=str(INSTALL_DIR),
        shell=False,
    )


def main() -> None:
    """
    setup_installer.py > main
    Punto de entrada del instalador.
    """
    root = Tk()
    root.withdraw()

    try:
        ruta_instalada = instalar_o_actualizar()

        abrir = messagebox.askyesno(
            "DraftSender instalado",
            "DraftSender fue instalado o actualizado correctamente.\n\n"
            f"Aplicación:\n{ruta_instalada}\n\n"
            f"Configuración:\n{ENV_FILE}\n\n"
            "Se creó el acceso directo en el escritorio.\n\n"
            "¿Quieres abrir DraftSender ahora?",
        )

        if abrir:
            abrir_app(ruta_instalada)

    except Exception as e:
        messagebox.showerror(
            "Error de instalación",
            f"No se pudo instalar o actualizar DraftSender.\n\n{e}",
        )
        raise

    finally:
        root.destroy()


if __name__ == "__main__":
    main()