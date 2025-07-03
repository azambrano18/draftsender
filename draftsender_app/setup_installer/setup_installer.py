import os
import requests
import subprocess
import sys
import shutil
import winshell
from win32com.client import Dispatch

INSTALL_DIR = r"C:\DraftSender_app"
DATA_DIR = os.path.join(INSTALL_DIR, "data")
LOGS_DIR = os.path.join(DATA_DIR, "logs")
ENV_FILE_NAME = ".env"
ENV_TEMPLATE = """
# DraftSender configuration
DATABASE_URL=postgresql://tracking_db_e8b1_user:xtjrT1S65w61IC3HfyR5txLYjeXmSkQJ@dpg-d1a8hjh5pdvs73aj7fd0-a.oregon-postgres.render.com:5432/tracking_db_e8b1
"""

GITHUB_API = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"
HEADERS = {"Accept": "application/vnd.github.v3+json"}


def crear_directorios():
    print("[+] Creando directorios...")
    os.makedirs(LOGS_DIR, exist_ok=True)
    print(f"   - {INSTALL_DIR}")
    print(f"   - {DATA_DIR}")
    print(f"   - {LOGS_DIR}")


def crear_env():
    env_path = os.path.join(INSTALL_DIR, ENV_FILE_NAME)
    if not os.path.exists(env_path):
        print("[+] Creando archivo .env de configuración...")
        with open(env_path, "w", encoding="utf-8") as f:
            f.write(ENV_TEMPLATE.strip())
    else:
        print("[=] Archivo .env ya existe, no se sobrescribe.")


def obtener_url_ultimo_exe():
    print("[+] Consultando última versión en GitHub...")
    response = requests.get(GITHUB_API, headers=HEADERS)
    response.raise_for_status()
    data = response.json()
    for asset in data["assets"]:
        if asset["name"].endswith(".exe"):
            print(f"   - Última versión: {asset['name']}")
            return asset["browser_download_url"], asset["name"]
    raise Exception("No se encontró archivo .exe en el último release.")


def descargar_exe(url, nombre):
    exe_path = os.path.join(INSTALL_DIR, nombre)
    print(f"[+] Descargando {nombre}...")
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(exe_path, 'wb') as f:
            shutil.copyfileobj(r.raw, f)
    print("[✓] Descarga completa.")
    return exe_path


def crear_acceso_directo(path_exe):
    print("[+] Creando acceso directo en el escritorio...")
    escritorio = winshell.desktop()
    acceso_path = os.path.join(escritorio, "DraftSender.lnk")
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortcut(acceso_path)
    shortcut.TargetPath = path_exe
    shortcut.WorkingDirectory = INSTALL_DIR
    shortcut.IconLocation = path_exe
    shortcut.Save()
    print("[✓] Acceso directo creado.")


def ejecutar_app(path_exe):
    print("[+] Ejecutando aplicación...")
    subprocess.Popen([path_exe], shell=True)


def autodestruir():
    print("[+] Preparando autodestrucción...")
    bat_path = os.path.join(INSTALL_DIR, "cleanup.bat")
    with open(bat_path, 'w') as f:
        f.write(f"""@echo off
timeout /t 3 > nul
del "{sys.argv[0]}"
del "%~f0"
""")
    subprocess.Popen([bat_path], shell=True)


def main():
    crear_directorios()
    crear_env()
    url_exe, nombre_exe = obtener_url_ultimo_exe()
    path_exe = descargar_exe(url_exe, nombre_exe)
    crear_acceso_directo(path_exe)
    ejecutar_app(path_exe)
    autodestruir()


if __name__ == "__main__":
    main()