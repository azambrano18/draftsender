import os
import requests
import subprocess
import sys
import shutil
import winshell
from win32com.client import Dispatch

INSTALL_DIR = r"C:\DraftSender_app"
GITHUB_API = "https://api.github.com/repos/azambrano18/draftsender/releases/latest"
HEADERS = {"Accept": "application/vnd.github.v3+json"}

def crear_directorio():
    os.makedirs(INSTALL_DIR, exist_ok=True)

def obtener_url_ultimo_exe():
    response = requests.get(GITHUB_API, headers=HEADERS)
    response.raise_for_status()
    data = response.json()
    for asset in data["assets"]:
        if asset["name"].endswith(".exe"):
            return asset["browser_download_url"], asset["name"]
    raise Exception("No se encontró archivo .exe en el último release.")

def descargar_exe(url, nombre):
    exe_path = os.path.join(INSTALL_DIR, nombre)
    print(f"Descargando {nombre}...")
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(exe_path, 'wb') as f:
            shutil.copyfileobj(r.raw, f)
    return exe_path

def ejecutar_app(path_exe):
    print("Ejecutando aplicación...")
    subprocess.Popen([path_exe], shell=True)

def crear_acceso_directo(path_exe):
    print("Creando acceso directo en el escritorio...")
    escritorio = winshell.desktop()
    acceso_path = os.path.join(escritorio, "DraftSender.lnk")
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortcut(acceso_path)
    shortcut.TargetPath = path_exe
    shortcut.WorkingDirectory = INSTALL_DIR
    shortcut.IconLocation = path_exe
    shortcut.Save()

def autodestruir():
    print("Autodestruyendo instalador...")
    bat_path = os.path.join(INSTALL_DIR, "cleanup.bat")
    with open(bat_path, 'w') as f:
        f.write(f"""@echo off
timeout /t 3 > nul
del "{sys.argv[0]}"
del "%~f0"
""")
    subprocess.Popen([bat_path], shell=True)

def main():
    crear_directorio()
    url_exe, nombre_exe = obtener_url_ultimo_exe()
    path_exe = descargar_exe(url_exe, nombre_exe)
    crear_acceso_directo(path_exe)
    ejecutar_app(path_exe)
    autodestruir()

if __name__ == "__main__":
    main()