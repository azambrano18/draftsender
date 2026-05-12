import os
import re
import sys
from typing import Tuple


VERSION_DEFAULT = "v1.5"


def obtener_base_app() -> str:
    """
    version.py > obtener_base_app
    Devuelve la carpeta base real de ejecución.

    En modo ejecutable PyInstaller:
        carpeta donde está DraftSender_vX.Y.exe

    En modo desarrollo:
        raíz del proyecto.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))

    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def obtener_data_path() -> str:
    """
    version.py > obtener_data_path
    Devuelve la ruta absoluta de la carpeta data.
    """
    data_path = os.path.join(obtener_base_app(), "data")
    os.makedirs(data_path, exist_ok=True)
    return data_path


def obtener_version_file_path() -> str:
    """
    version.py > obtener_version_file_path
    Devuelve la ruta absoluta del archivo version.txt.
    """
    return os.path.join(obtener_data_path(), "version.txt")


def normalizar_version(version: str) -> str:
    """
    version.py > normalizar_version
    Normaliza versiones tipo:
        1.5     -> v1.5
        v1.5    -> v1.5
        1.5.0   -> v1.5.0
        v1.5.0  -> v1.5.0

    Si no es válida, devuelve VERSION_DEFAULT.
    """
    if not version:
        return VERSION_DEFAULT

    version_limpia = str(version).strip()

    if re.match(r"^\d+\.\d+$", version_limpia):
        return f"v{version_limpia}"

    if re.match(r"^v\d+\.\d+$", version_limpia):
        return version_limpia

    if re.match(r"^\d+\.\d+\.\d+$", version_limpia):
        return f"v{version_limpia}"

    if re.match(r"^v\d+\.\d+\.\d+$", version_limpia):
        return version_limpia

    return VERSION_DEFAULT


def version_a_tupla(version: str) -> Tuple[int, int, int]:
    """
    version.py > version_a_tupla
    Convierte una versión vX.Y o vX.Y.Z a tupla numérica.

    Ejemplos:
        v1.5   -> (1, 5, 0)
        v1.5.1 -> (1, 5, 1)
    """
    version_normalizada = normalizar_version(version)
    numeros = version_normalizada.lstrip("v").split(".")

    if len(numeros) == 2:
        numeros.append("0")

    return int(numeros[0]), int(numeros[1]), int(numeros[2])


def es_version_mayor(version_remota: str, version_local: str) -> bool:
    """
    version.py > es_version_mayor
    Devuelve True si version_remota es mayor que version_local.
    """
    return version_a_tupla(version_remota) > version_a_tupla(version_local)


def obtener_version_local() -> str:
    """
    version.py > obtener_version_local
    Lee la versión local desde data/version.txt.

    Si no existe, crea el archivo con VERSION_DEFAULT.
    """
    version_path = obtener_version_file_path()

    try:
        if not os.path.exists(version_path):
            guardar_version_local(VERSION_DEFAULT)
            return VERSION_DEFAULT

        with open(version_path, "r", encoding="utf-8") as archivo:
            version = archivo.read().strip()

        return normalizar_version(version)

    except Exception:
        return VERSION_DEFAULT


def guardar_version_local(version: str) -> None:
    """
    version.py > guardar_version_local
    Guarda la versión local en data/version.txt.
    """
    version_path = obtener_version_file_path()
    version_normalizada = normalizar_version(version)

    os.makedirs(os.path.dirname(version_path), exist_ok=True)

    with open(version_path, "w", encoding="utf-8") as archivo:
        archivo.write(version_normalizada)