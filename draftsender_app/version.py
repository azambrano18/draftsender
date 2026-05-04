import os
import re
import sys
from typing import Tuple


VERSION_DEFAULT = "v1.0.0"


def obtener_base_app() -> str:
    """
    Devuelve la carpeta base real de la aplicación.

    En modo ejecutable PyInstaller:
        C:\\ruta\\donde\\esta\\DraftSender_vX.Y.Z.exe

    En modo desarrollo:
        raíz del proyecto, es decir, la carpeta que contiene draftsender_app.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))

    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def obtener_data_path() -> str:
    """
    Devuelve la ruta absoluta de la carpeta data de la aplicación.
    """
    data_path = os.path.join(obtener_base_app(), "data")
    os.makedirs(data_path, exist_ok=True)
    return data_path


def obtener_version_file_path() -> str:
    """
    Devuelve la ruta absoluta del archivo version.txt.
    """
    return os.path.join(obtener_data_path(), "version.txt")


def normalizar_version(version: str) -> str:
    """
    Normaliza una versión al formato vX.Y.Z.

    Ejemplos:
        1.2.0  -> v1.2.0
        v1.2.0 -> v1.2.0

    Si el valor no es válido, devuelve VERSION_DEFAULT.
    """
    if not version:
        return VERSION_DEFAULT

    version_limpia = str(version).strip()

    if re.match(r"^\d+\.\d+\.\d+$", version_limpia):
        return f"v{version_limpia}"

    if re.match(r"^v\d+\.\d+\.\d+$", version_limpia):
        return version_limpia

    return VERSION_DEFAULT


def version_a_tupla(version: str) -> Tuple[int, int, int]:
    """
    Convierte una versión vX.Y.Z a tupla numérica.

    Ejemplo:
        v1.2.10 -> (1, 2, 10)
    """
    version_normalizada = normalizar_version(version)
    numeros = version_normalizada.lstrip("v").split(".")
    return int(numeros[0]), int(numeros[1]), int(numeros[2])


def es_version_mayor(version_remota: str, version_local: str) -> bool:
    """
    Devuelve True si version_remota es mayor que version_local.
    """
    return version_a_tupla(version_remota) > version_a_tupla(version_local)


def obtener_version_local() -> str:
    """
    Lee la versión local desde data/version.txt.

    Si el archivo no existe, lo crea con VERSION_DEFAULT.
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
    Guarda la versión local en data/version.txt.
    """
    version_path = obtener_version_file_path()
    version_normalizada = normalizar_version(version)

    os.makedirs(os.path.dirname(version_path), exist_ok=True)

    with open(version_path, "w", encoding="utf-8") as archivo:
        archivo.write(version_normalizada)