import os
import re
import sys
from typing import Tuple


VERSION_DEFAULT = "v0.0.0"


def normalizar_version(version: str) -> str:
    """
    version.py > normalizar_version
    Normaliza versiones tipo:
        1.5     -> v1.5
        v1.5    -> v1.5
        1.5.1   -> v1.5.1
        v1.5.1  -> v1.5.1
    """
    if not version:
        return VERSION_DEFAULT

    version_limpia = str(version).strip().replace("\ufeff", "")

    if re.match(r"^\d+\.\d+$", version_limpia):
        return f"v{version_limpia}"

    if re.match(r"^v\d+\.\d+$", version_limpia):
        return version_limpia

    if re.match(r"^\d+\.\d+\.\d+$", version_limpia):
        return f"v{version_limpia}"

    if re.match(r"^v\d+\.\d+\.\d+$", version_limpia):
        return version_limpia

    return VERSION_DEFAULT


def obtener_ruta_version_embebida() -> str:
    """
    version.py > obtener_ruta_version_embebida
    Devuelve la ruta del version.txt incluido dentro del .exe por PyInstaller.

    En GitHub Actions, este archivo se genera desde el tag:
        github.ref_name

    Luego PyInstaller lo incluye con:
        --add-data "draftsender_app/data/version.txt;data"
    """
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, "data", "version.txt")


def obtener_version_local() -> str:
    """
    version.py > obtener_version_local
    Lee exclusivamente la versión embebida dentro del ejecutable.

    No lee C:\\DraftSender_app\\data\\version.txt para evitar usar una versión antigua
    persistida de una instalación previa.
    """
    ruta_version = obtener_ruta_version_embebida()

    try:
        if not os.path.exists(ruta_version):
            return VERSION_DEFAULT

        with open(ruta_version, "r", encoding="utf-8-sig") as archivo:
            version = archivo.read().strip()

        return normalizar_version(version)

    except Exception:
        return VERSION_DEFAULT


def version_a_tupla(version: str) -> Tuple[int, int, int]:
    """
    version.py > version_a_tupla
    Convierte vX.Y o vX.Y.Z a tupla numérica.

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