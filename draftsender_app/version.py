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


def obtener_base_app() -> str:
    """
    version.py > obtener_base_app
    Devuelve la carpeta base externa de ejecución.

    Se mantiene por compatibilidad con actualizacion.py.

    En modo .exe:
        carpeta donde está DraftSender_vX.X.exe

    En modo desarrollo:
        raíz del proyecto.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))

    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def obtener_data_path() -> str:
    """
    version.py > obtener_data_path
    Devuelve la carpeta data externa.

    Se mantiene por compatibilidad con módulos existentes.
    No se usa para leer la versión del título.
    """
    data_path = os.path.join(obtener_base_app(), "data")
    os.makedirs(data_path, exist_ok=True)
    return data_path


def obtener_version_file_path() -> str:
    """
    version.py > obtener_version_file_path
    Devuelve la ruta externa data/version.txt.

    Se mantiene por compatibilidad, pero obtener_version_local()
    NO usa esta ruta como fuente de verdad.
    """
    return os.path.join(obtener_data_path(), "version.txt")


def obtener_ruta_version_embebida() -> str:
    """
    version.py > obtener_ruta_version_embebida
    Devuelve la ruta del version.txt incluido dentro del .exe por PyInstaller.

    GitHub Actions genera este archivo desde el tag:
        github.ref_name

    PyInstaller lo incluye con:
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

    No lee C:\\DraftSender_app\\data\\version.txt para evitar mostrar
    una versión antigua persistida de instalaciones anteriores.
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


def guardar_version_local(version: str) -> None:
    """
    version.py > guardar_version_local
    Guarda una versión en data/version.txt externo.

    Se mantiene por compatibilidad con código existente, pero no debe usarse
    como fuente del título de la ventana.
    """
    version_normalizada = normalizar_version(version)
    version_path = obtener_version_file_path()

    os.makedirs(os.path.dirname(version_path), exist_ok=True)

    with open(version_path, "w", encoding="utf-8") as archivo:
        archivo.write(version_normalizada)


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