import re
import os
import sys

def obtener_version_local() -> str:
    """
    Devuelve la versión actual leyendo el archivo 'data/version.txt',
    ya sea en entorno de desarrollo o desde ejecutable PyInstaller.
    """
    try:
        if hasattr(sys, '_MEIPASS'):
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.abspath(".")

        version_path = os.path.join(base_dir, "data", "version.txt")

        if os.path.exists(version_path):
            with open(version_path, "r", encoding="utf-8") as f:
                version = f.read().strip()
                if re.match(r"^v?\d+\.\d+\.\d+$", version):
                    return version

        # Si no existe o es inválido, usar valor por defecto
        return "v"

    except Exception:
        return "versión desconocida"