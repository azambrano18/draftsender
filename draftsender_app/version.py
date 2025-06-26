import re
import os
import sys

def obtener_version_local() -> str:
    """
    Devuelve la versión actual leyendo el archivo 'data/version.txt'.
    Si no existe, lo crea con una versión por defecto ("v1.0.0").
    """
    try:
        if hasattr(sys, '_MEIPASS'):
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.abspath(".")

        version_path = os.path.join(base_dir, "data", "version.txt")
        os.makedirs(os.path.dirname(version_path), exist_ok=True)

        if not os.path.exists(version_path):
            with open(version_path, "w", encoding="utf-8") as f:
                f.write("v1.0.0")
            return "v1.0.0"

        with open(version_path, "r", encoding="utf-8") as f:
            version = f.read().strip()
            if re.match(r"^v?\d+\.\d+\.\d+$", version):
                return version
            else:
                return "v1.0.0"

    except Exception:
        return "versión desconocida"