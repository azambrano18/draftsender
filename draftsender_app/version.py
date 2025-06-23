import re
import os
from draftsender_app.ui_utils import get_data_path

def obtener_version_local() -> str:
    try:
        data_dir = get_data_path()
        version_file = os.path.join(data_dir, "version.txt")

        if os.path.exists(version_file):
            with open(version_file, "r", encoding="utf-8") as f:
                version = f.read().strip()
                if re.match(r"^v?\d+\.\d+\.\d+$", version):
                    return version

        # Si no existe o es inválido, crear versión por defecto
        os.makedirs(data_dir, exist_ok=True)
        default_version = "1.0.0"
        with open(version_file, "w", encoding="utf-8") as f:
            f.write(default_version)
        return default_version

    except Exception:
        return "versión desconocida"