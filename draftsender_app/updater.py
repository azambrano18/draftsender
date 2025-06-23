import os
import sys
import time
import shutil
import subprocess

def main():
    if len(sys.argv) != 3:
        print("Uso: updater.py <nuevo_exe> <nueva_version>")
        sys.exit(1)

    nuevo_exe = sys.argv[1]                  # Ej: "DraftSender v13.exe"
    nueva_version = sys.argv[2]              # Ej: "v13"
    version_txt = os.path.join("data", "version.txt")

    # Detectar el nombre del ejecutable actual (que se ejecutó antes de lanzar el updater)
    exe_antiguo = detectar_exe_actual()
    if exe_antiguo and os.path.exists(exe_antiguo):
        print(f"Esperando cierre del archivo actual: {exe_antiguo}")
        time.sleep(2)
        try:
            os.remove(exe_antiguo)
            print(f"Archivo antiguo eliminado: {exe_antiguo}")
        except Exception as e:
            print(f"No se pudo eliminar el ejecutable antiguo: {e}")
    else:
        print("No se detectó ejecutable anterior o ya fue eliminado.")

    # Asegurar que la carpeta data exista
    os.makedirs(os.path.dirname(version_txt), exist_ok=True)

    try:
        with open(version_txt, "w", encoding="utf-8") as f:
            f.write(nueva_version)
        print(f"Versión actualizada en {version_txt}")
    except Exception as e:
        print(f"Error al escribir version.txt: {e}")

    # Lanzar nuevo exe
    try:
        print(f"Ejecutando nueva versión: {nuevo_exe}")
        os.startfile(nuevo_exe)
    except Exception as e:
        print(f"No se pudo ejecutar el nuevo EXE: {e}")

def detectar_exe_actual():
    """
    Intenta determinar el .exe original que lanzó esta instancia del updater.
    """
    try:
        parent_path = sys.executable
        if parent_path.endswith(".exe") and "python" not in os.path.basename(parent_path).lower():
            return parent_path
    except Exception:
        pass
    return None

if __name__ == "__main__":
    main()
