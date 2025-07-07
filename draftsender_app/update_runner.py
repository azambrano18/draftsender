import sys
import os
import shutil
import time
import subprocess

def main():
    if len(sys.argv) != 4:
        print("Uso: update_runner.py <nuevo_exe> <exe_viejo> <nombre_final>")
        sys.exit(1)

    ruta_nuevo = sys.argv[1]
    ruta_viejo = sys.argv[2]
    nombre_final = sys.argv[3]

    carpeta_versiones = os.path.join("data", "versiones")
    os.makedirs(carpeta_versiones, exist_ok=True)

    # Espera a que el proceso viejo libere el archivo
    for i in range(10):
        if not is_file_locked(ruta_viejo):
            break
        time.sleep(1)
    else:
        print("El proceso anterior no se cerró a tiempo.")
        sys.exit(1)

    # Mover viejo a versiones
    nombre_viejo = os.path.basename(ruta_viejo)
    destino_viejo = os.path.join(carpeta_versiones, nombre_viejo)
    try:
        shutil.move(ruta_viejo, destino_viejo)
    except Exception as e:
        print(f"No se pudo mover el ejecutable anterior: {e}")

    # Mover nuevo al destino
    try:
        shutil.move(ruta_nuevo, ruta_viejo)
    except Exception as e:
        print(f"No se pudo mover el nuevo ejecutable: {e}")
        sys.exit(1)

    # Lanzar nuevo ejecutable
    subprocess.Popen([ruta_viejo], close_fds=True)

def is_file_locked(path):
    """Intenta abrir el archivo en modo escritura exclusiva para saber si está bloqueado."""
    try:
        fd = os.open(path, os.O_RDWR | os.O_EXCL)
        os.close(fd)
        return False
    except OSError:
        return True

if __name__ == "__main__":
    main()
