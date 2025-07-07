# update_runner.py
import sys
import os
import shutil
import time
import subprocess

def is_file_locked(path):
    """
    Intenta abrir el archivo en modo escritura exclusiva para saber si está bloqueado.
    """
    try:
        fd = os.open(path, os.O_RDWR | os.O_EXCL)
        os.close(fd)
        return False
    except OSError:
        return True

def main():
    if len(sys.argv) != 4:
        print("Uso: update_runner.py <nuevo_exe> <exe_viejo> <nombre_final>")
        sys.exit(1)

    ruta_nuevo = sys.argv[1]
    ruta_viejo = sys.argv[2]
    nombre_final = sys.argv[3]

    carpeta_versiones = os.path.join("data", "versiones")
    os.makedirs(carpeta_versiones, exist_ok=True)

    # Esperar que el proceso anterior libere el archivo
    for i in range(30):
        if not is_file_locked(ruta_viejo):
            break
        print(f"Esperando que el archivo se libere... intento {i+1}")
        time.sleep(1)
    else:
        print("El proceso anterior no se cerró a tiempo.")
        sys.exit(1)

    # Renombrar el ejecutable viejo (para evitar bloqueos)
    renombrado = ruta_viejo + ".old"
    try:
        os.rename(ruta_viejo, renombrado)
    except Exception as e:
        print(f"No se pudo renombrar el ejecutable anterior: {e}")
        sys.exit(1)

    # Mover el renombrado a la carpeta de versiones
    try:
        destino_viejo = os.path.join(carpeta_versiones, os.path.basename(renombrado))
        shutil.move(renombrado, destino_viejo)
    except Exception as e:
        print(f"No se pudo mover el ejecutable anterior a versiones: {e}")
        sys.exit(1)

    # Mover el nuevo ejecutable al path final
    try:
        shutil.move(ruta_nuevo, ruta_viejo)
    except Exception as e:
        print(f"No se pudo mover el nuevo ejecutable al destino final: {e}")
        sys.exit(1)

    # Lanzar el nuevo ejecutable
    try:
        subprocess.Popen([ruta_viejo], close_fds=True)
    except Exception as e:
        print(f"No se pudo lanzar el nuevo ejecutable: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
