# update_runner.py
import sys
import os
import shutil
import time
import subprocess

def main():
    if len(sys.argv) != 5:
        print("Uso: update_runner.py <nuevo_exe> <exe_viejo> <nombre_final> <pid_viejo>")
        sys.exit(1)

    ruta_nuevo = sys.argv[1]
    ruta_viejo = sys.argv[2]
    nombre_final = sys.argv[3]
    pid_viejo = sys.argv[4]

    carpeta_versiones = os.path.join("data", "versiones")
    os.makedirs(carpeta_versiones, exist_ok=True)

    # Intentar matar el proceso viejo
    print(f"Terminando proceso {pid_viejo}...")
    try:
        subprocess.run(["taskkill", "/PID", pid_viejo, "/F"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
    except Exception as e:
        print(f"No se pudo terminar el proceso: {e}")

    # Esperar que el archivo se libere
    for i in range(30):
        if not is_file_locked(ruta_viejo):
            break
        print(f"Esperando que se libere... intento {i+1}")
        time.sleep(1)
    else:
        print("El ejecutable sigue bloqueado.")
        sys.exit(1)

    # Renombrar el viejo
    renombrado = ruta_viejo + ".old"
    try:
        os.rename(ruta_viejo, renombrado)
    except Exception as e:
        print(f"No se pudo renombrar: {e}")
        sys.exit(1)

    # Mover a carpeta de versiones
    destino = os.path.join(carpeta_versiones, os.path.basename(renombrado))
    try:
        shutil.move(renombrado, destino)
    except Exception as e:
        print(f"No se pudo mover el viejo ejecutable: {e}")
        sys.exit(1)

    # Mover nuevo
    try:
        shutil.move(ruta_nuevo, ruta_viejo)
    except Exception as e:
        print(f"No se pudo mover el nuevo ejecutable: {e}")
        sys.exit(1)

    # Lanzar nuevo
    try:
        subprocess.Popen([ruta_viejo], close_fds=True)
    except Exception as e:
        print(f"No se pudo iniciar el nuevo ejecutable: {e}")

def is_file_locked(path):
    try:
        fd = os.open(path, os.O_RDWR | os.O_EXCL)
        os.close(fd)
        return False
    except OSError:
        return True

if __name__ == "__main__":
    main()