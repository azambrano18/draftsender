import os
import shutil
import subprocess
import sys
import time


def is_file_locked(path: str) -> bool:
    """
    Devuelve True si el archivo está bloqueado por otro proceso.
    """
    try:
        fd = os.open(path, os.O_RDWR | os.O_EXCL)
        os.close(fd)
        return False
    except OSError:
        return True


def esperar_archivo_liberado(path: str, intentos: int = 30, segundos: int = 1) -> bool:
    """
    Espera hasta que un archivo deje de estar bloqueado.
    """
    for intento in range(1, intentos + 1):
        if not is_file_locked(path):
            return True

        print(f"Esperando que se libere el ejecutable. Intento {intento}/{intentos}")
        time.sleep(segundos)

    return False


def terminar_proceso(pid: str) -> None:
    """
    Intenta terminar un proceso por PID en Windows.
    """
    try:
        subprocess.run(
            ["taskkill", "/PID", pid, "/F"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            text=True,
        )
    except Exception as error:
        print(f"No se pudo terminar el proceso {pid}: {error}")


def reemplazar_ejecutable(
    ruta_nuevo: str,
    ruta_viejo: str,
    nombre_final: str,
    pid_viejo: str,
) -> None:
    """
    Reemplaza el ejecutable viejo por el nuevo y reinicia la app.
    """
    if not os.path.exists(ruta_nuevo):
        raise FileNotFoundError(f"No existe el nuevo ejecutable: {ruta_nuevo}")

    carpeta_base = os.path.dirname(os.path.abspath(ruta_viejo))
    carpeta_versiones = os.path.join(carpeta_base, "data", "versiones")
    os.makedirs(carpeta_versiones, exist_ok=True)

    terminar_proceso(pid_viejo)

    if not esperar_archivo_liberado(ruta_viejo):
        raise RuntimeError("El ejecutable anterior sigue bloqueado.")

    ruta_backup = os.path.join(carpeta_versiones, f"{nombre_final}.old")

    if os.path.exists(ruta_backup):
        os.remove(ruta_backup)

    if os.path.exists(ruta_viejo):
        shutil.move(ruta_viejo, ruta_backup)

    shutil.move(ruta_nuevo, ruta_viejo)

    subprocess.Popen([ruta_viejo], close_fds=True)


def main() -> None:
    """
    Punto de entrada del runner.

    Uso:
        python update_runner.py ruta_nuevo ruta_viejo nombre_final pid_viejo
    """
    if len(sys.argv) != 5:
        print("Uso: update_runner.py <ruta_nuevo> <ruta_viejo> <nombre_final> <pid_viejo>")
        sys.exit(1)

    ruta_nuevo = sys.argv[1]
    ruta_viejo = sys.argv[2]
    nombre_final = sys.argv[3]
    pid_viejo = sys.argv[4]

    try:
        reemplazar_ejecutable(
            ruta_nuevo=ruta_nuevo,
            ruta_viejo=ruta_viejo,
            nombre_final=nombre_final,
            pid_viejo=pid_viejo,
        )
    except Exception as error:
        print(f"Error durante actualización: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()