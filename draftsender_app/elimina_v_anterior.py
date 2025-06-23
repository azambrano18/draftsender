import time
import logging
import os
import re
import sys

logger = logging.getLogger("DraftSender")

def eliminar_ejecutable_anterior_si_es_necesario():
    """
    Elimina automáticamente el ejecutable anterior de la aplicación si se encuentra en la carpeta temporal.

    Esto se usa cuando la app se actualiza a sí misma. Busca un archivo llamado 'DraftSender_OLD.exe'
    y lo elimina silenciosamente si existe.

    Nota: Esta función debe ejecutarse al inicio del programa, antes de cargar la GUI.

    Returns:
        None
    """
    if len(sys.argv) >= 3 and sys.argv[1] == "--delete":
        antiguo = sys.argv[2]
        time.sleep(2)

        nombre_archivo = os.path.basename(antiguo)
        match = re.search(r'draftsender[-_]?([\d.]+[\w\-]*)\.exe', nombre_archivo, re.IGNORECASE)
        version_detectada = match.group(1) if match else "desconocida"

        try:
            os.remove(antiguo)
            logger.info(f"Se eliminó la versión: v{version_detectada}")

        except Exception as e:
            logger.warning(f"No se pudo eliminar el ejecutable anterior ({antiguo}) [versión detectada: {version_detectada}]: {e}")

        # Eliminar otros ejecutables antiguos que no sean el actual ni el ya eliminado
        try:
            carpeta = os.path.dirname(sys.executable)
            actual = os.path.abspath(sys.executable)
            for archivo in os.listdir(carpeta):
                if archivo.endswith(".exe") and archivo.startswith("draftsender"):
                    ruta = os.path.abspath(os.path.join(carpeta, archivo))
                    if ruta not in [actual, os.path.abspath(antiguo)]:
                        try:
                            os.remove(ruta)
                            logger.info(f"Ejecutable antiguo eliminado automáticamente: {ruta}")
                        except Exception as e:
                            logger.warning(f"No se pudo eliminar {ruta}: {e}")
        except Exception as e:
            logger.warning(f"Error al intentar limpiar ejecutables antiguos: {e}")