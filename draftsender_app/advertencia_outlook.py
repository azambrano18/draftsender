import time
from pywinauto import Desktop

def detectar_advertencia_outlook(timeout=20) -> bool:
    """
    Revisa si Outlook lanza una advertencia de seguridad relacionada con accesos MAPI
    (como "un programa está intentando enviar un correo en su nombre").

    Espera hasta `timeout` segundos verificando si aparece una ventana de advertencia.

    Args:
        timeout (int): Tiempo máximo de espera en segundos. Default es 20.

    Returns:
        bool: True si se detecta una advertencia, False si no aparece dentro del tiempo.
    """
    for _ in range(timeout):
        try:
            ventana = Desktop(backend="uia").window(title_re=".*programa está intentando.*", visible_only=True)
            if ventana.exists(timeout=1):
                return True
        except Exception:
            pass
        time.sleep(1)
    return False