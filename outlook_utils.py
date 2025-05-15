import winreg  # Para interactuar con el registro de Windows y obtener información de los perfiles de Outlook
import subprocess  # Para ejecutar comandos del sistema
import os  # Para interactuar con el sistema de archivos
import time  # Para introducir tiempos de espera entre operaciones
import win32com.client  # Para interactuar con Outlook mediante COM
import pythoncom  # Para inicializar el modelo de objetos COM

from logger_utils import configurar_logger  # Función para configurar el logger
logger = configurar_logger("outlook_utils")  # Instancia del logger para este módulo

def obtener_perfiles_outlook():
    """
    Obtiene los perfiles de Outlook configurados en el sistema desde el registro de Windows.
    La función busca en diferentes versiones de Outlook (16.0, 15.0, 14.0) y extrae los nombres de los perfiles
    configurados en el sistema.
    Returns:
        list: Lista de perfiles encontrados (incluye la opción "Seleccione perfil..." al principio).
    """
    perfiles = ["Seleccione perfil..."]
    try:
        office_versions = ["16.0", "15.0", "14.0"]  # Versiones de Outlook que se van a revisar
        for version in office_versions:
            path = fr"Software\\Microsoft\\Office\\{version}\\Outlook\\Profiles"  # Ruta en el registro de Windows
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                    i = 0
                    while True:
                        try:
                            perfil = winreg.EnumKey(key, i)  # Obtiene el nombre de cada perfil
                            perfiles.append(perfil)
                            i += 1
                        except OSError:
                            break
                break  # Si encontró perfiles en una versión, no sigue buscando otras
            except FileNotFoundError:
                continue  # Si no encuentra la clave, continúa con la siguiente versión de Outlook
        if len(perfiles) == 1:
            logger.warning("No se encontraron perfiles de Outlook.")  # Registra una advertencia si no se encuentran perfiles
            perfiles.append("No se encontraron perfiles.")
    except Exception:
        logger.exception("Error al obtener perfiles de Outlook.")  # Registra excepciones si ocurren
        perfiles.append("Error al obtener perfiles")
    return perfiles

def cerrar_outlook():
    """
    Cierra la aplicación de Outlook si está en ejecución.
    Utiliza un comando del sistema para finalizar el proceso de Outlook.
    Returns: None
    """
    try:
        subprocess.run("taskkill /F /IM outlook.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logger.info("Outlook cerrado correctamente.")
    except Exception:
        logger.exception("Error al intentar cerrar Outlook.")  # Registra un error si no se puede cerrar Outlook

def abrir_outlook(self):
    rutas = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
    ]
    for ruta in rutas:
        if os.path.exists(ruta):
            try:
                subprocess.Popen(ruta)
                time.sleep(5)
                return
            except Exception as e:
                logger.error(f"No se pudo abrir Outlook: {e}")
                return

    logger.error("No se encontró Outlook en las rutas conocidas.")

def iniciar_outlook_con_perfil(perfil: str):
    """
    Inicia Outlook con el perfil especificado.
    Llama a la ruta de Outlook y le pasa el perfil seleccionado para iniciar la aplicación.
    Args: perfil (str): Nombre del perfil de Outlook a usar.
    Returns: None
    """
    try:
        ruta_outlook = abrir_outlook()  # Obtiene la ruta de instalación de Outlook
        subprocess.Popen([ruta_outlook, "/profile", perfil])  # Ejecuta Outlook con el perfil
        time.sleep(7)  # Espera a que Outlook se inicie completamente
        logger.info(f"Outlook iniciado con el perfil: {perfil}")
    except Exception:
        logger.exception(f"No se pudo iniciar Outlook con el perfil: {perfil}")  # Registra si hay un error al iniciar Outlook

def obtener_cuentas_activas(max_intentos: int = 10, intervalo: int = 1):
    """
    Obtiene las cuentas activas de Outlook.
    Utiliza la interfaz COM de Outlook para acceder a las cuentas configuradas en el perfil activo.
    Args: max_intentos (int): Número máximo de intentos para obtener las cuentas. intervalo (int): Intervalo de tiempo (en segundos) entre intentos.
    Returns: list: Lista de direcciones SMTP de las cuentas activas en Outlook.
    """
    pythoncom.CoInitialize()  # Inicializa el modelo COM
    try:
        for intento in range(max_intentos):
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")  # Conecta con la aplicación Outlook
                namespace = outlook.GetNamespace("MAPI")
                cuentas = [account.SmtpAddress for account in namespace.Accounts]  # Obtiene las cuentas configuradas
                if cuentas:
                    logger.info(f"Se encontraron cuentas activas: {cuentas}")
                    return cuentas
            except Exception:
                logger.warning("Intento fallido al acceder a las cuentas de Outlook", exc_info=True)  # Registra advertencias por fallos
                time.sleep(intervalo)
        logger.warning("No se encontraron cuentas activas tras múltiples intentos.")  # Si no se encuentran cuentas, lo registra
        return []
    finally:
        pythoncom.CoUninitialize()  # Libera los recursos de COM