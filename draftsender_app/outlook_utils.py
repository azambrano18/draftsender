import win32com.client
import subprocess
import psutil
import time
import os
import re
import logging

logger = logging.getLogger("DraftSender")

def obtener_cuentas_activas():
    """
    Devuelve una lista con las direcciones de correo de todas las cuentas activas en Outlook.

    Returns:
        list[str]: Lista de cuentas disponibles.
    """

    outlook = win32com.client.Dispatch("Outlook.Application")
    session = outlook.Session
    cuentas = [account.SmtpAddress for account in session.Accounts]
    return cuentas

def obtener_carpeta_borradores(namespace, cuenta):
    """
    Obtiene la carpeta de borradores (Drafts) correspondiente a una cuenta específica en Outlook.

    Args:
        outlook: Objeto Outlook COM.
        cuenta (str): Dirección de la cuenta a buscar.

    Returns:
        Folder COM object: Carpeta de borradores si se encuentra, o None.
    """
    for account in namespace.Application.Session.Accounts:
        if account.SmtpAddress.strip().lower() == cuenta.strip().lower():
            store = account.DeliveryStore
            root = store.GetRootFolder()

            # Buscar directamente por nombre
            for folder in root.Folders:
                if folder.Name.strip().lower() in ["borradores", "drafts"]:
                    return folder

            # Soporte especial para Gmail
            for folder in root.Folders:
                if folder.Name.strip() == "[Gmail]":
                    try:
                        sub = folder.Folders["Borradores"]
                        return sub
                    except:
                        continue

            # Fallback: búsqueda recursiva
            for folder in root.Folders:
                for subfolder in folder.Folders:
                    if subfolder.Name.strip().lower() in ["borradores", "drafts"]:
                        return subfolder

    print(f"No se encontró carpeta de borradores para la cuenta: {cuenta}")
    return None

def cerrar_outlook():
    """
    Fuerza el cierre de todas las instancias de Outlook activas mediante taskkill.
    """
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'outlook.exe' in proc.info['name'].lower():
            try:
                proc.terminate()
                proc.wait(timeout=5)
            except Exception:
                proc.kill()

def iniciar_outlook_con_perfil(perfil):
    """
    Inicia Outlook usando un perfil específico.

    Args:
        perfil (str): Nombre del perfil de Outlook que se desea abrir.

    Returns:
        bool: True si se lanzó correctamente, False en caso de error.
    """
    if not perfil:
        logger.error("No se especificó perfil para iniciar Outlook.")
        return None

    # Validación de caracteres seguros en nombre de perfil
    if not re.match(r"^[\w\s\-]+$", perfil):
        logger.error(f"Perfil contiene caracteres no válidos: '{perfil}'")
        return None

    rutas = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE"
    ]

    for ruta in rutas:
        if os.path.exists(ruta):
            try:
                subprocess.Popen([ruta, "/profile", perfil])
                time.sleep(5)
                logger.info(f"Outlook iniciado con el perfil: {perfil}")
                return ruta
            except Exception as e:
                logger.error(f"No se pudo iniciar Outlook con el perfil '{perfil}': {e}")
                return None

    logger.error("No se encontró Outlook en las rutas conocidas.")
    return None

import win32com.client

def obtener_cuenta_outlook(_unused, direccion_objetivo):
    """
Busca y retorna el objeto de cuenta de Outlook que coincide con la dirección indicada.

Args:
    outlook: Objeto Outlook COM.
    direccion (str): Dirección de correo buscada.

Returns:
    Account COM object si existe, o None.
"""
    direccion_objetivo = direccion_objetivo.strip().lower()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    for acc in namespace.Accounts:
        if acc.SmtpAddress.strip().lower() == direccion_objetivo:
            return acc
    return None

def obtener_perfiles_outlook():
    """
    Obtiene una lista de nombres de perfiles de Outlook registrados en el sistema.

    Returns:
        list[str]: Lista de perfiles detectados (vía registro de Windows).
    """
    posibles_versiones = ["16.0", "15.0", "14.0"]
    perfiles = []

    for version in posibles_versiones:
        clave = fr'HKEY_CURRENT_USER\Software\Microsoft\Office\{version}\Outlook\Profiles'
        try:
            resultado = subprocess.check_output(
                ["reg", "query", clave],
                text=True, encoding="utf-8", errors="ignore"
            )
            perfiles = [line.strip().split("\\")[-1] for line in resultado.splitlines() if line.strip()]
            if perfiles:
                return perfiles
        except subprocess.CalledProcessError as e:
            logger.warning(f"No se pudo acceder a los perfiles de Outlook para versión {version}: {e}")
            continue

    return []