import logging
import os
import re
import subprocess
import time
from typing import List, Optional

import psutil
import pythoncom
import win32com.client


logger = logging.getLogger("DraftSender")


def obtener_cuentas_activas() -> List[str]:
    """
    outlook_utils.py > obtener_cuentas_activas
    Devuelve las cuentas SMTP activas del perfil Outlook actualmente abierto.
    Inicializa COM para evitar errores cuando se llama desde hilos.
    """
    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        cuentas = []

        for account in namespace.Accounts:
            smtp = getattr(account, "SmtpAddress", "")

            if smtp:
                cuentas.append(str(smtp).strip())

        return cuentas

    except Exception as e:
        logger.exception("No se pudieron obtener cuentas activas de Outlook: %s", e)
        return []

    finally:
        pythoncom.CoUninitialize()


def obtener_carpeta_borradores(namespace, cuenta: str):
    """
    outlook_utils.py > obtener_carpeta_borradores
    Obtiene la carpeta Borradores de la cuenta indicada, usando DeliveryStore como primera opción.
    """
    cuenta_normalizada = (cuenta or "").strip().lower()

    if not cuenta_normalizada:
        return None

    try:
        for account in namespace.Accounts:
            smtp = str(getattr(account, "SmtpAddress", "")).strip().lower()

            if smtp != cuenta_normalizada:
                continue

            try:
                return account.DeliveryStore.GetDefaultFolder(16)
            except Exception:
                pass

            root = account.DeliveryStore.GetRootFolder()

            for folder in root.Folders:
                if str(folder.Name).strip().lower() in {"borradores", "drafts"}:
                    return folder

            for folder in root.Folders:
                if str(folder.Name).strip().lower() == "[gmail]":
                    for nombre in ("Borradores", "Drafts"):
                        try:
                            return folder.Folders[nombre]
                        except Exception:
                            pass

            for folder in root.Folders:
                try:
                    for subfolder in folder.Folders:
                        if str(subfolder.Name).strip().lower() in {"borradores", "drafts"}:
                            return subfolder
                except Exception:
                    continue

    except Exception as e:
        logger.exception("Error buscando carpeta de borradores para %s: %s", cuenta, e)

    logger.warning("No se encontró carpeta de borradores para la cuenta: %s", cuenta)

    return None


def cerrar_outlook() -> None:
    """
    outlook_utils.py > cerrar_outlook
    Cierra Outlook de forma controlada y fuerza cierre solo si no responde.
    """
    for proc in psutil.process_iter(["pid", "name"]):
        nombre = proc.info.get("name") or ""

        if "outlook.exe" not in nombre.lower():
            continue

        try:
            proc.terminate()
            proc.wait(timeout=8)

        except Exception:
            try:
                proc.kill()
            except Exception:
                logger.warning("No se pudo finalizar Outlook PID %s", proc.info.get("pid"))


def iniciar_outlook_con_perfil(perfil: str) -> Optional[str]:
    """
    outlook_utils.py > iniciar_outlook_con_perfil
    Inicia Outlook clásico con el perfil especificado.
    """
    perfil = (perfil or "").strip()

    if not perfil:
        logger.error("No se especificó perfil para iniciar Outlook.")
        return None

    if not re.match(r"^[\w\s\-.@]+$", perfil):
        logger.error("Perfil contiene caracteres no válidos: %r", perfil)
        return None

    rutas = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    ]

    for ruta in rutas:
        if os.path.exists(ruta):
            try:
                subprocess.Popen([ruta, "/profile", perfil], shell=False)
                time.sleep(6)

                logger.info("Outlook iniciado con el perfil: %s", perfil)

                return ruta

            except Exception as e:
                logger.exception("No se pudo iniciar Outlook con el perfil %s: %s", perfil, e)
                return None

    logger.error("No se encontró OUTLOOK.EXE en las rutas conocidas.")

    return None


def obtener_cuenta_outlook(_unused, direccion_objetivo: str):
    """
    outlook_utils.py > obtener_cuenta_outlook
    Busca y retorna el objeto Account de Outlook que coincide con la dirección indicada.
    """
    direccion = (direccion_objetivo or "").strip().lower()

    if not direccion:
        return None

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        for acc in namespace.Accounts:
            if str(getattr(acc, "SmtpAddress", "")).strip().lower() == direccion:
                return acc

        return None

    finally:
        pythoncom.CoUninitialize()


def obtener_perfiles_outlook() -> List[str]:
    """
    outlook_utils.py > obtener_perfiles_outlook
    Obtiene nombres de perfiles Outlook desde el registro de Windows.
    """
    posibles_versiones = ["16.0", "15.0", "14.0"]
    perfiles = []

    for version in posibles_versiones:
        clave = fr"HKEY_CURRENT_USER\Software\Microsoft\Office\{version}\Outlook\Profiles"

        try:
            resultado = subprocess.check_output(
                ["reg", "query", clave],
                text=True,
                encoding="utf-8",
                errors="ignore",
            )

            perfiles = [
                line.strip().split("\\")[-1]
                for line in resultado.splitlines()
                if line.strip()
            ]

            perfiles = [
                perfil
                for perfil in perfiles
                if perfil and not perfil.startswith("HKEY_")
            ]

            if perfiles:
                return perfiles

        except subprocess.CalledProcessError:
            continue

        except Exception as e:
            logger.warning("Error leyendo perfiles Outlook %s: %s", version, e)

    return []