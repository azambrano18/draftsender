import win32com.client
import pythoncom
import logging
import os
import re
from typing import List, Tuple
from urllib.parse import unquote
from draftsender_app.outlook_utils import obtener_carpeta_borradores

from dotenv import load_dotenv
load_dotenv()

logger = logging.getLogger("DraftSender")

def registrar_envio(
    evento: str,
    remitente: str,
    destinatario: str,
    metodo_envio: str = None,
    asunto: str = None,
    token: str = None,
    url_destino: str = None,
) -> bool:
    """
    envios.py > registrar_envio
    Registra envíos/clicks en PostgreSQL. No imprime DATABASE_URL ni credenciales.
    """
    import psycopg2
    from datetime import datetime
    from pytz import timezone

    database_url = os.getenv("DATABASE_URL")

    if not database_url:
        logger.warning("DATABASE_URL no configurado; se omite registro en base de datos.")
        return False

    fecha_evento = datetime.now(timezone("America/Santiago"))

    try:
        with psycopg2.connect(database_url) as conn:
            with conn.cursor() as cur:
                if evento == "envio":
                    cur.execute(
                        """
                        INSERT INTO envios_clicks (
                            remitente,
                            destinatario,
                            metodo_envio,
                            asunto,
                            fecha_envio,
                            token,
                            url_destino
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """,
                        (
                            remitente,
                            destinatario,
                            metodo_envio,
                            asunto,
                            fecha_evento,
                            token,
                            url_destino,
                        ),
                    )

                elif evento == "click":
                    cur.execute(
                        """
                        UPDATE envios_clicks
                        SET
                            clicks_count = COALESCE(clicks_count, 0) + 1,
                            last_click_at = %s,
                            url_destino = COALESCE(url_destino, %s)
                        WHERE token = %s
                        """,
                        (
                            fecha_evento,
                            url_destino,
                            token,
                        ),
                    )

                else:
                    logger.warning("Tipo de evento desconocido: %s", evento)
                    return False

        logger.info("Evento %s registrado para %s", evento, destinatario)
        return True

    except Exception as e:
        logger.exception(
            "Error al registrar evento %s para %s: %s",
            evento,
            destinatario,
            e,
        )
        return False

def obtener_cuenta_outlook(outlook, direccion_objetivo):
    """
Busca una cuenta de Outlook activa que coincida con el nombre proporcionado.

Args:
    outlook: Objeto Outlook COM.
    nombre_cuenta (str): Nombre o dirección de la cuenta buscada.

Returns:
    Outlook account object si se encuentra, o None si no existe.
"""
    direccion_objetivo = direccion_objetivo.strip().lower()
    for acc in outlook.Session.Accounts:
        if acc.SmtpAddress.strip().lower() == direccion_objetivo:
            logger.info(f"Cuenta utilizada para envío: {acc.SmtpAddress}")
            return acc
    logger.warning(f"[ADVERTENCIA] No se encontró cuenta Outlook para: {direccion_objetivo}")
    return None

def contar_borradores(cuenta: str) -> int:
    """
    envios.py > contar_borradores
    Cuenta borradores en la carpeta de la cuenta seleccionada, no en la cuenta por defecto.
    """
    import pywintypes

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        drafts_folder = obtener_carpeta_borradores(namespace, cuenta)

        return int(drafts_folder.Items.Count) if drafts_folder else 0

    except pywintypes.com_error as e:
        logger.error("Error COM al contar borradores para %s: %s", cuenta, e)
        return 0

    except Exception as e:
        logger.exception("Error al contar borradores para %s: %s", cuenta, e)
        return 0

    finally:
        pythoncom.CoUninitialize()

def _extraer_tracking(cuerpo: str) -> tuple[str | None, str | None]:
    match = re.search(r'https://[^\s"<>]+/click\?[^"\s<>]+', cuerpo or "")
    url_tracking = match.group(0) if match else None

    url_destino = None
    token = None

    if url_tracking:
        url_match = re.search(r"[&?]url=([^&]+)", url_tracking)

        if url_match:
            url_destino = unquote(url_match.group(1))

        token_match = re.search(r"token=([\w\d]+)", url_tracking)
        token = token_match.group(1) if token_match else None

    return token, url_destino


def _obtener_metodo_envio(item) -> str:
    try:
        prop = item.UserProperties["MetodoEnvio"]

        if prop and prop.Value:
            return str(prop.Value)

    except Exception:
        pass

    asunto = (getattr(item, "Subject", "") or "").lower()

    if asunto.startswith("re:"):
        return "Reenvio2"

    return "Envio1"

def enviar_borradores(perfil: str) -> Tuple[int, List[str], str]:
    """
    envios.py > enviar_borradores
    Envía todos los borradores pendientes desde la cuenta indicada.
    Usa la carpeta Borradores de esa cuenta específica.
    """
    enviados = 0
    errores: List[str] = []

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        cuenta_enviadora = obtener_cuenta_outlook(outlook, perfil)

        if not cuenta_enviadora:
            mensaje = f"No se encontró la cuenta de Outlook: {perfil}"
            return 0, [mensaje], mensaje

        bandeja_drafts = obtener_carpeta_borradores(namespace, perfil)

        if not bandeja_drafts:
            mensaje = f"No se encontró la carpeta Borradores para la cuenta: {perfil}"
            return 0, [mensaje], mensaje

        items = [
            item
            for item in bandeja_drafts.Items
            if getattr(item, "Class", None) == 43
        ]

        total = len(items)

        if total == 0:
            mensaje_final = "No hay borradores disponibles para enviar."
            return 0, [], mensaje_final

        for index, item in enumerate(items, start=1):
            destinatario = getattr(item, "To", "") or ""
            asunto = getattr(item, "Subject", "") or ""

            try:
                item._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_enviadora))

                metodo_envio = _obtener_metodo_envio(item)

                cuerpo = (
                    (getattr(item, "HTMLBody", "") or "")
                    + (getattr(item, "Body", "") or "")
                )

                token, url_destino = _extraer_tracking(cuerpo)

                item.Send()

                enviados += 1

                registrar_envio(
                    evento="envio",
                    remitente=cuenta_enviadora.SmtpAddress,
                    destinatario=destinatario,
                    metodo_envio=metodo_envio,
                    asunto=asunto,
                    token=token,
                    url_destino=url_destino,
                )

            except Exception as e:
                descripcion = f"Error al enviar borrador {index} ({destinatario}): {e}"
                logger.exception(descripcion)
                errores.append(descripcion)

        if errores:
            resumen = "\n".join(errores[:5])

            if len(errores) > 5:
                resumen += f"\n...y {len(errores) - 5} errores más."

            mensaje_final = (
                f"Se enviaron {enviados} de {total} borradores.\n\n"
                f"Errores:\n{resumen}"
            )

        else:
            mensaje_final = f"Se enviaron correctamente {enviados} de {total} borradores."

        return enviados, errores, mensaje_final

    except Exception as e:
        mensaje_error = f"Error inesperado al intentar enviar borradores: {e}"
        logger.exception(mensaje_error)
        return 0, [mensaje_error], mensaje_error

    finally:
        pythoncom.CoUninitialize()

def enviar_siguiente_borrador(cuenta, app, intervalo, etiqueta_estado, contador):
    """
    Envía el primer borrador disponible y registra el envío en DB.
    """
    import re
    from draftsender_app.envios import registrar_envio

    try:
        namespace = app.GetNamespace("MAPI")
        cuenta_outlook = next(
            (acct for acct in namespace.Accounts if acct.SmtpAddress == cuenta),
            None
        )
        if not cuenta_outlook:
            raise Exception(f"No se encontró la cuenta: {cuenta}")

        carpeta_borradores = cuenta_outlook.DeliveryStore.GetDefaultFolder(16)
        items = carpeta_borradores.Items
        items.Sort("[ReceivedTime]", False)

        for item in items:
            if item.Class == 43:
                destinatario = item.To
                asunto = item.Subject or ""

                metodo_envio = "Envio1"
                try:
                    metodo_envio = item.UserProperties["MetodoEnvio"].Value
                except:
                    if asunto.lower().startswith("re:"):
                        metodo_envio = "Reenvio2"

                cuerpo = (item.HTMLBody or "") + (item.Body or "")

                match = re.search(r'https://[^\s"<>]+/click\?[^"\s<>]+', cuerpo)
                url_destino = match.group(0) if match else None

                token = None
                if url_destino:
                    token_match = re.search(r'token=([\w\d]+)', url_destino)
                    token = token_match.group(1) if token_match else None

                item.Send()
                print(f"[✓] Enviado a {destinatario} ({metodo_envio})")

                print("[DEBUG] Voy a registrar_envio con estos datos:")
                print("  remitente:", cuenta_outlook.SmtpAddress)
                print("  destinatario:", destinatario)
                print("  metodo_envio:", metodo_envio)
                print("  asunto:", asunto)
                print("  token:", token)
                print("  url_destino:", url_destino)

                registrar_envio(
                    evento="envio",
                    remitente=cuenta_outlook.SmtpAddress,
                    destinatario=destinatario,
                    metodo_envio=metodo_envio,
                    asunto=asunto,
                    token=token,
                    url_destino=url_destino
                )

                contador["enviados"] += 1
                contador["restantes"] -= 1

                return True

        return False

    except Exception as e:
        import traceback
        traceback.print_exc()
        if etiqueta_estado:
            etiqueta_estado.config(text=f"Error al enviar desde {cuenta}: {e}")
        return False