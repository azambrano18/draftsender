import win32com.client
import pythoncom
import logging
import psycopg2
import os
import dns.resolver
from datetime import datetime
from pytz import timezone
from draftsender_app.outlook_utils import obtener_carpeta_borradores

logger = logging.getLogger("DraftSender")


def registrar_envio(remitente: str, destinatario: str, url_destino: str, tipo_envio: str, token: str):
    """
    Inserta o actualiza el registro de envío en la tabla 'clicks' evitando duplicados.
    """
    dominio = destinatario.split("@")[-1].lower()
    try:
        servidor = sorted(dns.resolver.resolve(dominio, 'MX'), key=lambda r: r.preference)[0].exchange.to_text().lower()
    except Exception:
        servidor = "desconocido"

    fecha_envio = datetime.now(timezone("America/Santiago"))
    DATABASE_URL = os.getenv("DATABASE_URL")

    try:
        with psycopg2.connect(DATABASE_URL) as conn:
            with conn.cursor() as cur:
                # Verificar si ya existe un registro con el mismo remitente + destinatario + token
                cur.execute("""
                    SELECT id FROM clicks
                    WHERE remitente = %s AND destinatario = %s AND token = %s
                """, (remitente, destinatario, token))
                existe = cur.fetchone()

                if existe:
                    # Actualizar registro existente
                    cur.execute("""
                        UPDATE clicks SET
                            fecha_envio = %s,
                            tipo_envio = %s,
                            dominio_destinatario = %s,
                            servidor_correo = %s,
                            url_destino = %s
                        WHERE id = %s
                    """, (
                        fecha_envio, tipo_envio, dominio, servidor,
                        url_destino, existe[0]
                    ))
                    print(f"[✓] Registro actualizado para: {remitente} → {destinatario}")
                else:
                    # Insertar nuevo registro
                    cur.execute("""
                        INSERT INTO clicks (
                            remitente, fecha_envio, tipo_envio, destinatario,
                            dominio_destinatario, servidor_correo, click_apertura,
                            url_destino, navegador, so, pais,
                            ip_public, token, user_agent
                        ) VALUES (%s, %s, %s, %s, %s, %s, NULL, %s, NULL, NULL, NULL, NULL, %s, NULL)
                    """, (
                        remitente, fecha_envio, tipo_envio, destinatario,
                        dominio, servidor, url_destino, token
                    ))
                    print(f"[✓] Nuevo envío registrado: {remitente} → {destinatario}")

    except Exception as e:
        print(f"[X] Error al registrar envío: {e}")

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

def contar_borradores(cuenta):
    """
    Cuenta cuántos correos hay en la carpeta Borradores de la cuenta seleccionada.

    Args:
        outlook: Objeto Outlook COM.
        cuenta (str): Dirección de la cuenta de correo.

    Returns:
        int: Número de correos en la carpeta de borradores.
    """
    import pywintypes
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        drafts_folder = obtener_carpeta_borradores(namespace, cuenta)
        return drafts_folder.Items.Count if drafts_folder else 0
    except pywintypes.com_error as e:
        logger.error(f"Error al contar borradores para {cuenta}: {e}")
        return 0
    finally:
        pythoncom.CoUninitialize()

def enviar_borradores(perfil: str) -> tuple[int, list[str], str]:
    """
    Envía todos los borradores pendientes desde la cuenta especificada y registra cada envío en la base de datos.

    Args:
        perfil (str): Cuenta de Outlook (correo) utilizada para enviar.

    Returns:
        tuple: (cantidad_enviados, lista_errores, mensaje_final)
    """
    import pythoncom
    import win32com.client
    import re

    enviados = 0
    errores = []

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        cuenta_enviadora = next(
            (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == perfil.lower()), None
        )

        if not cuenta_enviadora:
            mensaje = f"No se encontró la cuenta de Outlook: {perfil}"
            logger.error(mensaje)
            return 0, [mensaje], mensaje

        bandeja_drafts = namespace.GetDefaultFolder(16)  # olFolderDrafts
        items = bandeja_drafts.Items
        items = [item for item in items if item.Class == 43 and perfil.lower() in item.SenderEmailAddress.lower()]

        total = len(items)

        for index, item in enumerate(items):
            try:
                item._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_enviadora))
                item.Send()
                enviados += 1

                # LOG: información básica del correo
                logger.info(f"[{perfil}] Enviado a: {item.To}")
                logger.info(f"[{perfil}] Asunto: {item.Subject}")

                # Intentar leer tanto texto como HTML
                cuerpo = (item.Body or "") + (item.HTMLBody or "")
                logger.debug(f"[{perfil}] Cuerpo combinado: {cuerpo[:500]}")

                # Buscar URL con tracking
                match = re.search(r'https://[^\s"<>]+/click\?[^"\s<>]+', cuerpo)
                url_destino = match.group(0) if match else None

                # Buscar token
                token = None
                if url_destino:
                    token_match = re.search(r'token=([\w\d]+)', url_destino)
                    token = token_match.group(1) if token_match else None

                logger.info(f"[{perfil}] URL destino encontrada: {url_destino}")
                logger.info(f"[{perfil}] Token extraído: {token}")

                # Validación mínima antes de registrar
                if not url_destino or not token:
                    logger.warning(f"[{perfil}] No se pudo extraer URL o token del correo a {item.To}")
                    print(">>> DEBUG URL o TOKEN NO EXTRAÍDOS")
                    print("Asunto:", item.Subject)
                    print("Cuerpo detectado:")
                    print(cuerpo[:500])
                    continue  # Salta al siguiente sin registrar

                # Detectar tipo de envío
                tipo_envio = "seguimiento" if item.Subject.lower().startswith("re:") else "envio1"

                # Registrar en base de datos
                registrar_envio(
                    remitente=cuenta_enviadora.SmtpAddress,
                    destinatario=item.To,
                    url_destino=url_destino,
                    tipo_envio=tipo_envio,
                    token=token
                )

            except Exception as e:
                descripcion = f"Error al enviar borrador {index + 1} ({item.To}): {e}"
                errores.append(descripcion)
                logger.error(f"[{perfil}] {descripcion}", exc_info=True)

        if errores:
            resumen = "\n".join(errores[:5])
            if len(errores) > 5:
                resumen += f"\n...y {len(errores) - 5} errores más."
            mensaje_final = f"Se enviaron {enviados} de {total} borradores.\n\nErrores:\n{resumen}"
        else:
            mensaje_final = f"Se enviaron correctamente {enviados} borradores."

        return enviados, errores, mensaje_final

    except Exception as e:
        mensaje_error = f"Error inesperado al intentar enviar borradores: {e}"
        logger.exception(f"[{perfil}] {mensaje_error}")
        return 0, [mensaje_error], mensaje_error

    finally:
        pythoncom.CoUninitialize()

def enviar_siguiente_borrador(cuenta, app, intervalo, etiqueta_estado, contador):
    """
    Envía el primer borrador disponible desde la carpeta Borradores de la cuenta especificada.

    Args:
        outlook: Objeto Outlook COM.
        cuenta (str): Dirección de correo usada para enviar.

    Returns:
        bool: True si se envió un correo, False si no había borradores.
    """
    try:
        namespace = app.GetNamespace("MAPI")
        cuenta_outlook = next((acct for acct in namespace.Accounts if acct.SmtpAddress == cuenta), None)
        if not cuenta_outlook:
            raise Exception(f"No se encontró la cuenta: {cuenta}")

        carpeta_borradores = cuenta_outlook.DeliveryStore.GetDefaultFolder(16)  # olFolderDrafts
        items = carpeta_borradores.Items
        items.Sort("[ReceivedTime]", False)

        for item in items:
            if item.Class == 43:  # 43 = olMailItem
                destinatario = item.To
                asunto = item.Subject
                item.Send()

                contador["enviados"] += 1
                contador["restantes"] -= 1

                return True

        logger.warning(f"No hay borradores disponibles para enviar desde {cuenta}")
        return False

    except Exception as e:
        logger.error(f"Error al enviar borrador desde cuenta {cuenta}: {e}", exc_info=True)
        if etiqueta_estado:
            etiqueta_estado.config(text=f"Error al enviar desde {cuenta}")
        return False