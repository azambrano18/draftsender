import win32com.client
import pythoncom
import logging
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
    url_destino: str = None
):
    """
    Registra un envío o un click en la tabla 'envios_clicks'.
    """
    import psycopg2
    import os
    from datetime import datetime
    from pytz import timezone

    fecha_evento = datetime.now(timezone("America/Santiago"))
    DATABASE_URL = os.getenv("DATABASE_URL")

    print("\n[DEBUG] === registrar_envio() ===")
    print(f"  DATABASE_URL: {DATABASE_URL}")
    print(f"  Evento: {evento}")
    print(f"  Remitente: {remitente}")
    print(f"  Destinatario: {destinatario}")
    print(f"  MetodoEnvio: {metodo_envio}")
    print(f"  Asunto: {asunto}")
    print(f"  Token: {token}")
    print(f"  URL destino: {url_destino}\n")

    try:
        with psycopg2.connect(DATABASE_URL) as conn:
            with conn.cursor() as cur:
                if evento == "envio":
                    cur.execute("""
                        INSERT INTO envios_clicks (
                            remitente,
                            destinatario,
                            metodo_envio,
                            asunto,
                            fecha_envio,
                            token,
                            url_destino
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """, (
                        remitente,
                        destinatario,
                        metodo_envio,
                        asunto,
                        fecha_evento,
                        token,
                        url_destino
                    ))
                    print(f"[✓] Envío registrado correctamente para {destinatario}")

                elif evento == "click":
                    cur.execute("""
                        UPDATE envios_clicks
                        SET
                            clicks_count = COALESCE(clicks_count,0) + 1,
                            last_click_at = %s
                        WHERE token = %s
                    """, (
                        fecha_evento,
                        token
                    ))
                    if cur.rowcount == 0:
                        print(f"[!] No se encontró envío con token {token} para actualizar click.")
                    else:
                        print(f"[✓] Click registrado correctamente para token {token}")

                else:
                    print("[X] Tipo de evento desconocido.")

    except Exception as e:
        import traceback
        print("[X] Error al registrar en 'envios_clicks':", e)
        traceback.print_exc()

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
    Envía todos los borradores pendientes desde la cuenta especificada
    y registra cada envío en la base de datos.

    Args:
        perfil (str): Cuenta de Outlook (correo) utilizada para enviar.

    Returns:
        tuple: (cantidad_enviados, lista_errores, mensaje_final)
    """
    import pythoncom
    import win32com.client
    import re
    from urllib.parse import unquote
    from draftsender_app.envios import registrar_envio

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
            print(f"[ERROR] {mensaje}")
            return 0, [mensaje], mensaje

        bandeja_drafts = namespace.GetDefaultFolder(16)  # olFolderDrafts
        items = bandeja_drafts.Items
        items = [item for item in items if item.Class == 43]

        total = len(items)
        print(f"[DEBUG] Total borradores encontrados: {total}")

        if total == 0:
            mensaje_final = "No hay borradores disponibles para enviar."
            return 0, [], mensaje_final

        for index, item in enumerate(items):
            try:
                print(f"\n[DEBUG] Procesando borrador {index+1}/{total}")
                print(f"  To: {item.To}")
                print(f"  Subject: {item.Subject}")

                item._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_enviadora))

                print("[DEBUG] Llamando item.Send()...")
                item.Send()
                print("[DEBUG] Envío completado.")

                enviados += 1

                metodo_envio = "Envio1"
                try:
                    metodo_envio = item.UserProperties["MetodoEnvio"].Value
                except:
                    if item.Subject and item.Subject.lower().startswith("re:"):
                        metodo_envio = "Reenvio2"

                cuerpo = (item.HTMLBody or "") + (item.Body or "")

                # Buscar URL de tracking en el cuerpo
                match = re.search(r'https://[^\s"<>]+/click\?[^"\s<>]+', cuerpo)
                url_tracking = match.group(0) if match else None

                # Extraer URL destino limpia
                url_destino = None
                if url_tracking:
                    url_match = re.search(r'[&?]url=([^&]+)', url_tracking)
                    if url_match:
                        url_destino = unquote(url_match.group(1))

                # Extraer token
                token = None
                if url_tracking:
                    token_match = re.search(r'token=([\w\d]+)', url_tracking)
                    token = token_match.group(1) if token_match else None

                print("[DEBUG] url_tracking:", url_tracking)
                print("[DEBUG] url_destino:", url_destino)
                print("[DEBUG] token:", token)

                registrar_envio(
                    evento="envio",
                    remitente=cuenta_enviadora.SmtpAddress,
                    destinatario=item.To,
                    metodo_envio=metodo_envio,
                    asunto=item.Subject,
                    token=token,
                    url_destino=url_destino
                )
                print("[DEBUG] Registro completado.\n")

            except Exception as e:
                descripcion = f"Error al enviar borrador {index + 1} ({item.To}): {e}"
                print(f"[ERROR] {descripcion}")
                errores.append(descripcion)

        if errores:
            resumen = "\n".join(errores[:5])
            if len(errores) > 5:
                resumen += f"\n...y {len(errores) - 5} errores más."
            mensaje_final = f"Se enviaron {enviados} de {total} borradores.\n\nErrores:\n{resumen}"
        else:
            mensaje_final = f"Se enviaron correctamente {enviados} de {total} borradores."

        return enviados, errores, mensaje_final

    except Exception as e:
        mensaje_error = f"Error inesperado al intentar enviar borradores: {e}"
        print(f"[ERROR] {mensaje_error}")
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