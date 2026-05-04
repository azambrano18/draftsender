import os
from dotenv import load_dotenv
import pythoncom
import win32com.client
import psycopg2
from datetime import datetime
from pytz import timezone
import re

# Cargar variable de entorno
load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL")

def registrar_envio_test(
    remitente,
    destinatario,
    metodo_envio,
    asunto,
    cuerpo,
    token,
    url_destino
):
    print("[DEBUG] registrar_envio_test() INVOCADO:")
    print("  remitente:", remitente)
    print("  destinatario:", destinatario)
    print("  metodo_envio:", metodo_envio)
    print("  asunto:", asunto)
    print("  token:", token)
    print("  url_destino:", url_destino)

    fecha_evento = datetime.now(timezone("America/Santiago"))
    try:
        with psycopg2.connect(DATABASE_URL) as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO envios_clicks (
                        remitente,
                        destinatario,
                        metodo_envio,
                        asunto,
                        cuerpo,
                        fecha_envio,
                        token,
                        url_destino
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    remitente,
                    destinatario,
                    metodo_envio,
                    asunto,
                    cuerpo,
                    fecha_evento,
                    token,
                    url_destino
                ))
                print(f"[✓] Registro insertado en envios_clicks para {destinatario}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print("[X] Error al registrar:", e)


def crear_y_enviar_borrador():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Configura estos datos según tu cuenta y destino
    cuenta = "aazs90@gmail.com"
    destinatario = "alcides.zambrano@outlook.com"
    asunto = "Prueba automática de envío"
    cuerpo_html = "<p>Este es un test automático con tracking</p><p><a href='https://miapp.com/click?token=testtoken123'>Ver enlace</a></p>"
    metodo_envio = "Envio1"

    # Busca cuenta
    cuenta_encontrada = next(
        (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == cuenta.lower()), None
    )
    if not cuenta_encontrada:
        print("[X] No se encontró la cuenta:", cuenta)
        return

    # Crear nuevo MailItem
    mensaje = outlook.CreateItem(0)
    mensaje._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_encontrada))

    mensaje.To = destinatario
    mensaje.Subject = asunto
    mensaje.HTMLBody = cuerpo_html

    # Asignar propiedad MetodoEnvio
    mensaje.UserProperties.Add("MetodoEnvio", 1, True)
    mensaje.UserProperties["MetodoEnvio"].Value = metodo_envio

    # Guardar y enviar
    mensaje.Save()
    print("[✓] Borrador creado correctamente")
    mensaje.Send()
    print("[✓] Correo enviado correctamente")

    # Extraer datos para registro
    cuerpo = mensaje.HTMLBody or ""
    match = re.search(r'https://[^\s"<>]+/click\?[^"\s<>]+', cuerpo)
    url_destino = match.group(0) if match else None
    token = None
    if url_destino:
        token_match = re.search(r'token=([\w\d]+)', url_destino)
        token = token_match.group(1) if token_match else None

    # Registrar
    registrar_envio_test(
        remitente=cuenta,
        destinatario=destinatario,
        metodo_envio=metodo_envio,
        asunto=asunto,
        cuerpo=cuerpo,
        token=token,
        url_destino=url_destino
    )

if __name__ == "__main__":
    crear_y_enviar_borrador()