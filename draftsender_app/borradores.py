import os
import re
import io
import mammoth
import logging
import warnings
from tkinter import messagebox
from draftsender_app.validaciones import es_email_valido, validar_columnas_obligatorias

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.reader.drawings")

logger = logging.getLogger("DraftSender")

def cargar_cuerpo_desde_docx(archivo_docx: str, variables: dict) -> str:
    """
    Carga el contenido HTML de un archivo Word (.docx) y reemplaza variables dentro del cuerpo.

    Args:
        archivo_docx (str): Ruta al archivo Word.
        variables (dict): Diccionario con claves que serán reemplazadas en el contenido.

    Returns:
        str: Cuerpo del correo en HTML con estilo embebido.
    """
    if not os.path.exists(archivo_docx):
        raise FileNotFoundError(f"El archivo '{archivo_docx}' no existe.")

    with open(archivo_docx, "rb") as f:
        contenido_docx = f.read()
        buffer = io.BytesIO(contenido_docx)
        resultado = mammoth.convert_to_html(buffer)
        cuerpo = resultado.value or ""

        # Reemplazo de variables personalizadas
    for clave, valor in variables.items():
        cuerpo = cuerpo.replace(f"[{clave}]", str(valor))
        cuerpo = cuerpo.replace(f"{{{{{clave}}}}}", str(valor))

        # Detectar etiquetas no reemplazadas
    etiquetas_no_reemplazadas = re.findall(r"\[\w+\]|\{\{\w+\}\}", cuerpo)
    if etiquetas_no_reemplazadas:
        etiquetas_unicas = list(set(etiquetas_no_reemplazadas))
        logger.warning(f"Etiquetas no reemplazadas detectadas: {etiquetas_unicas}")

    # 🔧 Limpieza de espacios excesivos después de "Saludos,"
    cuerpo = re.sub(r"(Saludos,)</p>\s*<p>(&nbsp;|\s)*</p>", r"\1</p>", cuerpo, flags=re.IGNORECASE)

    # 🔧 Eliminación de párrafos vacíos y <br> redundantes
    cuerpo = re.sub(r"<p>(&nbsp;|\s)*</p>", "", cuerpo)
    cuerpo = re.sub(r"(\s*<br\s*/?>\s*){2,}", "<br>", cuerpo)

    # Estilo HTML envolvente
    cuerpo_html = f'<div style="font-family: Calibri, sans-serif; font-size: 11pt;">{cuerpo}</div>'
    return cuerpo_html

def indice_a_letra_columna(indice):
    """
    Convierte un índice numérico de columna (ej. 0, 1, 2...) a su representación en letras tipo Excel (A, B, C...).

    Args:
        indice (int): Índice de columna.

    Returns:
        str: Letra de columna correspondiente.
    """
    letras = ""
    while indice >= 0:
        letras = chr(indice % 26 + 65) + letras
        indice = indice // 26 - 1
    return letras

def crear_borrador(
    cuenta: str,
    destinatario: str,
    asunto: str,
    cuerpo_html: str,
    metodo_envio: str,
    perfil_outlook: str = ""
) -> bool:
    """
    Crea un borrador en Outlook usando el botón 'Nuevo correo electrónico' para que Outlook inserte la firma automáticamente.
    Luego añade el cuerpo HTML con tracking antes de la firma.
    """
    import psutil
    import pythoncom
    from datetime import datetime
    import win32com.client

    def outlook_responde() -> bool:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and "outlook.exe" in proc.info['name'].lower():
                return True
        return False

    try:
        if not outlook_responde():
            logger.warning(f"[{cuenta}] Outlook no está ejecutándose")
            messagebox.showwarning("Outlook no disponible", "Outlook no está abierto o no responde.")
            return False

        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if perfil_outlook:
            namespace.Logon(Profile=perfil_outlook, ShowDialog=False, NewSession=True)

        cuenta_encontrada = next(
            (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == cuenta.lower()), None
        )
        if not cuenta_encontrada:
            logger.error(f"[{cuenta}] No se encontró la cuenta en Outlook")
            messagebox.showerror("Cuenta no encontrada", f"No se encontró la cuenta de Outlook: {cuenta}")
            return False

        mensaje = outlook.CreateItem(0)
        mensaje._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_encontrada))

        # ⚠️ Activar el mensaje para que Outlook inserte la firma automáticamente
        mensaje.Display()  # Esto hace que Outlook inserte la firma configurada

        # Capturar la firma generada automáticamente
        firma = mensaje.HTMLBody or ""

        # Insertar cuerpo con tracking ANTES de la firma
        timestamp_envio = datetime.utcnow().isoformat()
        cuerpo_con_tracking = reemplazar_links_por_tracking(
            cuerpo_html, cuenta, destinatario, timestamp_envio
        )

        mensaje.Subject = asunto
        mensaje.To = destinatario
        mensaje.BodyFormat = 1  # fuerza a no usar RTF
        mensaje.HTMLBody = cuerpo_con_tracking + firma

        mensaje.Save()
        mensaje.Close(1)

        return True

    except Exception as e:
        logger.error(f"[{cuenta}] Error al crear borrador: {e}")
        messagebox.showerror("Error al crear borrador", f"Ocurrió un error:\n{e}")
        return False

def reemplazar_links_por_tracking(cuerpo_html: str, remitente: str, destinatario: str, timestamp: str) -> str:
    """
    Reemplaza todos los links en el HTML por links con tracking y token.
    Valida si el texto visible es una URL y si coincide con el dominio real.
    """
    import re
    import hashlib
    import logging
    from html import unescape
    from urllib.parse import quote, urlparse

    logger = logging.getLogger("DraftSender")

    def generar_token(remitente: str, destinatario: str, url: str, secreto: str = "clave-secreta") -> str:
        base = f"{remitente}-{destinatario}-{url}-{secreto}"
        return hashlib.sha256(base.encode()).hexdigest()

    def extraer_dominio(url: str) -> str:
        try:
            parsed = urlparse(url)
            return parsed.hostname or ""
        except:
            return ""

    def reemplazo(match):
        url_original = match.group(1).strip()
        texto_visible = unescape(match.group(2).strip())

        # Validar si el texto visible es una URL
        es_texto_url = re.match(r"https?://|www\.", texto_visible, re.IGNORECASE)
        dominio_visible = extraer_dominio("https://" + texto_visible) if es_texto_url else ""
        dominio_real = extraer_dominio(url_original)

        # Advertencia si no coinciden
        if es_texto_url and dominio_visible.lower() != dominio_real.lower():
            logger.warning(
                f"Advertencia: el texto visible '{texto_visible}' no coincide con el dominio del enlace real '{dominio_real}'"
            )
            # 🔄 Opción segura: reemplazar por texto genérico para evitar phishing flags
            texto_visible = "Ver más"

        # Crear URL con tracking
        token = generar_token(remitente, destinatario, url_original)
        tracking_url = (
            "https://click-tracker-vszi.onrender.com/click"
            f"?from={quote(remitente)}"
            f"&to={quote(destinatario)}"
            f"&sent={quote(timestamp)}"
            f"&url={quote(url_original)}"
            f"&token={token}"
        )

        return f'<a href="{tracking_url}">{texto_visible}</a>'

    return re.sub(
        r'<a\s+href="(https?://[^"]+)"[^>]*>(.*?)</a>',
        reemplazo,
        cuerpo_html,
        flags=re.IGNORECASE | re.DOTALL
    )

def crear_borrador_respuesta(
    cuenta: str,
    destinatario: str,
    cuerpo_html: str,
    metodo_envio: str,
    perfil_outlook: str = ""
) -> bool:
    """
    Crea una respuesta real (Reply) al último correo enviado al destinatario.
    Outlook inserta automáticamente la firma configurada. El cuerpo se inserta antes.
    """
    import pythoncom
    import win32com.client
    from datetime import datetime

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if perfil_outlook:
            namespace.Logon(Profile=perfil_outlook, ShowDialog=False, NewSession=True)

        cuenta_encontrada = next(
            (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == cuenta.lower()),
            None
        )
        if not cuenta_encontrada:
            raise Exception(f"No se encontró la cuenta: {cuenta}")

        # Buscar último correo enviado a ese destinatario
        sent_folder = cuenta_encontrada.DeliveryStore.GetDefaultFolder(5)  # olFolderSentMail
        items = sent_folder.Items
        items.Sort("[SentOn]", True)

        correo_anterior = None
        for item in items:
            if item.Class == 43 and destinatario.lower() in (item.To or "").lower():
                correo_anterior = item
                break

        if not correo_anterior:
            raise Exception(f"No se encontró correo previo enviado a {destinatario}")

        # Crear una respuesta real (esto abre la ventana y Outlook inserta firma)
        reply = correo_anterior.Reply()
        reply.Display()  # Outlook inserta automáticamente la firma

        # Generar cuerpo con tracking
        timestamp_envio = datetime.utcnow().isoformat()
        cuerpo_tracking = reemplazar_links_por_tracking(
            cuerpo_html, remitente=cuenta, destinatario=destinatario, timestamp=timestamp_envio
        )

        # Insertar el nuevo cuerpo antes del contenido original (y de la firma)
        reply.HTMLBody = cuerpo_tracking + reply.HTMLBody

        # Asegurarse que el destinatario esté bien
        reply.To = destinatario

        # Guardar sin usar UserProperties
        reply.Save()
        reply.Close(1)

        return True

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[X] Error al crear borrador respuesta: {e}")
        return False

def generar_borradores(
    cuenta: str,
    perfil: str,
    ruta_excel: str,
    ruta_docx: str,
    modo_envio: str = "Envíos 1",
    callback_progreso=None
) -> tuple[int, list[str], str]:
    """
    Genera borradores de correos a partir de un Excel y una plantilla Word.

    Args:
        cuenta (str): Cuenta de Outlook desde la que se enviará.
        perfil (str): Perfil de Outlook a usar.
        ruta_excel (str): Ruta al archivo Excel.
        ruta_docx (str): Ruta al archivo Word.
        modo_envio (str): "Envíos 1", "Reenviados 2", "Reenviados 3".
        callback_progreso (func, optional): Función para actualizar progreso.

    Returns:
        tuple: (cantidad_enviados, lista_errores, mensaje_final)
    """
    import os
    import pandas as pd
    import openpyxl

    if not os.path.exists(ruta_excel):
        raise FileNotFoundError("No se encontró el archivo Excel.")
    if not os.path.exists(ruta_docx):
        raise FileNotFoundError("No se encontró el archivo Word.")

    # Leer Excel
    df = pd.read_excel(ruta_excel, sheet_name=1)

    if not validar_columnas_obligatorias(set(df.columns)):
        raise ValueError("El Excel no contiene las columnas obligatorias.")

    enviados = 0
    errores = []

    # Leer variable disponibilidad
    try:
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
        hoja = wb.active
        valor_disponibilidad = hoja.cell(row=2, column=5).value or ""
    except Exception as e:
        logger.warning(f"[{cuenta}] No se pudo leer disponibilidad desde E2: {e}")
        valor_disponibilidad = ""

    # Procesar cada fila
    for index, fila in df.iterrows():
        fila_excel = index + 2
        try:
            destinatario = str(fila["Correo"]).strip()
            asunto = str(fila["Asunto"]).strip()
            nombre = str(fila["Nombre"]).strip()

            if not destinatario:
                raise ValueError(f"Campo vacío en columna 'Correo', fila {fila_excel}")
            if not es_email_valido(destinatario):
                raise ValueError(f"Correo inválido: {destinatario}")
            if not asunto:
                raise ValueError(f"Asunto vacío en fila {fila_excel}")
            if not nombre:
                raise ValueError(f"Nombre vacío en fila {fila_excel}")

            # Preparar variables para la plantilla
            variables = {col: str(fila[col]).strip() for col in df.columns}
            variables["Disponibilidad"] = str(valor_disponibilidad).strip()
            cuerpo_html = cargar_cuerpo_desde_docx(ruta_docx, variables)

            # Determinar metodo de envío
            if modo_envio.lower() in ("reenviados 2", "reenviados 3"):
                metodo_envio = "Reenvio2" if "2" in modo_envio else "Reenvio3"
            else:
                metodo_envio = "Envio1"

            # Crear borrador según el método
            if modo_envio.lower() in ("reenviados 2", "reenviados 3"):
                exito = crear_borrador_respuesta(
                    cuenta=cuenta,
                    destinatario=destinatario,
                    cuerpo_html=cuerpo_html,
                    metodo_envio=metodo_envio,
                    perfil_outlook=perfil
                )
            else:
                exito = crear_borrador(
                    cuenta=cuenta,
                    destinatario=destinatario,
                    asunto=asunto,
                    cuerpo_html=cuerpo_html,
                    metodo_envio=metodo_envio,
                    perfil_outlook=perfil
                )

            if exito:
                enviados += 1
            else:
                mensaje = f"Fila {fila_excel}: No se pudo crear borrador para {destinatario}"
                errores.append(mensaje)
                logger.warning(f"[{cuenta}] {mensaje}")

        except Exception as e:
            mensaje = f"Fila {fila_excel}: {e}"
            errores.append(mensaje)
            logger.error(f"[{cuenta}] {mensaje}")

        if callback_progreso:
            callback_progreso(index + 1, len(df))

    # Mensaje final
    if errores:
        resumen = "\n".join(errores[:5])
        if len(errores) > 5:
            resumen += f"\n...y {len(errores) - 5} errores más."
        mensaje_final = f"Se generaron {enviados} borradores.\n\nErrores detectados:\n{resumen}"
    else:
        mensaje_final = f"Se crearon {enviados} borradores correctamente."

    return enviados, errores, mensaje_final