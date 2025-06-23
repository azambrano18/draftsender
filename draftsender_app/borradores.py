import os
import re
import io
import mammoth
import logging
import warnings
import win32com.client
import pythoncom
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

def crear_borrador(cuenta: str, destinatario: str, asunto: str, cuerpo_html: str, perfil_outlook: str = "") -> bool:
    """
    Crea un borrador de correo en Outlook con firma e inserta enlaces de rastreo personalizados.

    Args:
        cuenta (str): Dirección de correo desde la cual se enviará.
        destinatario (str): Dirección de correo del destinatario.
        asunto (str): Asunto del correo.
        cuerpo_html (str): Cuerpo del mensaje en HTML.
        perfil_outlook (str): Nombre del perfil de Outlook, si es necesario para iniciar sesión.

    Returns:
        bool: True si se creó correctamente, False si hubo error.
    """
    import psutil
    import pythoncom
    from datetime import datetime
    from urllib.parse import quote
    import hashlib
    import re

    def outlook_responde() -> bool:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and "outlook.exe" in proc.info['name'].lower():
                return True
        return False

    def generar_token(remitente: str, destinatario: str, url: str, secreto: str = "clave-secreta") -> str:
        base = f"{remitente}-{destinatario}-{url}-{secreto}"
        return hashlib.sha256(base.encode()).hexdigest()

    def reemplazar_links_por_tracking(cuerpo_html: str, remitente: str, destinatario: str, timestamp: str) -> str:
        def reemplazo(match):
            url_original = match.group(1)
            token = generar_token(remitente, destinatario, url_original)
            tracking_url = (
                "https://tracking-mr-jt9a.onrender.com/click"
                f"?from={quote(remitente)}"
                f"&to={quote(destinatario)}"
                f"&sent={quote(timestamp)}"
                f"&url={quote(url_original)}"
                f"&token={token}"
            )
            return f'href="{tracking_url}"'
        return re.sub(r'href="(https?://[^"]+)"', reemplazo, cuerpo_html)

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

        try:
            mensaje.Display()
            firma = mensaje.HTMLBody or ""
        except Exception as e:
            logger.error(f"[{cuenta}] Error al obtener firma con Display(): {e}")
            messagebox.showerror("Firma no disponible", f"No se pudo obtener la firma desde Outlook para {destinatario}.")
            return False

        timestamp_envio = datetime.utcnow().isoformat()

        # Reemplazar todos los enlaces del cuerpo por enlaces con tracking
        cuerpo_con_tracking = reemplazar_links_por_tracking(cuerpo_html, cuenta, destinatario, timestamp_envio)

        # Armar mensaje
        mensaje.Subject = asunto
        mensaje.To = destinatario
        mensaje.BodyFormat = 2
        mensaje.HTMLBody = cuerpo_con_tracking + firma
        mensaje.Save()
        mensaje.Close(1)

        return True

    except Exception as e:
        logger.error(f"[{cuenta}] Error crítico al crear borrador para {destinatario}: {e}")
        messagebox.showerror("Error al crear borrador", f"Ocurrió un error al crear el borrador:\n{e}")
        return False

def crear_borrador_respuesta(cuenta: str, destinatario: str, cuerpo_html: str, perfil_outlook: str = "") -> bool:
    """
    Crea un borrador como respuesta al último correo enviado al destinatario (modo Seguimiento),
    insertando enlaces de rastreo personalizados.

    Args:
        cuenta (str): Cuenta de envío.
        destinatario (str): Correo destino.
        cuerpo_html (str): Contenido HTML que se agregará al mensaje.
        perfil_outlook (str): Nombre del perfil de Outlook.

    Returns:
        bool: True si se creó exitosamente, False si falló.
    """
    import win32com.client
    import pythoncom
    from datetime import datetime
    from urllib.parse import quote
    import hashlib
    import re

    def generar_token(remitente: str, destinatario: str, url: str, secreto: str = "clave-secreta") -> str:
        base = f"{remitente}-{destinatario}-{url}-{secreto}"
        return hashlib.sha256(base.encode()).hexdigest()

    def reemplazar_links_por_tracking(cuerpo_html: str, remitente: str, destinatario: str, timestamp: str) -> str:
        def reemplazo(match):
            url_original = match.group(1)
            token = generar_token(remitente, destinatario, url_original)
            tracking_url = (
                "https://tracking-mr-jt9a.onrender.com/click"
                f"?from={quote(remitente)}"
                f"&to={quote(destinatario)}"
                f"&sent={quote(timestamp)}"
                f"&url={quote(url_original)}"
                f"&token={token}"
            )
            return f'href="{tracking_url}"'
        return re.sub(r'href="(https?://[^"]+)"', reemplazo, cuerpo_html)

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if perfil_outlook:
            namespace.Logon(Profile=perfil_outlook, ShowDialog=False, NewSession=True)

        cuenta_encontrada = next(
            (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == cuenta.lower()), None
        )
        if not cuenta_encontrada:
            logger.error(f"[{cuenta}] Cuenta no encontrada al generar respuesta")
            return False

        store = cuenta_encontrada.DeliveryStore
        root_folder = namespace.Folders[store.DisplayName]

        sent_items = None
        for folder in root_folder.Folders:
            if folder.Name.lower() in ["sent items", "elementos enviados"]:
                sent_items = folder
                break
        if not sent_items:
            for folder in root_folder.Folders:
                if "[gmail]" in folder.Name.lower():
                    for subfolder in folder.Folders:
                        if subfolder.Name.lower() == "enviados":
                            sent_items = subfolder
                            break
        if not sent_items:
            logger.error(f"[{cuenta}] No se encontró la carpeta de enviados")
            return False

        items = sent_items.Items
        items.Sort("[SentOn]", True)

        correo_anterior = None
        for item in items:
            if item.Class == 43 and destinatario.lower() in item.To.lower():
                correo_anterior = item
                break

        if not correo_anterior:
            logger.warning(f"[{cuenta}] No se encontró correo previo a {destinatario}")
            return False

        # Obtener timestamp actual para tracking
        timestamp_envio = datetime.utcnow().isoformat()

        # Reemplazar enlaces por links con tracking
        cuerpo_con_tracking = reemplazar_links_por_tracking(cuerpo_html, cuenta, destinatario, timestamp_envio)

        respuesta = correo_anterior.Reply()
        respuesta.HTMLBody = cuerpo_con_tracking + "<br><br>" + respuesta.HTMLBody
        respuesta.To = destinatario
        respuesta._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_encontrada))
        respuesta.Save()
        respuesta.Close(1)

        logger.info(f"[{cuenta}] Borrador de seguimiento creado para {destinatario}")
        return True

    except Exception as e:
        logger.error(f"[{cuenta}] Error al crear respuesta para {destinatario}: {e}")
        return False

def generar_borradores(cuenta: str, perfil: str, ruta_excel: str, ruta_docx: str, modo_envio: str = "Envíos 1", callback_progreso=None) -> tuple[int, list[str], str]:
    """
    Genera borradores de correos a partir de un Excel y una plantilla Word.

    Args:
        cuenta (str): Cuenta de Outlook desde la que se enviará.
        perfil (str): Perfil de Outlook a usar.
        ruta_excel (str): Ruta al archivo Excel.
        ruta_docx (str): Ruta al archivo Word.
        modo_envio (str): "Envíos 1" o "Seguimiento".
        callback_progreso (func, optional): Función para actualizar progreso.

    Returns:
        tuple: (cantidad_enviados, lista_errores, mensaje_final)
    """
    import pandas as pd
    import openpyxl

    if not os.path.exists(ruta_excel):
        raise FileNotFoundError("No se encontró el archivo Excel.")
    if not os.path.exists(ruta_docx):
        raise FileNotFoundError("No se encontró el archivo Word.")

    df = pd.read_excel(ruta_excel, sheet_name=1)

    if not validar_columnas_obligatorias(set(df.columns)):
        raise ValueError("El Excel no contiene las columnas obligatorias.")

    enviados = 0
    errores = []

    try:
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
        hoja = wb.active
        valor_disponibilidad = hoja.cell(row=2, column=5).value or ""
    except Exception as e:
        logger.warning(f"[{cuenta}] No se pudo leer disponibilidad desde E2: {e}")
        valor_disponibilidad = ""

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

            variables = {col: str(fila[col]).strip() for col in df.columns}
            variables["Disponibilidad"] = str(valor_disponibilidad).strip()
            cuerpo_html = cargar_cuerpo_desde_docx(ruta_docx, variables)

            if modo_envio == "Seguimiento":
                exito = crear_borrador_respuesta(cuenta, destinatario, cuerpo_html, perfil_outlook=perfil)
            else:
                exito = crear_borrador(cuenta, destinatario, asunto, cuerpo_html, perfil_outlook=perfil)

            if exito:
                enviados += 1
            else:
                mensaje = f"Fila {fila_excel}: no se pudo crear borrador para {destinatario}"
                errores.append(mensaje)
                logger.warning(f"[{cuenta}] {mensaje}")

        except Exception as e:
            mensaje = f"Fila {fila_excel}: {e}"
            errores.append(mensaje)
            logger.error(f"[{cuenta}] {mensaje}")

        if callback_progreso:
            callback_progreso(index + 1, len(df))

    if errores:
        resumen = "\n".join(errores[:5])
        if len(errores) > 5:
            resumen += f"\n...y {len(errores) - 5} errores más."
        mensaje_final = f"Se generaron {enviados} borradores.\n\nErrores detectados:\n{resumen}"
    else:
        mensaje_final = f"Se crearon {enviados} borradores correctamente."

    return enviados, errores, mensaje_final

# pr