import os
import re
import io
import mammoth
import logging
import warnings
from tkinter import messagebox
from draftsender_app.validaciones import es_email_valido, validar_columnas_obligatorias
from draftsender_app.archivos import leer_excel_destinatarios

TRACKING_ENABLED = False     # ← Pon en True si quieres reactivar el tracking de links

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

        # Activar el mensaje para que Outlook inserte la firma automáticamente
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
    [DESACTIVADA POR FLAG] Devuelve el HTML sin modificar cuando TRACKING_ENABLED es False.
    Mantiene la firma y compatibilidad con llamadas existentes.
    """
    if not TRACKING_ENABLED:
        return cuerpo_html  # ← NO cambia enlaces; deja el original

    # --- código actual debajo se mantiene intacto para reactivación futura ---
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
            # Opción segura: reemplazar por texto genérico para evitar phishing flags
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
    Crea una respuesta real al último correo enviado al destinatario.

    La búsqueda se realiza en:
    1. La carpeta de enviados predeterminada que Outlook asigna a la cuenta.
    2. Carpetas encontradas dentro de la cuenta con nombres como:
       - Enviados
       - Elementos enviados
       - Sent Mail
       - Sent Items
       - Sent Messages
       - Correo enviado

    Además, obtiene las direcciones SMTP desde la colección Recipients,
    evitando depender exclusivamente de item.To.
    """
    import re
    import unicodedata
    import traceback
    from datetime import datetime

    import pythoncom
    import win32com.client

    com_inicializado = False

    def normalizar_texto(valor: str) -> str:
        """
        Normaliza nombres de carpetas para compararlos sin depender
        de mayúsculas, espacios o tildes.
        """
        texto = str(valor or "").strip().casefold()
        texto = unicodedata.normalize("NFD", texto)

        return "".join(
            caracter
            for caracter in texto
            if unicodedata.category(caracter) != "Mn"
        )

    def extraer_email(valor: str) -> str:
        """
        Extrae una dirección de correo desde textos como:

        Juan Pérez <juan@empresa.cl>
        SMTP:juan@empresa.cl
        juan@empresa.cl
        """
        texto = str(valor or "").strip()

        coincidencia = re.search(
            r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
            texto,
            flags=re.IGNORECASE
        )

        if coincidencia:
            return coincidencia.group(0).strip().casefold()

        return ""

    def obtener_smtp_address_entry(address_entry) -> str:
        """
        Intenta obtener la dirección SMTP real desde un AddressEntry
        de Outlook, incluyendo cuentas Exchange e IMAP.
        """
        if address_entry is None:
            return ""

        candidatos = []

        try:
            candidatos.append(address_entry.Address)
        except Exception:
            pass

        try:
            candidatos.append(
                address_entry.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                )
            )
        except Exception:
            pass

        try:
            tipo = str(address_entry.Type or "").upper()

            if tipo == "EX":
                exchange_user = address_entry.GetExchangeUser()

                if exchange_user:
                    candidatos.append(exchange_user.PrimarySmtpAddress)
        except Exception:
            pass

        try:
            exchange_lista = address_entry.GetExchangeDistributionList()

            if exchange_lista:
                candidatos.append(exchange_lista.PrimarySmtpAddress)
        except Exception:
            pass

        for candidato in candidatos:
            correo = extraer_email(candidato)

            if correo:
                return correo

        return ""

    def obtener_smtp_recipient(recipient) -> str:
        """
        Obtiene el correo SMTP real de un Recipient de Outlook.
        """
        candidatos = []

        try:
            candidatos.append(recipient.Address)
        except Exception:
            pass

        try:
            candidatos.append(recipient.Name)
        except Exception:
            pass

        try:
            smtp_address_entry = obtener_smtp_address_entry(
                recipient.AddressEntry
            )

            if smtp_address_entry:
                candidatos.append(smtp_address_entry)
        except Exception:
            pass

        for candidato in candidatos:
            correo = extraer_email(candidato)

            if correo:
                return correo

        return ""

    def correo_enviado_al_destinatario(item, correo_buscado: str) -> bool:
        """
        Comprueba si el mensaje fue enviado al destinatario indicado.

        Primero inspecciona Recipients, que es más confiable que item.To.
        Luego utiliza item.To como respaldo.
        """
        correo_buscado = correo_buscado.strip().casefold()

        try:
            recipients = item.Recipients

            for indice in range(1, recipients.Count + 1):
                try:
                    recipient = recipients.Item(indice)
                    smtp_recipient = obtener_smtp_recipient(recipient)

                    if smtp_recipient == correo_buscado:
                        return True

                except Exception as error_recipient:
                    logger.debug(
                        f"[{cuenta}] No se pudo inspeccionar un destinatario: "
                        f"{error_recipient}"
                    )

        except Exception as error_recipients:
            logger.debug(
                f"[{cuenta}] No se pudo acceder a Recipients: "
                f"{error_recipients}"
            )

        # Respaldo para cuentas en las que item.To contiene el SMTP.
        try:
            texto_to = str(item.To or "").casefold()

            if correo_buscado in texto_to:
                return True

            correos_to = {
                correo.casefold()
                for correo in re.findall(
                    r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
                    texto_to,
                    flags=re.IGNORECASE
                )
            }

            return correo_buscado in correos_to

        except Exception:
            return False

    def recorrer_carpetas(carpeta):
        """
        Recorre recursivamente todas las carpetas y subcarpetas
        pertenecientes al almacén de la cuenta.
        """
        yield carpeta

        try:
            subcarpetas = carpeta.Folders

            for indice in range(1, subcarpetas.Count + 1):
                try:
                    subcarpeta = subcarpetas.Item(indice)
                    yield from recorrer_carpetas(subcarpeta)

                except Exception as error_subcarpeta:
                    logger.debug(
                        f"[{cuenta}] No se pudo inspeccionar una subcarpeta: "
                        f"{error_subcarpeta}"
                    )

        except Exception:
            return

    def obtener_clave_carpeta(carpeta) -> str:
        """
        Obtiene una clave estable para evitar inspeccionar dos veces
        la misma carpeta.
        """
        try:
            return str(carpeta.EntryID)
        except Exception:
            pass

        try:
            return str(carpeta.FolderPath)
        except Exception:
            pass

        return str(id(carpeta))

    def obtener_ruta_carpeta(carpeta) -> str:
        """
        Obtiene una descripción legible de una carpeta Outlook.
        """
        try:
            return str(carpeta.FolderPath)
        except Exception:
            pass

        try:
            return str(carpeta.Name)
        except Exception:
            return "Carpeta desconocida"

    def obtener_carpetas_enviados(cuenta_outlook) -> list:
        """
        Obtiene todas las carpetas que podrían contener mensajes enviados
        para la cuenta seleccionada.
        """
        nombres_enviados = {
            "enviados",
            "elementos enviados",
            "correo enviado",
            "correos enviados",
            "mensajes enviados",
            "sent",
            "sent mail",
            "sent items",
            "sent messages"
        }

        carpetas_encontradas = []
        claves_agregadas = set()

        def agregar_carpeta(carpeta, origen: str) -> None:
            if carpeta is None:
                return

            clave = obtener_clave_carpeta(carpeta)

            if clave in claves_agregadas:
                return

            claves_agregadas.add(clave)
            carpetas_encontradas.append(carpeta)

            logger.info(
                f"[{cuenta}] Carpeta candidata de enviados "
                f"({origen}): {obtener_ruta_carpeta(carpeta)}"
            )

        # Primera opción: carpeta predeterminada de enviados de esa cuenta.
        try:
            carpeta_predeterminada = (
                cuenta_outlook.DeliveryStore.GetDefaultFolder(5)
            )

            agregar_carpeta(
                carpeta_predeterminada,
                "predeterminada Outlook"
            )

        except Exception as error_default:
            logger.warning(
                f"[{cuenta}] Outlook no entregó la carpeta predeterminada "
                f"de enviados: {error_default}"
            )

        # Segunda opción: recorrer todas las carpetas del almacén.
        try:
            carpeta_raiz = cuenta_outlook.DeliveryStore.GetRootFolder()

            for carpeta in recorrer_carpetas(carpeta_raiz):
                try:
                    nombre_normalizado = normalizar_texto(carpeta.Name)

                    if nombre_normalizado in nombres_enviados:
                        agregar_carpeta(
                            carpeta,
                            f"detectada por nombre: {carpeta.Name}"
                        )

                except Exception as error_nombre:
                    logger.debug(
                        f"[{cuenta}] No se pudo leer el nombre de una carpeta: "
                        f"{error_nombre}"
                    )

        except Exception as error_raiz:
            logger.warning(
                f"[{cuenta}] No se pudo recorrer el almacén de la cuenta: "
                f"{error_raiz}"
            )

        return carpetas_encontradas

    def buscar_ultimo_correo_enviado(
        carpetas: list,
        correo_buscado: str
    ):
        """
        Busca el correo más reciente enviado al destinatario entre todas
        las carpetas candidatas.

        En cada carpeta ordena por SentOn descendente y conserva el mensaje
        más reciente encontrado.
        """
        mejor_correo = None
        mejor_fecha = None

        for carpeta in carpetas:
            ruta_carpeta = obtener_ruta_carpeta(carpeta)

            try:
                items = carpeta.Items
                cantidad_items = items.Count

                logger.info(
                    f"[{cuenta}] Buscando correo para {correo_buscado} en "
                    f"{ruta_carpeta}. Elementos: {cantidad_items}"
                )

                if cantidad_items == 0:
                    continue

                items.Sort("[SentOn]", True)

                # Al estar ordenada, la primera coincidencia de cada carpeta
                # será la más reciente dentro de esa carpeta.
                for indice in range(1, cantidad_items + 1):
                    try:
                        item = items.Item(indice)

                        # 43 corresponde a olMail.
                        if getattr(item, "Class", None) != 43:
                            continue

                        if not correo_enviado_al_destinatario(
                            item,
                            correo_buscado
                        ):
                            continue

                        try:
                            fecha_envio = item.SentOn
                        except Exception:
                            fecha_envio = None

                        if (
                            mejor_correo is None
                            or mejor_fecha is None
                            or (
                                fecha_envio is not None
                                and fecha_envio > mejor_fecha
                            )
                        ):
                            mejor_correo = item
                            mejor_fecha = fecha_envio

                        logger.info(
                            f"[{cuenta}] Correo previo encontrado para "
                            f"{correo_buscado} en {ruta_carpeta}. "
                            f"Fecha: {fecha_envio}"
                        )

                        break

                    except Exception as error_item:
                        logger.debug(
                            f"[{cuenta}] No se pudo inspeccionar un mensaje "
                            f"de {ruta_carpeta}: {error_item}"
                        )

            except Exception as error_carpeta:
                logger.warning(
                    f"[{cuenta}] No se pudo revisar la carpeta "
                    f"{ruta_carpeta}: {error_carpeta}"
                )

        return mejor_correo

    try:
        destinatario_normalizado = destinatario.strip().casefold()

        if not destinatario_normalizado:
            raise ValueError("El destinatario está vacío.")

        pythoncom.CoInitialize()
        com_inicializado = True

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        if perfil_outlook:
            namespace.Logon(
                Profile=perfil_outlook,
                ShowDialog=False,
                NewSession=True
            )

        cuenta_encontrada = None

        for account in namespace.Accounts:
            try:
                smtp_account = str(account.SmtpAddress or "").strip().casefold()

                if smtp_account == cuenta.strip().casefold():
                    cuenta_encontrada = account
                    break

            except Exception:
                continue

        if cuenta_encontrada is None:
            raise Exception(
                f"No se encontró la cuenta de Outlook: {cuenta}"
            )

        carpetas_enviados = obtener_carpetas_enviados(
            cuenta_encontrada
        )

        if not carpetas_enviados:
            raise Exception(
                "No se encontró ninguna carpeta de correos enviados para "
                f"la cuenta {cuenta}."
            )

        correo_anterior = buscar_ultimo_correo_enviado(
            carpetas=carpetas_enviados,
            correo_buscado=destinatario_normalizado
        )

        if correo_anterior is None:
            rutas_revisadas = ", ".join(
                obtener_ruta_carpeta(carpeta)
                for carpeta in carpetas_enviados
            )

            raise Exception(
                f"No se encontró un correo previo enviado a "
                f"{destinatario}. Carpetas revisadas: {rutas_revisadas}"
            )

        # Reply conserva la conversación y los identificadores del hilo.
        reply = correo_anterior.Reply()

        # Intentar asegurar que se utilice la cuenta correcta.
        try:
            reply._oleobj_.Invoke(
                *(64209, 0, 8, 0, cuenta_encontrada)
            )
        except Exception as error_cuenta_envio:
            logger.warning(
                f"[{cuenta}] No se pudo fijar SendUsingAccount en la "
                f"respuesta: {error_cuenta_envio}"
            )

        # Display permite que Outlook inserte la firma configurada.
        reply.Display()

        timestamp_envio = datetime.utcnow().isoformat()

        cuerpo_tracking = reemplazar_links_por_tracking(
            cuerpo_html,
            remitente=cuenta,
            destinatario=destinatario,
            timestamp=timestamp_envio
        )

        # 2 corresponde a olFormatHTML.
        try:
            reply.BodyFormat = 2
        except Exception:
            pass

        cuerpo_existente = reply.HTMLBody or ""

        # Se agrega el cuerpo nuevo antes de la firma y del mensaje anterior.
        reply.HTMLBody = cuerpo_tracking + cuerpo_existente

        # Reply sobre un elemento enviado puede dejar la cuenta propia como
        # destinatario. Por eso se fuerza explícitamente el destinatario real.
        reply.To = destinatario

        try:
            reply.Recipients.ResolveAll()
        except Exception:
            pass

        reply.Save()
        reply.Close(1)

        logger.info(
            f"[{cuenta}] Borrador {metodo_envio} creado correctamente "
            f"para {destinatario}"
        )

        return True

    except Exception as error:
        traceback.print_exc()

        logger.error(
            f"[{cuenta}] Error al crear borrador respuesta para "
            f"{destinatario}: {error}"
        )

        print(
            f"[X] Error al crear borrador respuesta para "
            f"{destinatario}: {error}"
        )

        return False

    finally:
        if com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

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
    df = leer_excel_destinatarios(ruta_excel)

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