import os  # Para manejar rutas de archivos y comprobar su existencia
import pandas as pd  # Para trabajar con archivos Excel
import win32com.client  # Para interactuar con Outlook
import mammoth  # Para convertir archivos DOCX a HTML
import re

from logger_utils import configurar_logger  # Para la configuración del logger
logger = configurar_logger("borradores")  # Instancia del logger para este módulo

def es_email_valido(email: str) -> bool:
    patron = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(patron, email) is not None

def cargar_cuerpo_desde_docx(archivo_docx: str, variables: dict) -> str:
    if not os.path.exists(archivo_docx):
        raise FileNotFoundError(f"El archivo '{archivo_docx}' no existe.")

    with open(archivo_docx, "rb") as docx_file:
        resultado = mammoth.convert_to_html(docx_file)
        cuerpo = resultado.value

    # Reemplazar etiquetas
    for clave, valor in variables.items():
        cuerpo = cuerpo.replace(f"[{clave}]", str(valor))
        cuerpo = cuerpo.replace(f"{{{{{clave}}}}}", str(valor))  # {{Nombre}}

    # Buscar etiquetas no reemplazadas
    etiquetas_no_reemplazadas = re.findall(r"\[\w+\]|\{\{\w+\}\}", cuerpo)
    if etiquetas_no_reemplazadas:
        etiquetas_unicas = list(set(etiquetas_no_reemplazadas))
        logger.warning(f"Etiquetas no reemplazadas detectadas: {etiquetas_unicas}")

    # Aplicar estilo
    cuerpo = f'<div style="font-family: Calibri, sans-serif; font-size: 11pt;">{cuerpo}</div>'
    return cuerpo

def crear_borrador(cuenta, destinatario, asunto, cuerpo_html, perfil_outlook=""):
    """
    Crea un borrador de correo en Outlook.
    Args:
        cuenta (str): Dirección de correo desde la cual se enviará el mensaje.
        destinatario (str): Dirección de correo del destinatario.
        asunto (str): Asunto del correo.
        cuerpo_html (str): Cuerpo del correo en formato HTML.
        perfil_outlook (str, opcional): Perfil de Outlook a usar. Si no se proporciona, se usará el perfil por defecto.
    Returns: None
    """
    # Conecta con Outlook mediante COM
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    if perfil_outlook:
        namespace.Logon(Profile=perfil_outlook, ShowDialog=False, NewSession=True)  # Inicia Outlook con el perfil especificado

    # Busca la cuenta de Outlook proporcionada
    cuenta_encontrada = next(
        (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == cuenta.lower()),
        None
    )

    if not cuenta_encontrada:
        raise RuntimeError(f"No se encontró la cuenta de Outlook: {cuenta}")  # Si no se encuentra la cuenta, lanza un error

    # Crea un nuevo mensaje de correo
    mensaje = outlook.CreateItem(0)
    mensaje._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_encontrada))  # Asocia la cuenta encontrada al mensaje
    mensaje.Display()  # Muestra el mensaje en pantalla
    firma = mensaje.HTMLBody or ""  # Obtiene la firma si está presente
    mensaje.Subject = asunto  # Asigna el asunto
    mensaje.To = destinatario  # Asigna el destinatario
    mensaje.BodyFormat = 2  # Establece el formato del cuerpo como HTML
    mensaje.HTMLBody = cuerpo_html + firma  # Asigna el cuerpo HTML junto con la firma
    mensaje.Save()  # Guarda el borrador
    mensaje.Close(1)  # Cierra el mensaje


def generar_borradores(cuenta: str, perfil: str, ruta_excel: str, ruta_docx: str, callback_progreso=None) -> int:
    """
    Genera borradores de correos electrónicos a partir de un archivo Excel y un archivo DOCX.
    Args:
        cuenta (str): Dirección de correo desde la cual se enviarán los mensajes.
        perfil (str): Perfil de Outlook a usar.
        ruta_excel (str): Ruta del archivo Excel con los destinatarios y asuntos.
        ruta_docx (str): Ruta del archivo DOCX que contiene el cuerpo del correo.
    Returns:
        int: Número de borradores generados.
    """
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError("No se encontró el archivo Excel.")  # Verifica que el archivo Excel exista
    if not os.path.exists(ruta_docx):
        raise FileNotFoundError("No se encontró el archivo Word.")  # Verifica que el archivo DOCX exista

    # Carga los datos desde el archivo Excel
    df = pd.read_excel(ruta_excel, sheet_name=0)
    columnas_obligatorias = {"Correo", "Asunto", "Nombre"}  # Las columnas que deben estar presentes en el Excel
    if not columnas_obligatorias.issubset(df.columns):
        raise ValueError("El Excel debe contener las columnas: Correo, Asunto, Nombre")

    enviados = 0  # Contador de borradores generados

    for index, fila in df.iterrows():
        try:
            # Validación de campos obligatorios
            destinatario = str(fila["Correo"]).strip()
            asunto = str(fila["Asunto"]).strip()
            nombre = str(fila["Nombre"]).strip()

            if not destinatario or not asunto or not nombre:
                raise ValueError("Uno de los campos obligatorios (Correo, Asunto, Nombre) está vacío.")

            if not es_email_valido(destinatario):
                raise ValueError(f"Correo inválido: {destinatario}")

            variables = {col: str(fila[col]).strip() for col in df.columns}
            cuerpo_html = cargar_cuerpo_desde_docx(ruta_docx, variables)
            crear_borrador(cuenta, destinatario, asunto, cuerpo_html, perfil_outlook=perfil)
            enviados += 1

        except Exception as e:
            logger.error(f"Error en fila {index + 1}: {e}", exc_info=True)

        if callback_progreso:
            callback_progreso(index + 1, len(df))

    logger.info(f"Se generaron {enviados} borradores.")
    return enviados