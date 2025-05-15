import os
import pandas as pd
import win32com.client
import mammoth
import re
from tkinter import messagebox
from logger_utils import configurar_logger

logger = configurar_logger("borradores")

def es_email_valido(email: str) -> bool:
    patron = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(patron, email) is not None

def cargar_cuerpo_desde_docx(archivo_docx: str, variables: dict) -> str:
    if not os.path.exists(archivo_docx):
        raise FileNotFoundError(f"El archivo '{archivo_docx}' no existe.")

    with open(archivo_docx, "rb") as docx_file:
        resultado = mammoth.convert_to_html(docx_file)
        cuerpo = resultado.value

    for clave, valor in variables.items():
        cuerpo = cuerpo.replace(f"[{clave}]", str(valor))
        cuerpo = cuerpo.replace(f"{{{{{clave}}}}}", str(valor))  # {{Nombre}}

    etiquetas_no_reemplazadas = re.findall(r"\[\w+\]|\{\{\w+\}\}", cuerpo)
    if etiquetas_no_reemplazadas:
        etiquetas_unicas = list(set(etiquetas_no_reemplazadas))
        logger.warning(f"Etiquetas no reemplazadas detectadas: {etiquetas_unicas}")

    cuerpo = f'<div style="font-family: Calibri, sans-serif; font-size: 11pt;">{cuerpo}</div>'
    return cuerpo

def indice_a_letra_columna(indice):
    letras = ""
    while indice >= 0:
        letras = chr(indice % 26 + 65) + letras
        indice = indice // 26 - 1
    return letras

def crear_borrador(cuenta, destinatario, asunto, cuerpo_html, perfil_outlook=""):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    if perfil_outlook:
        namespace.Logon(Profile=perfil_outlook, ShowDialog=False, NewSession=True)

    cuenta_encontrada = next(
        (acc for acc in namespace.Accounts if acc.SmtpAddress.lower() == cuenta.lower()), None
    )

    if not cuenta_encontrada:
        raise RuntimeError(f"No se encontró la cuenta de Outlook: {cuenta}")

    mensaje = outlook.CreateItem(0)
    mensaje._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_encontrada))
    mensaje.Display()
    firma = mensaje.HTMLBody or ""
    mensaje.Subject = asunto
    mensaje.To = destinatario
    mensaje.BodyFormat = 2
    mensaje.HTMLBody = cuerpo_html + firma
    mensaje.Save()
    mensaje.Close(1)

def generar_borradores(cuenta: str, perfil: str, ruta_excel: str, ruta_docx: str, callback_progreso=None) -> int:
    if not os.path.exists(ruta_excel):
        raise FileNotFoundError("No se encontró el archivo Excel.")
    if not os.path.exists(ruta_docx):
        raise FileNotFoundError("No se encontró el archivo Word.")

    df = pd.read_excel(ruta_excel, sheet_name=0)
    columnas_obligatorias = {"Correo", "Asunto", "Nombre"}
    if not columnas_obligatorias.issubset(df.columns):
        raise ValueError("El Excel debe contener las columnas: Correo, Asunto, Nombre")

    enviados = 0

    for index, fila in df.iterrows():
        try:
            fila_excel = index + 2

            destinatario = str(fila["Correo"]).strip()
            asunto = str(fila["Asunto"]).strip()
            nombre = str(fila["Nombre"]).strip()

            if not destinatario:
                raise ValueError(f"Campo vacío en columna 'Correo', celda A{fila_excel}")
            if not es_email_valido(destinatario):
                raise ValueError(f"Correo inválido en columna 'Correo', celda A{fila_excel}: {destinatario}")
            if not asunto:
                raise ValueError(f"Asunto vacío en celda B{fila_excel}")
            if not nombre:
                raise ValueError(f"Nombre vacío en celda C{fila_excel}")

            variables = {col: str(fila[col]).strip() for col in df.columns}
            cuerpo_html = cargar_cuerpo_desde_docx(ruta_docx, variables)
            crear_borrador(cuenta, destinatario, asunto, cuerpo_html, perfil_outlook=perfil)
            enviados += 1

        except Exception as e:
            mensaje_error = f"Error en fila {fila_excel}: {e}"
            logger.error(mensaje_error, exc_info=True)
            messagebox.showerror("Error al generar borradores", mensaje_error)

        if callback_progreso:
            callback_progreso(index + 1, len(df))

    logger.info(f"Se generaron {enviados} borradores.")
    return enviados