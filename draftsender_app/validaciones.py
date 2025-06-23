import os
import re
import logging
from tkinter import messagebox

logger = logging.getLogger("DraftSender")

def es_email_valido(correo: str) -> bool:
    """
    Valida si una dirección de correo tiene un formato válido usando expresión regular.

    Args:
        correo (str): Dirección de correo a evaluar.

    Returns:
        bool: True si el formato es válido, False si no lo es.
    """
    if not correo:
        return False
    patron = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return re.match(patron, correo) is not None

def validar_rutas_archivos(ruta_excel: str, ruta_docx: str) -> bool:
    """
    Valida que las rutas de los archivos Excel y Word estén presentes y tengan extensiones correctas.

    Args:
        ruta_excel (str): Ruta del archivo Excel.
        ruta_docx (str): Ruta del archivo Word.

    Returns:
        tuple[bool, str]: (True, "") si las rutas son válidas; (False, "mensaje de error") en caso contrario.
    """
    ruta_excel = ruta_excel.strip() if ruta_excel else ""
    ruta_docx = ruta_docx.strip() if ruta_docx else ""

    if not ruta_excel or not ruta_docx:
        messagebox.showerror("Error", "Debes cargar Excel y Word antes de continuar.")
        logger.error(f"Faltan archivos - Excel: {bool(ruta_excel)}, Word: {bool(ruta_docx)}")
        return False
    if not os.path.exists(ruta_excel):
        messagebox.showerror("Error", f"No se encontró el archivo Excel:\n{ruta_excel}")
        logger.error("Archivo Excel no encontrado: %s", ruta_excel.encode('ascii', errors='ignore').decode())
        return False
    if not os.path.exists(ruta_docx):
        messagebox.showerror("Error", f"No se encontró el archivo Word:\n{ruta_docx}")
        logger.error("Archivo Word no encontrado: %s", ruta_docx.encode('ascii', errors='ignore').decode())
        return False
    return True

def validar_perfil(perfil: str) -> bool:
    """
    Valida que el perfil seleccionado no esté vacío ni sea un marcador de posición.

    Args:
        perfil (str): Nombre del perfil seleccionado.

    Returns:
        bool: True si el perfil es válido, False si no.
    """
    if not perfil or perfil == "Selecciona un correo...":
        messagebox.showerror("Error", "Selecciona un correo válido.")
        logger.error(f"Correo no válido: {perfil}")
        return False
    return True

def validar_columnas_obligatorias(columnas_en_excel: set) -> bool:
    """
    Verifica que las columnas obligatorias existan en el DataFrame cargado desde Excel.

    Args:
        df (pd.DataFrame): DataFrame con los datos a validar.
        columnas_requeridas (list[str]): Lista de nombres de columnas requeridas.

    Returns:
        tuple[bool, str]: True y mensaje vacío si están todas; False y mensaje de error si faltan.
    """
    columnas_requeridas = {"Correo", "Asunto", "Nombre"}
    faltantes = columnas_requeridas - columnas_en_excel
    if faltantes:
        messagebox.showerror("Error", f"Faltan columnas obligatorias: {', '.join(faltantes)}")
        logger.error(f"Columnas faltantes: {faltantes}")
        return False
    return True