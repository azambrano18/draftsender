import os
import logging
import pandas as pd
from draftsender_app.validaciones import validar_columnas_obligatorias
from tkinter import filedialog, messagebox
from docx import Document

logger = logging.getLogger("DraftSender")

def cargar_excel(ruta_visible_var, ruta_real_var) -> bool:
    """
    Abre un diálogo para seleccionar un archivo Excel válido.

    Si el archivo es válido, actualiza las variables visibles y reales con la ruta del archivo.

    Args:
        ruta_visible_var (tk.StringVar): Variable para mostrar el nombre del archivo seleccionado.
        ruta_real_var (tk.StringVar): Variable para guardar la ruta completa del archivo.

    Returns:
        bool: True si se cargó correctamente, False si hubo error o cancelación.
    """
    archivo = filedialog.askopenfilename(
        filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")]
    )

    if not archivo:
        return False

    try:
        if not archivo.lower().endswith((".xlsx", ".xls", ".xlsm")):
            raise ValueError("Extensión no válida. Usa un archivo .xlsx, .xls o .xlsm.")

        if not os.path.exists(archivo):
            raise FileNotFoundError("El archivo no existe.")

        if os.path.getsize(archivo) > 10 * 1024 * 1024:
            raise ValueError("Archivo demasiado grande. Máximo permitido: 10 MB.")

        nombre = os.path.basename(archivo)
        ruta_visible_var.set(f"... {nombre}")
        ruta_real_var.set(archivo.strip())

        logger.info(f"Archivo Excel cargado: {archivo}")
        return True

    except Exception as e:
        logger.error(f"Error al cargar Excel: {e}")
        messagebox.showerror(
            "Error al cargar Excel",
            f"No se pudo cargar el archivo:\n{e}"
        )
        return False

def cargar_docx(ruta_visible_var, ruta_real_var):
    """
Abre un diálogo para seleccionar un archivo Word (.docx) válido.
Si el archivo es válido, actualiza las variables visibles y reales con la ruta del archivo.

Args:
    ruta_visible_var (tk.StringVar): Variable para mostrar el nombre del archivo seleccionado.
    ruta_real_var (tk.StringVar): Variable para guardar la ruta completa del archivo.
"""
    archivo = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])
    if archivo:
        try:
            if not archivo.lower().endswith(".docx"):
                raise ValueError("Extensión no válida")
            if not os.path.exists(archivo):
                raise FileNotFoundError("El archivo no existe")
            if os.path.getsize(archivo) > 5 * 1024 * 1024:
                raise ValueError("Archivo demasiado grande")

            doc = Document(archivo)
            contenido = "\n".join(p.text for p in doc.paragraphs).strip()
            if not contenido:
                raise ValueError("El archivo Word no contiene texto legible")

            nombre = os.path.basename(archivo)
            ruta_visible_var.set(f"... {nombre}")
            ruta_real_var.set(archivo.strip())
            logger.info(f"Archivo Word cargado: {archivo}")

        except Exception as e:
            logger.exception("No se pudo validar el archivo .docx")
            messagebox.showerror("Error al cargar Word", f"No se pudo validar el archivo Word:\n{e}")

def validar_excel_gui(ruta_excel: str) -> bool:
    """
    Verifica que el archivo Excel tenga una hoja válida con las columnas obligatorias.

    Args:
        ruta_excel (str): Ruta completa del archivo Excel.

    Returns:
        bool: True si es válido, False si no.
    """
    try:
        df = leer_excel_destinatarios(ruta_excel)

        if not validar_columnas_obligatorias(set(df.columns)):
            messagebox.showerror(
                "Error en Excel",
                "El Excel no contiene las columnas necesarias: Correo, Asunto y Nombre."
            )
            return False

        return True

    except Exception as e:
        logger.error(f"Error al validar Excel: {e}")
        messagebox.showerror(
            "Error al leer Excel",
            f"No se pudo abrir el archivo Excel:\n{e}"
        )
        return False

def leer_excel_destinatarios(ruta_excel: str) -> pd.DataFrame:
    """
    Lee el Excel de destinatarios buscando automáticamente una hoja válida.

    Una hoja válida debe contener las columnas obligatorias:
    - Correo
    - Asunto
    - Nombre

    Args:
        ruta_excel (str): Ruta del archivo Excel.

    Returns:
        pd.DataFrame: DataFrame con los destinatarios.

    Raises:
        FileNotFoundError: Si el archivo no existe.
        ValueError: Si no existe una hoja válida.
    """
    if not ruta_excel or not os.path.exists(ruta_excel):
        raise FileNotFoundError("No se encontró el archivo Excel.")

    try:
        hojas = pd.read_excel(ruta_excel, sheet_name=None)
    except Exception as e:
        raise ValueError(f"No se pudo abrir el archivo Excel: {e}")

    columnas_obligatorias = {"Correo", "Asunto", "Nombre"}

    for nombre_hoja, df in hojas.items():
        columnas = set(str(col).strip() for col in df.columns)

        if columnas_obligatorias.issubset(columnas):
            df.columns = [str(col).strip() for col in df.columns]
            return df

    hojas_disponibles = ", ".join(hojas.keys())

    raise ValueError(
        "No se encontró ninguna hoja válida en el Excel.\n\n"
        "La hoja debe contener las columnas obligatorias: "
        "Correo, Asunto y Nombre.\n\n"
        f"Hojas encontradas: {hojas_disponibles}"
    )