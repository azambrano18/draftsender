import os
import logging
import pandas as pd
from draftsender_app.validaciones import validar_columnas_obligatorias
from tkinter import filedialog, messagebox
from docx import Document

logger = logging.getLogger("DraftSender")

def cargar_excel(ruta_visible_var, ruta_real_var):
    """
Abre un diálogo para seleccionar un archivo Excel (.xlsx) válido.
Si el archivo es válido, actualiza las variables visibles y reales con la ruta del archivo.

Args:
    ruta_visible_var (tk.StringVar): Variable para mostrar el nombre del archivo seleccionado.
    ruta_real_var (tk.StringVar): Variable para guardar la ruta completa del archivo.
"""
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")])
    if archivo:
        try:
            if not archivo.lower().endswith((".xlsx", ".xls", ".xlsm")):
                raise ValueError("Extensión no válida")
            if not os.path.exists(archivo):
                raise FileNotFoundError("El archivo no existe")
            if os.path.getsize(archivo) > 10 * 1024 * 1024:
                raise ValueError("Archivo demasiado grande")

            nombre = os.path.basename(archivo)
            ruta_visible_var.set(f"... {nombre}")
            ruta_real_var.set(archivo.strip())
            logger.info(f"Archivo Excel cargado: {archivo}")

        except Exception as e:
            logger.error(f"Error al cargar Excel: {e}")
            messagebox.showerror("Error al cargar Excel", f"No se pudo cargar el archivo:\n{e}")

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
    Verifica que el archivo Excel tenga las columnas obligatorias requeridas.

    Args:
        ruta_excel (str): Ruta completa del archivo Excel.

    Returns:
        bool: True si es válido, False si no.
    """
    try:
        df = pd.read_excel(ruta_excel, sheet_name=1)
        if not validar_columnas_obligatorias(set(df.columns)):
            messagebox.showerror("Error en Excel", "El Excel no contiene las columnas necesarias.")
            return False
        return True
    except Exception as e:
        logger.error(f"Error al validar Excel: {e}")
        messagebox.showerror("Error al leer Excel", f"No se pudo abrir el archivo Excel:\n{e}")
        return False