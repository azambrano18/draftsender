import os
from tkinter import filedialog, messagebox
from docx import Document
from logger_utils import configurar_logger

logger = configurar_logger("archivos")

def cargar_excel(ruta_visible_var, ruta_real_var):
    """
    Abre un diálogo para seleccionar un archivo Excel y guarda su ruta completa y nombre visible por separado.
    """
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")])
    if archivo:
        nombre = os.path.basename(archivo)
        ruta_visible_var.set(f"... {nombre}")
        ruta_real_var.set(archivo.strip())
        logger.info(f"Archivo Excel cargado: {archivo}")

def cargar_docx(ruta_visible_var, ruta_real_var):
    """
    Selecciona y valida un archivo Word .docx. Guarda la ruta real y nombre visible por separado.
    """
    archivo = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])
    if archivo:
        try:
            if os.path.getsize(archivo) == 0:
                logger.error("El archivo .docx está vacío")
                messagebox.showerror("Archivo vacío", "El archivo Word seleccionado está vacío.")
                return

            doc = Document(archivo)
            contenido = "\n".join(p.text for p in doc.paragraphs).strip()
            if not contenido:
                logger.error("El archivo .docx no tiene contenido legible")
                messagebox.showerror("Sin contenido", "El archivo Word no contiene texto legible.")
                return

        except Exception as e:
            logger.exception("No se pudo validar el archivo .docx")
            messagebox.showerror("Error al validar", f"No se pudo validar el archivo Word:\n{e}")
            return

        nombre = os.path.basename(archivo)
        ruta_visible_var.set(f"... {nombre}")
        ruta_real_var.set(archivo.strip())
        logger.info(f"Archivo Word cargado: {archivo}")