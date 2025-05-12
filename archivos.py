import os  # Para interactuar con el sistema de archivos
from tkinter import filedialog, messagebox  # Para abrir el diálogo de selección de archivos y mostrar mensajes
from docx import Document  # Para trabajar con archivos .docx
import estado  # Para acceder a las variables de estado globales

from logger_utils import configurar_logger  # Para configurar el logger
logger = configurar_logger("archivos")  # Instancia del logger para este módulo

def cargar_excel(ruta_var):
    """
    Abre un diálogo para seleccionar un archivo Excel (.xlsx, .xls, .xlsm) y guarda su ruta en el estado global.
    Actualiza la variable `ruta_var` con el nombre del archivo seleccionado.

    Args:
        ruta_var (tk.StringVar): Variable de Tkinter para mostrar la ruta del archivo en la interfaz.

    Returns:
        None
    """
    # Abre el diálogo de selección de archivo para elegir un archivo Excel
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")])
    if archivo:
        # Actualiza la variable global 'estado.ruta_excel' con la ruta seleccionada
        estado.ruta_excel = archivo
        nombre = os.path.basename(archivo)  # Obtiene el nombre del archivo
        ruta_var.set(f"... {nombre}   ✔️")  # Actualiza la interfaz con el nombre del archivo
        logger.info(f"Archivo Excel cargado: {archivo}")  # Registra la carga del archivo

def cargar_docx(ruta_var):
    """
    Abre un diálogo para seleccionar un archivo Word (.docx) y valida su contenido.
    Si el archivo es válido, guarda su ruta en el estado global y actualiza la variable `ruta_var`.

    Args:
        ruta_var (tk.StringVar): Variable de Tkinter para mostrar la ruta del archivo en la interfaz.

    Returns:
        None
    """
    # Abre el diálogo de selección de archivo para elegir un archivo Word
    archivo = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])
    if archivo:
        try:
            # Verifica si el archivo está vacío
            if os.path.getsize(archivo) == 0:
                logger.error("El archivo .docx está vacío")
                messagebox.showerror("Archivo vacío", "El archivo Word seleccionado está vacío.")  # Muestra un mensaje de error
                return

            # Intenta leer el contenido del archivo .docx
            doc = Document(archivo)
            contenido = "\n".join(p.text for p in doc.paragraphs).strip()
            if not contenido:
                logger.error("El archivo .docx no tiene contenido legible")
                messagebox.showerror("Sin contenido", "El archivo Word no contiene texto legible.")  # Muestra un mensaje de error
                return

        except Exception as e:
            # Si ocurre un error al validar el archivo, muestra un mensaje de error
            logger.exception("No se pudo validar el archivo .docx")
            messagebox.showerror("Error al validar", f"No se pudo validar el archivo Word:\n{e}")
            return

        # Actualiza la variable global 'estado.ruta_docx' con la ruta seleccionada
        estado.ruta_docx = archivo
        nombre = os.path.basename(archivo)  # Obtiene el nombre del archivo
        ruta_var.set(f"... {nombre}   ✔️")  # Actualiza la interfaz con el nombre del archivo
        logger.info(f"Archivo Word cargado: {archivo}")  # Registra la carga del archivo