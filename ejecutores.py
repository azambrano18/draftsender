from tkinter import messagebox
import os
from borradores import generar_borradores
from envios import enviar_borradores
from logger_utils import configurar_logger

logger = configurar_logger("ejecutores")

def validar_datos(perfil: str, requiere_archivos: bool = True, ruta_excel: str = None, ruta_docx: str = None) -> bool:
    if not perfil or perfil == "Selecciona un correo...":
        messagebox.showerror("Error", "Selecciona un correo válido.")
        logger.error(f"Correo no válido: {perfil}")
        return False

    if requiere_archivos:
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

def ejecutar_script(nombre_script_txt: str, perfil: str, ruta_excel: str = None, ruta_docx: str = None, callback_progreso=None):
    solo_envio = "envios" in nombre_script_txt.lower() or "envios" in nombre_script_txt.lower()

    if not validar_datos(perfil, requiere_archivos=not solo_envio, ruta_excel=ruta_excel, ruta_docx=ruta_docx):
        return

    try:
        if "crear_borradores" in nombre_script_txt.lower():
            cantidad = generar_borradores(
                cuenta=perfil,
                perfil=perfil,
                ruta_excel=ruta_excel,
                ruta_docx=ruta_docx,
                callback_progreso=callback_progreso
            )
            logger.info(f"{cantidad} borradores creados correctamente.")
            return

        if solo_envio:
            cantidad = enviar_borradores(perfil)
            messagebox.showinfo("Éxito", f"Se enviaron {cantidad} borradores correctamente.")
            logger.info(f"{cantidad} borradores enviados correctamente.")
            return

        messagebox.showerror("Error", f"No se reconoce el script: {nombre_script_txt}")
        logger.error(f"Script no reconocido: {nombre_script_txt}")

    except Exception as e:
        logger.exception(f"Fallo al ejecutar {nombre_script_txt}")
        messagebox.showerror("Error", f"No se pudo ejecutar {nombre_script_txt}:\n{e}")

def ejecutar_envios(perfil: str):
    ejecutar_script("envios", perfil)