import logging
from draftsender_app.ui_utils import mostrar_advertencia
from draftsender_app.ui_utils import mostrar_error, mostrar_info
from draftsender_app.validaciones import validar_perfil, validar_rutas_archivos

logger = logging.getLogger("DraftSender")

def validar_datos(perfil: str, requiere_archivos: bool = True, ruta_excel: str = None, ruta_docx: str = None) -> bool:
    """
Valida que se hayan proporcionado todas las rutas y selecciones necesarias para ejecutar un envío.

Args:
    cuenta (str): Dirección de correo seleccionada.
    ruta_excel (str): Ruta al archivo Excel con los destinatarios.
    ruta_docx (str): Ruta al archivo Word con el cuerpo del mensaje.

Returns:
    bool: True si los datos están completos y válidos, False en caso contrario.
"""

    if not validar_perfil(perfil):
        return False
    if requiere_archivos and not validar_rutas_archivos(ruta_excel, ruta_docx):
        return False
    return True

def ejecutar_script(nombre_script_txt: str, perfil: str, ruta_excel: str = None, ruta_docx: str = None, callback_progreso=None, modo_envio: str = "Envíos 1"):
    """
    Ejecuta lógicamente la creación o el envío de borradores.

    Args:
        nombre_script_txt (str): Nombre lógico del script a ejecutar.
        perfil (str): Perfil de Outlook (correo usado para identificar la cuenta).
        ruta_excel (str): Ruta al archivo Excel.
        ruta_docx (str): Ruta al archivo Word.
        callback_progreso (callable): Función callback para progreso visual.
        modo_envio (str): "Envíos 1" o "Seguimiento".
    """
    try:
        es_envio = "envios" in nombre_script_txt.lower()
        es_borradores = "crear_borradores" in nombre_script_txt.lower()

        if not validar_datos(perfil, requiere_archivos=not es_envio, ruta_excel=ruta_excel, ruta_docx=ruta_docx):
            return

        if es_borradores:
            from draftsender_app.borradores import generar_borradores
            cantidad, errores, mensaje = generar_borradores(
                cuenta=perfil,
                perfil=perfil,
                ruta_excel=ruta_excel,
                ruta_docx=ruta_docx,
                modo_envio=modo_envio,
                callback_progreso=callback_progreso)
            if errores:
                mostrar_advertencia("Errores en la creación", mensaje)
            else:
                mostrar_info("Borradores creados", mensaje)
            return

        if es_envio:
            from draftsender_app.envios import enviar_borradores
            enviados, errores, mensaje = enviar_borradores(perfil)
            if errores:
                mostrar_advertencia("Errores durante el envío", mensaje)
            else:
                mostrar_info("Envío completado", mensaje)
            return

        mostrar_error("Script no reconocido", f"No se reconoce el script: {nombre_script_txt}")

    except Exception as e:
        logger.exception(f"Error al ejecutar script '{nombre_script_txt}'")
        mostrar_error("Fallo al ejecutar script", f"No se pudo ejecutar '{nombre_script_txt}':\n{e}")

def ejecutar_envios(perfil: str):
    """
    Lanza el ejecutable 'envios.exe' si existe, permitiendo iniciar el envío de borradores manualmente.
    """
    ejecutar_script("envios", perfil)