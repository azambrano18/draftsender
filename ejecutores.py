from tkinter import messagebox  # Para mostrar cuadros de mensaje en la interfaz gráfica
import estado  # Para acceder a las variables de estado globales
from borradores import generar_borradores  # Función para crear borradores de correos
from envios import enviar_borradores  # Función para enviar borradores de correos

from logger_utils import configurar_logger  # Para configurar el logger
logger = configurar_logger("ejecutores")  # Instancia del logger para este módulo

def validar_datos(perfil: str, requiere_archivos: bool = True) -> bool:
    """
    Valida que los datos necesarios estén presentes antes de ejecutar un script.
    Args:
        perfil (str): Perfil de Outlook seleccionado por el usuario.
        requiere_archivos (bool): Si se requieren archivos (Excel y Word) para el script.
    Returns:
        bool: True si todos los datos son válidos, False en caso contrario.
    """
    if perfil == "Seleccione perfil...":
        messagebox.showerror("Error", "Selecciona un perfil válido.")  # Muestra un mensaje de error si el perfil no es válido
        logger.error(f"Perfil inválido: {perfil}")  # Registra el error en el logger
        return False

    if not estado.cuenta_seleccionada:
        messagebox.showerror("Error", "Selecciona una cuenta asociada.")  # Muestra un mensaje de error si no se seleccionó cuenta
        logger.error("Cuenta asociada no seleccionada")  # Registra el error en el logger
        return False

    if requiere_archivos and (not estado.ruta_excel or not estado.ruta_docx):
        messagebox.showerror("Error", "Debes cargar Excel y Word antes de continuar.")  # Muestra un mensaje de error si faltan archivos
        logger.error(f"Faltan archivos - Excel: {bool(estado.ruta_excel)}, Word: {bool(estado.ruta_docx)}")  # Registra el error en el logger
        return False

    return True

def ejecutar_script(nombre_script_txt: str, perfil: str):
    """
    Ejecuta un script según su nombre y el perfil de Outlook.
    Args:
        nombre_script_txt (str): Nombre del script a ejecutar (por ejemplo, "crear_borradores" o "timer_sent").
        perfil (str): Perfil de Outlook seleccionado.
    Returns:
        None
    """
    solo_envio = "timer_sent" in nombre_script_txt.lower()  # Determina si es solo un envío de borradores

    if not validar_datos(perfil, requiere_archivos=not solo_envio):  # Valida los datos antes de ejecutar el script
        return

    try:
        if "crear_borradores" in nombre_script_txt.lower():  # Si el script es para crear borradores
            cantidad = generar_borradores(
                cuenta=estado.cuenta_seleccionada,
                perfil=perfil,
                ruta_excel=estado.ruta_excel,
                ruta_docx=estado.ruta_docx
            )
            messagebox.showinfo("Éxito", f"Se crearon {cantidad} borradores correctamente.")  # Muestra un mensaje de éxito
            logger.info(f"{cantidad} borradores creados correctamente.")  # Registra la creación de borradores en el logger
            return

        if "timer_sent" in nombre_script_txt.lower():  # Si el script es para enviar borradores
            cantidad = enviar_borradores(estado.cuenta_seleccionada)
            messagebox.showinfo("Éxito", f"Se enviaron {cantidad} borradores correctamente.")  # Muestra un mensaje de éxito
            logger.info(f"{cantidad} borradores enviados correctamente.")  # Registra el envío de borradores en el logger
            return

        messagebox.showerror("Error", f"No se reconoce el script: {nombre_script_txt}")  # Muestra un mensaje de error si el script no se reconoce
        logger.error(f"Script no reconocido: {nombre_script_txt}")  # Registra el error en el logger

    except Exception as e:
        logger.exception(f"Fallo al ejecutar {nombre_script_txt}")  # Registra la excepción en el logger
        messagebox.showerror("Error", f"No se pudo ejecutar {nombre_script_txt}:\n{e}")  # Muestra un mensaje de error en caso de fallo

def ejecutar_timer_send(perfil: str):
    """
    Ejecuta el script de envío de borradores "timer_sent".
    Args:
        perfil (str): Perfil de Outlook seleccionado.
    Returns:
        None
    """
    ejecutar_script("timer_sent", perfil)  # Llama a la función 'ejecutar_script' con el nombre del script "timer_sent"