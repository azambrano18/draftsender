import win32com.client  # Para interactuar con Outlook mediante COM
import pythoncom  # Para inicializar y liberar el modelo de objetos COM
import time  # Para gestionar los intervalos de espera
import threading  # Para ejecutar tareas en segundo plano
import sys  # Para gestionar argumentos de línea de comando y rutas
import os  # Para interactuar con el sistema de archivos

import tkinter as tk  # Para la interfaz gráfica
from tkinter import ttk, messagebox  # Para crear widgets y mostrar cuadros de mensaje

from logger_utils import configurar_logger  # Para configurar el logger
logger = configurar_logger("envios")  # Instancia del logger para este módulo

# Obtener la cuenta seleccionada desde los argumentos de la línea de comando
if len(sys.argv) > 1:
    cuenta_seleccionada = sys.argv[1]  # Recibe la cuenta seleccionada desde los argumentos de la línea de comando
else:
    cuenta_seleccionada = ""  # Si no se pasa ningún argumento, la cuenta será vacía

# Registra la cuenta seleccionada en el log
logger.info(f"Cuenta seleccionada: {cuenta_seleccionada}")

# Evento global utilizado para controlar el proceso de envío
enviar_event = threading.Event()

# Función para actualizar el contador (intervalo)
def actualizar_contador(event=None):
    """
    Actualiza el contador de borradores y el tiempo estimado basado en el intervalo seleccionado.
    Args:
        event: Evento opcional que puede ser usado para disparar la actualización.
    Returns:
        None
    """
    intervalo_str = combo_intervalo.get()
    if intervalo_str == "Seleccione intervalo." or not intervalo_str.isdigit():
        intervalo = 15
    else:
        intervalo = int(intervalo_str)

    combo_intervalo.set(str(intervalo))  # Actualiza el valor en el combo
    validar_intervalo()  # Valida el intervalo

    total_borradores = contar_borradores(cuenta_seleccionada)
    status_label.config(text=f"Borradores restantes: {total_borradores} | Enviados: 0")

    if total_borradores > 0:
        tiempo_total = intervalo * total_borradores
        horas, resto = divmod(tiempo_total, 3600)
        minutos, segundos = divmod(resto, 60)
        estimado_label.config(text=f"Tiempo total estimado: {horas:02}:{minutos:02}:{segundos:02}")
    else:
        estimado_label.config(text="Tiempo restante: 00:00:00")

# Función para validar el intervalo
def validar_intervalo():
    """
    Valida el intervalo de envío seleccionado en la interfaz gráfica.
    Si el intervalo es válido (un número), habilita el botón para iniciar el envío.
    Si el intervalo es inválido, deshabilita el botón de inicio.
    Returns:
        None
    """
    intervalo_str = combo_intervalo.get()
    if intervalo_str == "Seleccione intervalo." or not intervalo_str.isdigit():
        start_button.config(state="disabled")
    else:
        start_button.config(state="normal")

# Función para iniciar el temporizador dinámico
def iniciar_temporizador_dinamico(tiempo_total):
    """
    Inicia un temporizador dinámico que actualiza la interfaz con el tiempo restante.
    Args:
        tiempo_total (int): Tiempo total estimado para el envío en segundos.
    Returns:
        None
    """
    def actualizar_reloj():
        nonlocal tiempo_total
        if tiempo_total <= 0 or not enviar_event.is_set():
            estimado_label.config(text="Tiempo restante: 00:00:00")
            return
        horas, resto = divmod(tiempo_total, 3600)
        minutos, segundos = divmod(resto, 60)
        estimado_label.config(text=f"Tiempo restante: {horas:02}:{minutos:02}:{segundos:02}")
        tiempo_total -= 1
        root.after(1000, actualizar_reloj)

    actualizar_reloj()

# Función para obtener la carpeta de borradores en Outlook
def obtener_carpeta_borradores(namespace, cuenta_smtp):
    """
    Busca la carpeta de borradores en la cuenta de Outlook proporcionada.
    Args:
        namespace: Espacio de nombres de Outlook.
        cuenta_smtp (str): Dirección de correo de la cuenta a buscar.
    Returns:
        folder: Carpeta de borradores de la cuenta de Outlook.
    Raises:
        RuntimeError: Si no se encuentra la carpeta de borradores o la cuenta.
    """
    def buscar_recursivo(folder):
        for subfolder in folder.Folders:
            if subfolder.Name in ["Borradores", "Drafts"]:
                return subfolder
            resultado = buscar_recursivo(subfolder)
            if resultado:
                return resultado
        return None

    for account in namespace.Accounts:
        if account.SmtpAddress.lower() == cuenta_smtp.lower():
            try:
                root_folder = account.DeliveryStore.GetRootFolder()
                carpeta = buscar_recursivo(root_folder)
                if carpeta:
                    return carpeta
                raise LookupError(f"No se encontró la carpeta 'Borradores' en la cuenta {cuenta_smtp}")
            except Exception as e:
                logger.exception(f"No se pudo acceder a la carpeta raíz de la cuenta {cuenta_smtp}")
                raise RuntimeError(f"No se pudo acceder a la carpeta raíz de la cuenta {cuenta_smtp}: {e}")
    raise RuntimeError(f"No se encontró la cuenta en Outlook: {cuenta_smtp}")

# Función para contar los borradores
def contar_borradores(cuenta):
    """
    Cuenta el número de borradores en la carpeta de borradores de la cuenta seleccionada.
    Args:
        cuenta (str): Dirección de correo de la cuenta.
    Returns:
        int: Número de borradores en la carpeta de borradores.
    """
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        drafts_folder = obtener_carpeta_borradores(namespace, cuenta)
        if drafts_folder:
            return drafts_folder.Items.Count
        else:
            messagebox.showerror("Error", "No se encontró la carpeta de borradores.")
            return 0
    except Exception as e:
        logger.exception("Error al contar borradores")
        messagebox.showerror("Error al contar borradores", str(e))
        return 0
    finally:
        pythoncom.CoUninitialize()

# Función para enviar los borradores
def enviar_borradores(cuenta, status_label):
    """
    Envía los borradores de la cuenta seleccionada con un intervalo definido.
    Args:
        cuenta (str): Dirección de correo de la cuenta de Outlook.
        status_label (tk.Label): Etiqueta de la interfaz que muestra el estado del envío.
    Returns:
        None
    """
    enviar_event.set()  # Inicia el evento de envío
    intervalo = int(combo_intervalo.get())  # Obtiene el intervalo de envío
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        drafts_folder = obtener_carpeta_borradores(namespace, cuenta)

        if not drafts_folder:
            logger.error("No se encontró la carpeta de borradores.")
            messagebox.showerror("Error", "No se encontró la carpeta de borradores.")
            return

        total_borradores = drafts_folder.Items.Count
        enviados = 0

        time.sleep(intervalo)  # Espera antes de comenzar a enviar

        while enviados < total_borradores and enviar_event.is_set():
            item = drafts_folder.Items.GetFirst()  # Obtiene el primer borrador
            if not item:
                break

            try:
                sin_resolver = [r.Name for r in item.Recipients if not r.Resolve()]
                if sin_resolver:
                    raise Exception(f"Destinatarios no resueltos: {', '.join(sin_resolver)}")
                item.Send()  # Envía el borrador
                enviados += 1
            except Exception as e:
                logger.error(f"Error en el borrador #{enviados + 1}: {e}")
                messagebox.showerror("Error en borrador", f"Error en el borrador #{enviados + 1}: {e}")

            restantes = total_borradores - enviados
            status_label.config(text=f"Borradores restantes: {restantes} | Enviados: {enviados}")
            time.sleep(intervalo)  # Espera antes de enviar el siguiente borrador

        logger.info(f"Se enviaron {enviados} de {total_borradores} borradores.")

    except Exception as e:
        logger.exception("Error general durante el envío de borradores")
        messagebox.showerror("Error general", f"{e}")
    finally:
        status_label.config(text="Proceso finalizado")
        estimado_label.config(text="Tiempo restante: 00:00:00")
        enviar_event.clear()
        pythoncom.CoUninitialize()

# Función para iniciar el proceso de envío
def iniciar_envio():
    """
    Inicia el proceso de envío de borradores en segundo plano.
    Calcula el tiempo total necesario para el envío y lanza los hilos para el temporizador y el envío.
    Returns:
        None
    """
    total_borradores = contar_borradores(cuenta_seleccionada)
    intervalo = int(combo_intervalo.get())
    if total_borradores > 0:
        enviar_event.set()
        tiempo_total = intervalo * total_borradores
        threading.Thread(target=lambda: iniciar_temporizador_dinamico(tiempo_total), daemon=True).start()
        threading.Thread(target=enviar_borradores, args=(cuenta_seleccionada, status_label), daemon=True).start()

# Función para detener el proceso de envío
def detener_envio():
    """
    Detiene el proceso de envío de borradores.
    Cancela el evento de envío y actualiza la interfaz con el estado detenido.
    Returns:
        None
    """
    enviar_event.clear()
    status_label.config(text="Envío detenido")
    estimado_label.config(text="Tiempo restante: --")
    logger.info("Envío detenido manualmente.")

# EJECUCIÓN SOLO SI ES LLAMADO DIRECTAMENTE
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Enviar Borradores Outlook")
    root.geometry("400x280")

    try:
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        icon_path = os.path.join(base_path, "config", "icono.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        logger.warning(f"No se pudo cargar el ícono: {e}")

    tk.Label(root, text=f"Correo seleccionado: {cuenta_seleccionada}", font=("Arial", 12), fg="green").pack(pady=5)

    tk.Label(root, text="Intervalo de envío (segundos):", font=("Arial", 12)).pack(pady=5)
    combo_intervalo = ttk.Combobox(root, values=["Seleccione intervalo.", "60", "120", "180"], state="readonly",
                                   font=("Arial", 10), width=20)
    combo_intervalo.bind("<<ComboboxSelected>>", actualizar_contador)
    combo_intervalo.pack(pady=5)
    combo_intervalo.current(0)

    start_button = tk.Button(root, text="Iniciar Envío", command=iniciar_envio, font=("Arial", 12), bg="lightgreen",
                             state="disabled")
    start_button.pack(pady=5)

    tk.Button(root, text="Detener Envío", command=detener_envio, font=("Arial", 10), bg="red").pack(pady=5)

    status_label = tk.Label(root, text="", font=("Arial", 14))
    status_label.pack(pady=5)

    estimado_label = tk.Label(root, text="Tiempo restante: --", font=("Arial", 12), fg="blue")
    estimado_label.pack(pady=5)

    validar_intervalo()
    root.mainloop()