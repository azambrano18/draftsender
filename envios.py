import win32com.client
import pythoncom
import time
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys
import os

from logger_utils import configurar_logger
logger = configurar_logger("envios")

cuenta_seleccionada = ""
enviar_event = threading.Event()

def obtener_carpeta_borradores(namespace, cuenta_smtp):
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

def contar_borradores(cuenta):
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

def validar_intervalo():
    intervalo_str = combo_intervalo.get()
    if intervalo_str == "Seleccione intervalo." or not intervalo_str.isdigit():
        start_button.config(state="disabled")
    else:
        start_button.config(state="normal")

def actualizar_contador(event=None):
    intervalo_str = combo_intervalo.get()
    if intervalo_str == "Seleccione intervalo." or not intervalo_str.isdigit():
        intervalo = 15
    else:
        intervalo = int(intervalo_str)

    combo_intervalo.set(str(intervalo))
    validar_intervalo()

    total_borradores = contar_borradores(cuenta_seleccionada)
    status_label.config(text=f"Borradores restantes: {total_borradores} | Enviados: 0")

    if total_borradores > 0:
        tiempo_total = intervalo * total_borradores
        horas, resto = divmod(tiempo_total, 3600)
        minutos, segundos = divmod(resto, 60)
        estimado_label.config(text=f"Tiempo total estimado: {horas:02}:{minutos:02}:{segundos:02}")
    else:
        estimado_label.config(text="Tiempo restante: 00:00:00")

def iniciar_temporizador_dinamico(tiempo_total):
    def actualizar_reloj():
        nonlocal tiempo_total
        global after_id

        try:
            if tiempo_total <= 0 or not enviar_event.is_set():
                if estimado_label.winfo_exists():
                    estimado_label.config(text="Tiempo restante: 00:00:00")
                return

            horas, resto = divmod(tiempo_total, 3600)
            minutos, segundos = divmod(resto, 60)

            if estimado_label.winfo_exists():
                estimado_label.config(text=f"Tiempo restante: {horas:02}:{minutos:02}:{segundos:02}")

            tiempo_total -= 1

            if root.winfo_exists():
                after_id = root.after(1000, actualizar_reloj)

        except tk.TclError:
            logger.warning("after() cancelado tras cierre de ventana")

    actualizar_reloj()

def enviar_borradores(cuenta, status_label):
    enviar_event.set()
    intervalo = int(combo_intervalo.get())
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

        time.sleep(intervalo)

        while enviados < total_borradores and enviar_event.is_set():
            item = drafts_folder.Items.GetFirst()
            if not item:
                break

            try:
                sin_resolver = [r.Name for r in item.Recipients if not r.Resolve()]
                if sin_resolver:
                    raise Exception(f"Destinatarios no resueltos: {', '.join(sin_resolver)}")
                item.Send()
                enviados += 1
            except Exception as e:
                logger.error(f"Error en el borrador #{enviados + 1}: {e}")
                messagebox.showerror("Error en borrador", f"Error en el borrador #{enviados + 1}: {e}")

            restantes = total_borradores - enviados
            status_label.config(text=f"Borradores restantes: {restantes} | Enviados: {enviados}")
            time.sleep(intervalo)

        logger.info(f"Se enviaron {enviados} de {total_borradores} borradores.")
    except Exception as e:
        logger.exception("Error general durante el envío de borradores")
        messagebox.showerror("Error general", f"{e}")
    finally:
        status_label.config(text="Proceso finalizado")
        estimado_label.config(text="Tiempo restante: 00:00:00")
        enviar_event.clear()
        pythoncom.CoUninitialize()

def iniciar_envio():
    total_borradores = contar_borradores(cuenta_seleccionada)
    intervalo = int(combo_intervalo.get())
    if total_borradores > 0:
        enviar_event.set()
        tiempo_total = intervalo * total_borradores
        threading.Thread(target=lambda: iniciar_temporizador_dinamico(tiempo_total), daemon=True).start()
        threading.Thread(target=enviar_borradores, args=(cuenta_seleccionada, status_label), daemon=True).start()

def detener_envio():
    global after_id
    enviar_event.clear()

    if root.winfo_exists() and after_id:
        try:
            root.after_cancel(after_id)
        except Exception:
            pass

    status_label.config(text="Envío detenido")
    estimado_label.config(text="Tiempo restante: --")
    logger.info("Envío detenido manualmente.")

# Función principal para ser llamada desde gui.py
def lanzar_envio_desde_gui(cuenta):
    global root, combo_intervalo, start_button, status_label, estimado_label, cuenta_seleccionada
    cuenta_seleccionada = cuenta

    root = tk.Tk()
    root.title("Enviar Borradores Outlook")
    root.geometry("400x280")

    try:
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        icon_path = os.path.join(base_path, "config", "icono.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        logger.warning(f"No se pudo cargar el ícono: {e}")

    tk.Label(root, text=f"Correo seleccionado: {cuenta}", font=("Arial", 12), fg="green").pack(pady=5)

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

# Soporte para ejecución directa en desarrollo
if __name__ == "__main__":
    cuenta = sys.argv[1] if len(sys.argv) > 1 else ""
    lanzar_envio_desde_gui(cuenta)