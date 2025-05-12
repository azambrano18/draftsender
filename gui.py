import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

from logger_utils import configurar_logger
from outlook_utils import obtener_perfiles_outlook, cerrar_outlook, iniciar_outlook_con_perfil, obtener_cuentas_activas
from ejecutores import ejecutar_script
from archivos import cargar_excel, cargar_docx
from actualizacion import verificar_actualizacion
import estado

logger = configurar_logger("gui")

combo_cuentas = None
combo_cuentas_asociadas = None
ruta_excel_var = None
ruta_docx_var = None
label_cuenta_var = None
barra_progreso = None
porcentaje_var = None
frame_progreso = None
status_var = None
status_label = None

PERFIL_DEFAULT = "Seleccione perfil..."

def construir_gui(root):
    global combo_cuentas, combo_cuentas_asociadas, ruta_excel_var, ruta_docx_var, label_cuenta_var

    label_cuenta_var = tk.StringVar()
    crear_menu(root)
    crear_encabezado(root)
    crear_selector_perfil(root)
    crear_cargadores_archivos(root)
    crear_botonera(root)
    crear_barra_progreso(root)

def crear_menu(root):
    menu_bar = tk.Menu(root)
    menu_archivo = tk.Menu(menu_bar, tearoff=0)
    menu_archivo.add_command(label="Actualizar", command=lambda: actualizar_aplicacion_intermedia(root, barra_progreso, porcentaje_var, frame_progreso, status_label, status_var))
    menu_archivo.add_command(label="Salir", command=root.quit)
    menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

    menu_ayuda = tk.Menu(menu_bar, tearoff=0)
    menu_ayuda.add_command(label="Acerca de", command=mostrar_acerca_de)
    menu_bar.add_cascade(label="Ayuda", menu=menu_ayuda)

    root.config(menu=menu_bar)

def crear_encabezado(root):
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        img_path = os.path.join(base_path, "config", "cover_borradores.jpg")
        img = Image.open(img_path).resize((500, 90))
        img_tk = ImageTk.PhotoImage(img)
        label_img = tk.Label(root, image=img_tk)
        label_img.image = img_tk
        label_img.pack(pady=10)
    except Exception as e:
        logger.error("No se pudo cargar la imagen de portada", exc_info=True)

def crear_selector_perfil(root):
    global combo_cuentas, combo_cuentas_asociadas
    perfiles = obtener_perfiles_outlook()
    max_length = max(len(p) for p in perfiles)
    tk.Label(root, text="Selecciona un perfil:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10)

    frame_perfiles = tk.Frame(root)
    frame_perfiles.pack(anchor="w", padx=10, pady=5)
    combo_cuentas = ttk.Combobox(frame_perfiles, values=perfiles, state="readonly", font=("Arial", 10), width=max_length + 2)
    combo_cuentas.pack(side="left")
    combo_cuentas.bind("<<ComboboxSelected>>", mostrar_cuenta_seleccionada)
    combo_cuentas.current(0)

    # Cuenta asociada
    tk.Label(root, textvariable=label_cuenta_var, font=("Arial", 10)).pack(pady=5)
    combo_cuentas_asociadas = ttk.Combobox(root, state="readonly", font=("Arial", 10))
    combo_cuentas_asociadas.bind("<<ComboboxSelected>>", lambda e: seleccionar_cuenta_asociada())
    combo_cuentas_asociadas.pack_forget()

def crear_cargadores_archivos(root):
    global ruta_excel_var, ruta_docx_var
    ruta_excel_var = tk.StringVar()
    ruta_docx_var = tk.StringVar()

    frame_excel = tk.Frame(root)
    frame_excel.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_excel, text="Cargar Excel", command=lambda: cargar_excel(ruta_excel_var), font=("Arial", 10)).pack(side="left")
    tk.Label(frame_excel, textvariable=ruta_excel_var, fg="green").pack(side="left", padx=10)

    frame_docx = tk.Frame(root)
    frame_docx.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_docx, text="Cargar Texto Mail", command=lambda: cargar_docx(ruta_docx_var), font=("Arial", 10)).pack(side="left")
    tk.Label(frame_docx, textvariable=ruta_docx_var, fg="green").pack(side="left", padx=10)

def crear_botonera(root):
    frame_boton = tk.Frame(root)
    frame_boton.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_boton, text="Crear Borradores", command=lambda: ejecutar_script("crear_borradores", combo_cuentas.get()), font=("Arial", 10)).pack(side="left")

    frame_enviar = tk.Frame(root)
    frame_enviar.pack(anchor="center", pady=5)
    tk.Button(frame_enviar, text="Enviar Borradores", command=lanzar_envio_gui, font=("Arial", 12), bg="purple", fg="white").pack()

def crear_barra_progreso(root):
    global barra_progreso, porcentaje_var, frame_progreso, status_var, status_label

    frame_progreso = tk.Frame(root)  # Oculto al inicio

    barra_progreso = ttk.Progressbar(frame_progreso, length=300, mode='determinate', maximum=100)
    barra_progreso.pack(side="left", padx=(0, 10))

    porcentaje_var = tk.StringVar(value="0%")
    tk.Label(frame_progreso, textvariable=porcentaje_var).pack(side="left")

    status_var = tk.StringVar(value="")
    status_label = tk.Label(root, textvariable=status_var, font=("Arial", 10), fg="blue")
    status_label.pack(side="bottom", pady=(0, 5))
    status_label.pack_forget()  # Ocultamos al iniciar

def actualizar_aplicacion_intermedia(root, barra_progreso, porcentaje_var, frame_progreso, status_label, status_var):
    verificar_actualizacion(root, barra_progreso, porcentaje_var, frame_progreso, status_label, status_var)

def mostrar_acerca_de():
    messagebox.showinfo("Acerca de", "DraftSender - Automatización de borradores y envíos de correos.")

def mostrar_cuenta_seleccionada(_event=None):
    perfil = combo_cuentas.get()
    estado.cuenta_seleccionada = None
    ruta_excel_var.set("")
    ruta_docx_var.set("")
    combo_cuentas_asociadas.pack_forget()
    combo_cuentas_asociadas.set("")

    if perfil == PERFIL_DEFAULT:
        label_cuenta_var.set("")
        return

    cerrar_outlook()
    iniciar_outlook_con_perfil(perfil)
    cuentas = obtener_cuentas_activas()

    if not cuentas:
        logger.error("No se encontraron cuentas.")
        messagebox.showerror("Error", "No se encontraron cuentas activas en Outlook.")
        label_cuenta_var.set("No se encontraron cuentas.")
    elif len(cuentas) == 1:
        estado.cuenta_seleccionada = cuentas[0]
        label_cuenta_var.set(f"{estado.cuenta_seleccionada} ✔️")
    else:
        label_cuenta_var.set("Selecciona una cuenta:")
        combo_cuentas_asociadas['values'] = cuentas
        combo_cuentas_asociadas.current(0)
        combo_cuentas_asociadas.pack()
        logger.info(f"Cuenta seleccionada: {estado.cuenta_seleccionada}")

def seleccionar_cuenta_asociada():
    estado.cuenta_seleccionada = combo_cuentas_asociadas.get()

def lanzar_envio_gui():
    if not estado.cuenta_seleccionada:
        messagebox.showerror("Error", "Debe seleccionar una cuenta antes de enviar borradores.")
        logger.error("No se ha seleccionado una cuenta para envío.")
        return

    from envios import lanzar_envio_desde_gui
    lanzar_envio_desde_gui(estado.cuenta_seleccionada)

# Inicialización directa (solo si se ejecuta este archivo)
if __name__ == "__main__":
    root = tk.Tk()
    root.title("DraftSender - Automatización de Borradores y Envíos")
    root.geometry("480x450")
    construir_gui(root)
    root.mainloop()