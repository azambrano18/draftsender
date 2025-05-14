import os  # Para manejo de rutas del sistema operativo
import sys  # Acceso a variables y funciones del sistema
import tkinter as tk  # Biblioteca principal para la interfaz gráfica
import estado
from tkinter import ttk, messagebox  # Widgets mejorados y cuadros de diálogo
from PIL import Image, ImageTk  # Para cargar y mostrar imágenes en Tkinter
from logger_utils import configurar_logger  # Función para iniciar el logger de eventos
from outlook_utils import obtener_perfiles_outlook, cerrar_outlook, iniciar_outlook_con_perfil, obtener_cuentas_activas  # Funciones relacionadas a perfiles Outlook
from ejecutores import ejecutar_script  # Ejecuta scripts personalizados
from archivos import cargar_excel, cargar_docx  # Funciones para cargar archivos
from actualizacion import verificar_actualizacion  # Lógica de actualización automática

logger = configurar_logger("gui")  # Función para iniciar el logger de eventos

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
label_cuenta_asociada = None
boton_crear_borradores = None

def wrapper_callback(fila_actual, total_filas):
    from progreso_utils import callback_progreso_gui
    callback_progreso_gui(
        fila_actual,
        total_filas,
        barra_progreso,
        porcentaje_var,
        status_var,
        status_label,
        frame_progreso
    )

PERFIL_DEFAULT = "Seleccione perfil..."  # Texto por defecto en selector de perfil

def construir_gui(root):
    global combo_cuentas, combo_cuentas_asociadas, ruta_excel_var, ruta_docx_var, label_cuenta_var, label_cuenta_asociada, boton_crear_borradores

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
    global combo_cuentas, combo_cuentas_asociadas, label_cuenta_asociada

    perfiles = obtener_perfiles_outlook()
    max_length = max(len(p) for p in perfiles)

    tk.Label(root, text="Selecciona un perfil y cuenta:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10)

    frame_perfiles = tk.Frame(root)
    frame_perfiles.pack(anchor="w", padx=10, pady=5)

    frame_perfil = tk.Frame(frame_perfiles)
    frame_perfil.pack(side="left", padx=(0, 10))

    tk.Label(frame_perfil, text="Perfil Outlook", font=("Arial", 10)).pack(anchor="w")
    combo_cuentas = ttk.Combobox(frame_perfil, values=perfiles, state="readonly", font=("Arial", 10), width=max_length + 2)
    combo_cuentas.pack()
    combo_cuentas.bind("<<ComboboxSelected>>", mostrar_cuenta_seleccionada)
    combo_cuentas.current(0)

    frame_cuenta = tk.Frame(frame_perfiles)
    frame_cuenta.pack(side="left")

    label_cuenta_asociada = tk.Label(frame_cuenta, text="Cuenta asociada", font=("Arial", 10))
    label_cuenta_asociada.pack(anchor="w")
    label_cuenta_asociada.pack_forget()

    combo_cuentas_asociadas = ttk.Combobox(frame_cuenta, state="readonly", font=("Arial", 10), width=30)
    combo_cuentas_asociadas.pack()
    combo_cuentas_asociadas.pack_forget()
    combo_cuentas_asociadas.bind("<<ComboboxSelected>>", lambda e: seleccionar_cuenta_asociada())

    tk.Label(root, textvariable=label_cuenta_var, font=("Arial", 10)).pack(anchor="w", padx=10, pady=(0, 5))

def crear_cargadores_archivos(root):
    global ruta_excel_var, ruta_docx_var
    ruta_excel_var = tk.StringVar()
    ruta_docx_var = tk.StringVar()

    frame_excel = tk.Frame(root)
    frame_excel.pack(anchor="w", padx=10, pady=5)

    def cargar_excel_y_validar():
        cargar_excel(ruta_excel_var)
        validar_estado_para_boton_borradores()

    tk.Button(frame_excel, text="Cargar Excel", command=cargar_excel_y_validar, font=("Arial", 10)).pack(side="left")
    tk.Label(frame_excel, textvariable=ruta_excel_var, fg="green").pack(side="left", padx=10)

    frame_docx = tk.Frame(root)
    frame_docx.pack(anchor="w", padx=10, pady=5)

    def cargar_docx_y_validar():
        cargar_docx(ruta_docx_var)
        validar_estado_para_boton_borradores()

    tk.Button(frame_docx, text="Cargar Texto Mail", command=cargar_docx_y_validar, font=("Arial", 10)).pack(side="left")
    tk.Label(frame_docx, textvariable=ruta_docx_var, fg="green").pack(side="left", padx=10)

def crear_botonera(root):
    global boton_crear_borradores

    frame_boton = tk.Frame(root)
    frame_boton.pack(anchor="w", padx=10, pady=5)

    boton_crear_borradores = tk.Button(
        frame_boton,
        text="Crear Borradores",
        command=lambda: ejecutar_script("crear_borradores", combo_cuentas.get(), wrapper_callback),
        font=("Arial", 10),
        state="disabled"
    )
    boton_crear_borradores.pack(side="left")

    frame_enviar = tk.Frame(root)
    frame_enviar.pack(anchor="center", pady=5)
    tk.Button(frame_enviar, text="Enviar Borradores", command=lanzar_envio_gui, font=("Arial", 12), bg="purple", fg="white").pack()

def crear_barra_progreso(root):
    global barra_progreso, porcentaje_var, frame_progreso, status_var, status_label

    frame_progreso = tk.Frame(root)
    barra_progreso = ttk.Progressbar(frame_progreso, length=300, mode='determinate', maximum=100)
    barra_progreso.pack(side="left", padx=(0, 10))

    porcentaje_var = tk.StringVar(value="0%")
    tk.Label(frame_progreso, textvariable=porcentaje_var).pack(side="left")

    status_var = tk.StringVar(value="")
    status_label = tk.Label(root, textvariable=status_var, font=("Arial", 10), fg="blue")
    status_label.pack(side="bottom", pady=(0, 5))
    status_label.pack_forget()

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
    label_cuenta_asociada.pack_forget()
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
        return

    if len(cuentas) == 1:
        estado.cuenta_seleccionada = cuentas[0]
        label_cuenta_var.set(f"{estado.cuenta_seleccionada} ✔️")
    else:
        opciones = ["Seleccione cuenta..."] + cuentas
        combo_cuentas_asociadas["values"] = opciones
        combo_cuentas_asociadas.current(0)
        label_cuenta_asociada.pack(anchor="w")
        combo_cuentas_asociadas.pack()
        label_cuenta_var.set("")

    validar_estado_para_boton_borradores()

def seleccionar_cuenta_asociada():
    seleccion = combo_cuentas_asociadas.get()
    estado.cuenta_seleccionada = seleccion

    if seleccion and seleccion != "Seleccione cuenta...":
        label_cuenta_var.set(f"{seleccion} ✔️")
    else:
        label_cuenta_var.set("")

    validar_estado_para_boton_borradores()

def lanzar_envio_gui():
    if not estado.cuenta_seleccionada:
        messagebox.showerror("Error", "Debe seleccionar una cuenta antes de enviar borradores.")
        logger.error("No se ha seleccionado una cuenta para envío.")
        return

    from envios import lanzar_envio_desde_gui
    lanzar_envio_desde_gui(estado.cuenta_seleccionada)

def validar_estado_para_boton_borradores():
    archivos_cargados = estado.ruta_excel and estado.ruta_docx
    cuenta_valida = estado.cuenta_seleccionada and estado.cuenta_seleccionada != "Seleccione cuenta..."
    if archivos_cargados and cuenta_valida:
        boton_crear_borradores.config(state="normal")
    else:
        boton_crear_borradores.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("DraftSender - Automatización de Borradores y Envíos")
    root.geometry("480x450")
    construir_gui(root)
    root.mainloop()