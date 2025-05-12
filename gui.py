import os
import sys
import tkinter as tk  # Importación de Tkinter para la interfaz gráfica
from tkinter import ttk, messagebox  # Importación de widgets adicionales para la GUI
import subprocess  # Para ejecutar scripts externos

from logger_utils import configurar_logger  # Función para configurar el logger
logger = configurar_logger("nombre_modulo")

from outlook_utils import obtener_perfiles_outlook, cerrar_outlook, iniciar_outlook_con_perfil, obtener_cuentas_activas
from ejecutores import ejecutar_script
from archivos import cargar_excel, cargar_docx
from actualizacion import verificar_actualizacion
import estado

# Widgets globales
combo_cuentas = None
combo_cuentas_asociadas = None
ruta_excel_var = None
ruta_docx_var = None
label_cuenta_var = None
barra_progreso = None
porcentaje_var = None
frame_progreso = None

def seleccionar_cuenta_asociada():
    """
    Función que asigna la cuenta seleccionada en el combobox de cuentas asociadas
    a la variable global 'estado.cuenta_seleccionada'.

    Args:
        None

    Returns:
        None
    """
    estado.cuenta_seleccionada = combo_cuentas_asociadas.get()

def mostrar_cuenta_seleccionada(_event=None):
    """
    Función que muestra la cuenta seleccionada en el combobox de perfiles de Outlook.
    Configura la cuenta asociada y las rutas de los archivos cargados.

    Args:
        _event: Evento que dispara la función al seleccionar un perfil de Outlook.

    Returns:
        None
    """
    perfil = combo_cuentas.get()
    estado.cuenta_seleccionada = None
    ruta_excel_var.set("")
    ruta_docx_var.set("")
    combo_cuentas_asociadas.pack_forget()
    combo_cuentas_asociadas.set("")

    if perfil == "Seleccione perfil...":
        label_cuenta_var.set("")
        return

    cerrar_outlook()  # Cierra Outlook si está abierto
    iniciar_outlook_con_perfil(perfil)  # Inicia Outlook con el perfil seleccionado
    cuentas = obtener_cuentas_activas()  # Obtiene las cuentas activas de Outlook

    if not cuentas:
        logger.error("No se encontraron cuentas.")
        label_cuenta_var.set("No se encontraron cuentas.")
        return
    elif len(cuentas) == 1:
        estado.cuenta_seleccionada = cuentas[0]
        label_cuenta_var.set(f"{estado.cuenta_seleccionada} ✔️")
    else:
        label_cuenta_var.set("Selecciona una cuenta:")
        combo_cuentas_asociadas['values'] = cuentas
        combo_cuentas_asociadas.current(0)
        combo_cuentas_asociadas.pack()

def lanzar_envio_gui():
    """
    Función que lanza el envío de borradores mediante la interfaz de envío de 'envios.py'.
    Verifica si se ha seleccionado una cuenta antes de proceder.

    Args:
        None

    Returns:
        None
    """
    try:
        if not estado.cuenta_seleccionada:
            messagebox.showerror("Error", "Debe seleccionar una cuenta antes de enviar borradores.")
            logger.error("No se ha seleccionado una cuenta para envío.")
            return

        ruta_script = os.path.join(os.path.dirname(__file__), "envios.py")
        subprocess.Popen([sys.executable, ruta_script, estado.cuenta_seleccionada])  # Ejecuta el script de envíos
    except Exception as e:
        logger.exception("No se pudo abrir la interfaz de envío")
        messagebox.showerror("Error", f"No se pudo abrir la interfaz de envío:\n{e}")

def construir_gui(root):
    """
    Función que construye la interfaz gráfica de usuario (GUI) para la aplicación DraftSender.
    Configura los widgets de la ventana, botones y entradas necesarias.

    Args:
        root (tk.Tk): Ventana principal de la aplicación.

    Returns:
        None
    """
    global combo_cuentas, combo_cuentas_asociadas
    global ruta_excel_var, ruta_docx_var, label_cuenta_var
    global barra_progreso, porcentaje_var, frame_progreso

    try:
        logger.info("Construyendo GUI de DraftSender")
        from PIL import Image, ImageTk  # Importación de PIL para manejo de imágenes
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))  # Ruta base para recursos
        img_path = os.path.join(base_path, "config", "cover_borradores.jpg")
        img = Image.open(img_path).resize((500, 90))  # Carga y redimensiona la imagen
        img_tk = ImageTk.PhotoImage(img)
        label_img = tk.Label(root, image=img_tk)
        label_img.image = img_tk  # Mantiene la referencia de la imagen
        label_img.pack(pady=10)
    except Exception as e:
        logger.error("No se pudo cargar la imagen de portada", exc_info=True)

    # Perfil de Outlook
    perfiles = obtener_perfiles_outlook()  # Obtiene los perfiles de Outlook
    max_length = max(len(p) for p in perfiles)  # Define el tamaño máximo para el combobox
    tk.Label(root, text="Selecciona un perfil:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10)

    frame_perfiles = tk.Frame(root)
    frame_perfiles.pack(anchor="w", padx=10, pady=5)
    combo_cuentas = ttk.Combobox(frame_perfiles, values=perfiles, state="readonly", font=("Arial", 10), width=max_length + 2)
    combo_cuentas.pack(side="left")
    combo_cuentas.bind("<<ComboboxSelected>>", mostrar_cuenta_seleccionada)
    combo_cuentas.current(0)

    # Cuenta asociada
    label_cuenta_var = tk.StringVar()
    tk.Label(root, textvariable=label_cuenta_var, font=("Arial", 10)).pack(pady=5)

    combo_cuentas_asociadas = ttk.Combobox(root, state="readonly", font=("Arial", 10))
    combo_cuentas_asociadas.bind("<<ComboboxSelected>>", lambda e: seleccionar_cuenta_asociada())
    combo_cuentas_asociadas.pack_forget()

    # Excel
    ruta_excel_var = tk.StringVar()
    frame_excel = tk.Frame(root)
    frame_excel.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_excel, text="Cargar Excel", command=lambda: cargar_excel(ruta_excel_var), font=("Arial", 10)).pack(side="left")
    tk.Label(frame_excel, textvariable=ruta_excel_var, fg="green").pack(side="left", padx=10)

    # Word
    ruta_docx_var = tk.StringVar()
    frame_docx = tk.Frame(root)
    frame_docx.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_docx, text="Cargar Texto Mail", command=lambda: cargar_docx(ruta_docx_var),
              font=("Arial", 10)).pack(side="left")
    tk.Label(frame_docx, textvariable=ruta_docx_var, fg="green").pack(side="left", padx=10)

    # Crear Borradores
    ruta_boton_var = tk.StringVar()
    frame_boton = tk.Frame(root)
    frame_boton.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_boton, text="Crear Borradores",
              command=lambda: ejecutar_script("crear_borradores", combo_cuentas.get()),
              font=("Arial", 10)).pack(side="left")

    # Botón "Enviar Borradores" (abajo y centrado)
    frame_enviar = tk.Frame(root)
    frame_enviar.pack(anchor="center", pady=20)
    tk.Button(frame_enviar, text="Enviar Borradores",
              command=lanzar_envio_gui,
              font=("Arial", 12), bg="purple", fg="white").pack()

    # Barra de progreso
    frame_progreso = tk.Frame(root)
    frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
    barra_progreso = ttk.Progressbar(frame_progreso, length=400, mode='determinate', maximum=100)
    barra_progreso.pack(side="left", padx=(0, 10))
    porcentaje_var = tk.StringVar(value="0%")
    tk.Label(frame_progreso, textvariable=porcentaje_var).pack(side="left")