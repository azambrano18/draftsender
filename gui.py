import os
import sys
import tkinter as tk  # Importación de Tkinter para la interfaz gráfica
from tkinter import ttk, messagebox  # Importación de widgets adicionales para la GUI
import subprocess  # Para ejecutar scripts externos
from PIL import Image, ImageTk  # Importación para trabajar con imágenes

from logger_utils import configurar_logger  # Función para configurar el logger
logger = configurar_logger("nombre_modulo")

from outlook_utils import obtener_perfiles_outlook, cerrar_outlook, iniciar_outlook_con_perfil, obtener_cuentas_activas
from ejecutores import ejecutar_script
from archivos import cargar_excel, cargar_docx
from actualizacion import verificar_actualizacion
import estado

# Widgets globales
combo_cuentas = None
combo_cuentas_asociadas = None  # Aseguramos que esta variable esté inicializada correctamente
ruta_excel_var = None
ruta_docx_var = None
label_cuenta_var = None  # Aseguramos que esta variable se inicialice correctamente
barra_progreso = None
porcentaje_var = None
frame_progreso = None

def construir_gui(root):
    """
    Función que construye la interfaz gráfica de usuario (GUI) para la aplicación DraftSender.
    Configura los widgets de la ventana, botones y entradas necesarias.
    """
    global combo_cuentas, combo_cuentas_asociadas, ruta_excel_var, ruta_docx_var, label_cuenta_var  # Aseguramos que estas variables globales se usen correctamente

    # Inicializamos label_cuenta_var correctamente
    label_cuenta_var = tk.StringVar()

    # Barra de Menú
    menu_bar = tk.Menu(root)

    # Menú "Archivo"
    menu_archivo = tk.Menu(menu_bar, tearoff=0)
    menu_archivo.add_command(label="Actualizar", command=actualizar_aplicacion)
    menu_archivo.add_command(label="Salir", command=root.quit)
    menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

    # Menú "Ayuda"
    menu_ayuda = tk.Menu(menu_bar, tearoff=0)
    menu_ayuda.add_command(label="Acerca de", command=mostrar_acerca_de)
    menu_bar.add_cascade(label="Ayuda", menu=menu_ayuda)

    # Configura la barra de menú en la ventana principal
    root.config(menu=menu_bar)

    # Cargar la imagen del cover
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))  # Ruta base para recursos
        img_path = os.path.join(base_path, "config", "cover_borradores.jpg")  # Ruta completa de la imagen
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
    tk.Label(root, textvariable=label_cuenta_var, font=("Arial", 10)).pack(pady=5)

    # Inicialización de combo_cuentas_asociadas aquí
    combo_cuentas_asociadas = ttk.Combobox(root, state="readonly", font=("Arial", 10))
    combo_cuentas_asociadas.bind("<<ComboboxSelected>>", lambda e: seleccionar_cuenta_asociada())
    combo_cuentas_asociadas.pack_forget()  # Este es el cambio, aseguramos que esté en la función

    # Inicialización de las variables de ruta (Excel y DOCX)
    ruta_excel_var = tk.StringVar()
    ruta_docx_var = tk.StringVar()

    # Excel
    frame_excel = tk.Frame(root)
    frame_excel.pack(anchor="w", padx=10, pady=5)
    tk.Button(frame_excel, text="Cargar Excel", command=lambda: cargar_excel(ruta_excel_var), font=("Arial", 10)).pack(side="left")
    tk.Label(frame_excel, textvariable=ruta_excel_var, fg="green").pack(side="left", padx=10)

    # Word
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


def actualizar_aplicacion():
    """
    Ejecuta el script de actualización (actualizacion.py) para forzar la actualización.
    """
    try:
        # Ejecutar el script de actualización
        script_path = os.path.join(os.path.dirname(__file__), "actualizacion.py")
        subprocess.Popen([sys.executable, script_path])  # Ejecuta el script de actualización
        messagebox.showinfo("Actualización", "El proceso de actualización ha comenzado.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar la actualización: {e}")


def mostrar_acerca_de():
    """
    Muestra un mensaje con información acerca de la aplicación.
    """
    messagebox.showinfo("Acerca de", "DraftSender - Automatización de borradores y envíos de correos.")


def mostrar_cuenta_seleccionada(_event=None):
    """
    Función que muestra la cuenta seleccionada en el combobox de perfiles de Outlook.
    Configura la cuenta asociada y las rutas de los archivos cargados.
    """
    perfil = combo_cuentas.get()
    estado.cuenta_seleccionada = None
    ruta_excel_var.set("")  # Ahora esto no da error, ya que ruta_excel_var está inicializado
    ruta_docx_var.set("")  # Lo mismo para esta variable
    combo_cuentas_asociadas.pack_forget()  # Asegurándonos de que esté inicializado
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


def seleccionar_cuenta_asociada():
    """
    Función que asigna la cuenta seleccionada en el combobox de cuentas asociadas
    a la variable global 'estado.cuenta_seleccionada'.
    """
    estado.cuenta_seleccionada = combo_cuentas_asociadas.get()


def lanzar_envio_gui():
    """
    Función que lanza el envío de borradores mediante la interfaz de envío de 'envios.py'.
    Verifica si se ha seleccionado una cuenta antes de proceder.
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


# Código para inicializar la GUI
if __name__ == "__main__":
    root = tk.Tk()
    root.title("DraftSender - Automatización de Borradores y Envíos")
    root.geometry("480x450")

    # Llamar a la función que construye la GUI
    construir_gui(root)

    root.mainloop()