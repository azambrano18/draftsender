import os
import sys
import time
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import psutil

from logger_utils import configurar_logger
from outlook_utils import obtener_cuentas_activas
from ejecutores import ejecutar_script
from archivos import cargar_excel, cargar_docx
from actualizacion import verificar_actualizacion
from progreso_utils import callback_progreso_gui
from envios import lanzar_envio_desde_gui

logger = configurar_logger("gui")

class DraftSenderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DraftSender - Automatización de Borradores y Envíos")
        self.root.geometry("480x450")

        # Variables visibles y reales
        self.ruta_excel_var = tk.StringVar()
        self.ruta_excel_path = tk.StringVar()
        self.ruta_docx_var = tk.StringVar()
        self.ruta_docx_path = tk.StringVar()
        self.label_cuenta_var = tk.StringVar()
        self.porcentaje_var = tk.StringVar(value="0%")
        self.status_var = tk.StringVar()
        self.cuenta_seleccionada = None

        self.combo_cuentas = None
        self.frame_progreso = None
        self.barra_progreso = None
        self.status_label = None
        self.boton_crear_borradores = None
        self.boton_cargar_excel = None
        self.boton_cargar_docx = None

        self.construir_gui()

        # Inicia monitoreo de Outlook
        self.monitor_outlook = threading.Thread(target=self.verificar_outlook_activo, daemon=True)
        self.monitor_outlook.start()

    def construir_gui(self):
        self.crear_menu()
        self.crear_encabezado()
        self.crear_selector_cuenta()
        self.crear_cargadores_archivos()
        self.crear_botonera()
        self.crear_barra_progreso()

    def crear_menu(self):
        menu_bar = tk.Menu(self.root)
        menu_archivo = tk.Menu(menu_bar, tearoff=0)
        menu_archivo.add_command(label="Actualizar", command=self.actualizar_aplicacion_intermedia)
        menu_archivo.add_command(label="Salir", command=self.root.quit)
        menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

        menu_ayuda = tk.Menu(menu_bar, tearoff=0)
        menu_ayuda.add_command(label="Acerca de", command=self.mostrar_acerca_de)
        menu_bar.add_cascade(label="Ayuda", menu=menu_ayuda)

        self.root.config(menu=menu_bar)

    def crear_encabezado(self):
        try:
            base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
            img_path = os.path.join(base_path, "config", "cover_1.jpg")
            img = Image.open(img_path).resize((480, 85))
            img_tk = ImageTk.PhotoImage(img)
            label_img = tk.Label(self.root, image=img_tk)
            label_img.image = img_tk
            label_img.pack(pady=10)
        except Exception as e:
            logger.error("No se pudo cargar la imagen de portada", exc_info=True)

    def crear_selector_cuenta(self):
        tk.Label(self.root, text="Selecciona una cuenta de Outlook:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10)
        opciones = ["Selecciona una cuenta..."]
        estado_inicial = "readonly" if self.outlook_esta_abierto() else "disabled"
        self.combo_cuentas = ttk.Combobox(self.root, values=opciones, state=estado_inicial, font=("Arial", 10), width=40)
        self.combo_cuentas.pack(padx=10, pady=5, anchor="w")
        self.combo_cuentas.bind("<<ComboboxSelected>>", self.seleccionar_cuenta)

        tk.Label(self.root, textvariable=self.label_cuenta_var, font=("Arial", 10)).pack(anchor="w", padx=10, pady=(0, 5))

    def seleccionar_cuenta(self, _event=None):
        seleccion = self.combo_cuentas.get()
        if not seleccion or seleccion == "Selecciona una cuenta...":
            self.cuenta_seleccionada = None
            self.label_cuenta_var.set("")
            self.deshabilitar_carga_archivos()
            return

        self.cuenta_seleccionada = seleccion
        self.label_cuenta_var.set(f"{seleccion} ✔️")
        logger.info(f"Cuenta seleccionada para generar borradores: {self.cuenta_seleccionada}")
        self.habilitar_carga_archivos()
        self.validar_estado_para_boton_borradores()

    def crear_cargadores_archivos(self):
        frame_excel = tk.Frame(self.root)
        frame_excel.pack(anchor="w", padx=10, pady=5)

        def cargar_excel_y_validar():
            cargar_excel(self.ruta_excel_var, self.ruta_excel_path)
            self.validar_estado_para_boton_borradores()

        self.boton_cargar_excel = tk.Button(frame_excel, text="Cargar Excel", command=cargar_excel_y_validar, font=("Arial", 10), state="disabled")
        self.boton_cargar_excel.pack(side="left")
        tk.Label(frame_excel, textvariable=self.ruta_excel_var, fg="green").pack(side="left", padx=10)

        frame_docx = tk.Frame(self.root)
        frame_docx.pack(anchor="w", padx=10, pady=5)

        def cargar_docx_y_validar():
            cargar_docx(self.ruta_docx_var, self.ruta_docx_path)
            self.validar_estado_para_boton_borradores()

        self.boton_cargar_docx = tk.Button(frame_docx, text="Cargar Texto Mail", command=cargar_docx_y_validar, font=("Arial", 10), state="disabled")
        self.boton_cargar_docx.pack(side="left")
        tk.Label(frame_docx, textvariable=self.ruta_docx_var, fg="green").pack(side="left", padx=10)

    def crear_botonera(self):
        frame_boton = tk.Frame(self.root)
        frame_boton.pack(anchor="w", padx=10, pady=5)

        def ejecutar_si_outlook_abierto():
            if not self.outlook_esta_abierto():
                messagebox.showwarning("Outlook cerrado", "Abre Outlook antes de generar los borradores.")
                return

            ejecutar_script(
                "crear_borradores",
                perfil=self.cuenta_seleccionada,
                ruta_excel=self.ruta_excel_path.get(),
                ruta_docx=self.ruta_docx_path.get(),
                callback_progreso=self.wrapper_callback
            )

        self.boton_crear_borradores = tk.Button(
            frame_boton,
            text="Crear Borradores",
            command=ejecutar_si_outlook_abierto,
            font=("Arial", 10),
            state="disabled"
        )
        self.boton_crear_borradores.pack(side="left")

        frame_enviar = tk.Frame(self.root)
        frame_enviar.pack(anchor="center", pady=5)
        tk.Button(frame_enviar, text="Enviar Borradores", command=self.lanzar_envio_gui, font=("Arial", 12), bg="purple", fg="white").pack()

    def crear_barra_progreso(self):
        self.frame_progreso = tk.Frame(self.root)
        self.barra_progreso = ttk.Progressbar(self.frame_progreso, length=300, mode='determinate', maximum=100)
        self.barra_progreso.pack(side="left", padx=(0, 10))

        tk.Label(self.frame_progreso, textvariable=self.porcentaje_var).pack(side="left")
        self.frame_progreso.pack()

        self.status_label = tk.Label(self.root, textvariable=self.status_var, font=("Arial", 10), fg="blue")
        self.status_label.pack(side="bottom", pady=(0, 5))
        self.status_label.pack_forget()

    def outlook_esta_abierto(self) -> bool:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and "outlook.exe" in proc.info['name'].lower():
                return True
        return False

    def cargar_cuentas_outlook(self):
        try:
            cuentas = obtener_cuentas_activas()
            if cuentas:
                opciones = ["Selecciona una cuenta..."] + cuentas
                self.combo_cuentas["values"] = opciones
                self.combo_cuentas.current(0)
        except Exception as e:
            logger.error(f"No se pudieron cargar las cuentas de Outlook: {e}")

    def verificar_outlook_activo(self):
        was_open = False
        while True:
            outlook_abierto = self.outlook_esta_abierto()

            if outlook_abierto and not was_open:
                self.root.after(0, self.cargar_cuentas_outlook)
                was_open = True
            elif not outlook_abierto and was_open:
                was_open = False

            if self.combo_cuentas:
                nuevo_estado = "readonly" if outlook_abierto else "disabled"
                self.root.after(0, lambda: self.combo_cuentas.config(state=nuevo_estado))

            if self.cuenta_seleccionada and outlook_abierto:
                self.root.after(0, self.habilitar_carga_archivos)
            else:
                self.root.after(0, self.deshabilitar_carga_archivos)

            time.sleep(5)

    def deshabilitar_carga_archivos(self):
        for widget in [self.boton_cargar_excel, self.boton_cargar_docx]:
            widget.config(state="disabled")

    def habilitar_carga_archivos(self):
        for widget in [self.boton_cargar_excel, self.boton_cargar_docx]:
            widget.config(state="normal")

    def actualizar_aplicacion_intermedia(self):
        verificar_actualizacion(self.root, self.barra_progreso, self.porcentaje_var, self.frame_progreso, self.status_label, self.status_var)

    def mostrar_acerca_de(self):
        messagebox.showinfo("Acerca de", "DraftSender - Automatización de borradores y envíos de correos.")

    def lanzar_envio_gui(self):
        if not self.cuenta_seleccionada:
            messagebox.showerror("Error", "Debe seleccionar una cuenta antes de enviar borradores.")
            logger.error("No se ha seleccionado una cuenta para envío.")
            return
        lanzar_envio_desde_gui(self.cuenta_seleccionada)

    def validar_estado_para_boton_borradores(self):
        archivos_cargados = self.ruta_excel_path.get() and self.ruta_docx_path.get()
        cuenta_valida = self.cuenta_seleccionada and self.cuenta_seleccionada != ""
        if archivos_cargados and cuenta_valida:
            self.boton_crear_borradores.config(state="normal")
        else:
            self.boton_crear_borradores.config(state="disabled")

    def wrapper_callback(self, fila_actual, total_filas):
        callback_progreso_gui(
            fila_actual,
            total_filas,
            self.barra_progreso,
            self.porcentaje_var,
            self.status_var,
            self.status_label,
            self.frame_progreso
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = DraftSenderApp(root)
    root.mainloop()