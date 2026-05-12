import time
import tkinter as tk
import psutil
import re
import logging
import os
import sys
import threading

from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from draftsender_app.envios_ui import lanzar_envio_desde_gui
from draftsender_app.archivos import cargar_excel, cargar_docx, validar_excel_gui
from draftsender_app.outlook_utils import obtener_cuentas_activas, obtener_perfiles_outlook, cerrar_outlook, iniciar_outlook_con_perfil
from draftsender_app.actualizacion import ejecutar_actualizacion, hay_nueva_version_disponible

logger = logging.getLogger("DraftSender")

class DraftSenderApp:
    """
        Inicializa la aplicación principal DraftSender

        Args:
            root (tk.Tk): Ventana raíz de Tkinter.
            titulo_ventana (str): Título de la ventana principal, puede incluir el nombre de usuario.

        Inicializa variables, crea la interfaz y lanza el monitor de Outlook en segundo plano.
        """
    def __init__(self, root: tk.Tk, *, titulo_ventana: str = "DraftSender"):
        self.root = root
        self.nombre_usuario = titulo_ventana.split(" - ")[-1].strip()

        # gui.py > configuración del título de ventana
        titulo_actual = self.root.title()

        partes_titulo = titulo_actual.split(" - ", 1)

        if len(partes_titulo) == 2:
            nombre_app, nombre_usuario = partes_titulo
            self.root.title(f"{nombre_app} - Creador de Borradores - {nombre_usuario}")
        else:
            self.root.title(f"{titulo_actual} - Creador de Borradores")

        self.root.geometry("600x400")
        self.tamano_inicial = self.root.winfo_geometry()

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

        self.monitor_outlook = threading.Thread(target=self.verificar_outlook_activo, daemon=True)
        self.monitor_outlook.start()

    def ocultar_barra_progreso(self):
        """
           Oculta la barra de progreso y el estado visual de la interfaz.
           Restaura el tamaño inicial de la ventana principal.
           """
        if self.frame_progreso:
            self.frame_progreso.pack_forget()
        if self.status_label:
            self.status_label.pack_forget()

        self.root.update_idletasks()
        self.root.geometry("600x400")

    def construir_gui(self):
        """
            Construye todos los componentes de la interfaz gráfica principal.

            Llama internamente a métodos de construcción modular como:
            - Menú superior
            - Encabezado
            - Selector de cuenta
            - Métodos de envío
            - Cargadores de archivos
            - Botonera de acciones
            - Barra de progreso
            """
        self.crear_menu()
        self.crear_encabezado()
        self.crear_selector_cuenta()
        self.crear_selector_metodo_envio()
        self.crear_cargadores_archivos()
        self.crear_botonera()
        self.crear_barra_progreso()

    def crear_menu(self):
        """
        gui.py > crear_menu
        Crea la barra de menú principal de la aplicación con las opciones:
        - Archivo: Actualizar, Salir
        - Ayuda: Ver instructivo, Acerca de

        Además verifica en segundo plano si existe una nueva versión.
        Si hay actualización disponible:
        - habilita el botón Actualizar
        - muestra un aviso al usuario
        """
        menu_bar = tk.Menu(self.root)

        menu_archivo = tk.Menu(menu_bar, tearoff=0)
        menu_archivo.add_command(
            label="Actualizar",
            command=self.actualizar_aplicacion_intermedia
        )
        menu_archivo.add_command(label="Salir", command=self.root.quit)
        menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

        menu_ayuda = tk.Menu(menu_bar, tearoff=0)
        menu_ayuda.add_command(label="Ver Instructivo", command=self.mostrar_readme)
        menu_ayuda.add_command(label="Acerca de", command=self.mostrar_acerca_de)
        menu_bar.add_cascade(label="Ayuda", menu=menu_ayuda)

        self.root.config(menu=menu_bar)

        self.menu_archivo = menu_archivo
        self.menu_archivo.entryconfig("Actualizar", state="disabled")

        self.root.after(3000, self.verificar_actualizacion_al_inicio)

    def verificar_actualizacion_al_inicio(self):
        """
        gui.py > verificar_actualizacion_al_inicio
        Verifica si hay una nueva versión disponible al iniciar la app.

        Si existe:
        - habilita Archivo > Actualizar
        - muestra aviso al usuario
        """

        def tarea():
            try:
                hay_actualizacion = hay_nueva_version_disponible()

                if not hay_actualizacion:
                    return

                def avisar():
                    try:
                        self.menu_archivo.entryconfig("Actualizar", state="normal")

                        respuesta = messagebox.askyesno(
                            "Nueva versión disponible",
                            "Hay una nueva versión de DraftSender disponible.\n\n"
                            "¿Quieres actualizar ahora?"
                        )

                        if respuesta:
                            self.actualizar_aplicacion_intermedia()

                    except Exception as e:
                        logger.error(f"Error mostrando aviso de actualización: {e}")

                self.root.after(0, avisar)

            except Exception as e:
                logger.warning(f"No se pudo verificar actualización al inicio: {e}")

        threading.Thread(target=tarea, daemon=True).start()

    def mostrar_readme(self):
        """
        Muestra el archivo README.txt en una ventana emergente de lectura.
        El instructivo se carga desde C:\DraftSender_app\data\README.txt.
        """
        try:
            # Ruta fija
            ruta_readme = r"C:\DraftSender_app\data\README.txt"

            if not os.path.exists(ruta_readme):
                messagebox.showwarning(
                    "Instructivo no encontrado",
                    f"No se pudo localizar el archivo README.txt en:\n{ruta_readme}"
                )
                return

            with open(ruta_readme, "r", encoding="utf-8") as f:
                contenido = f.read()

            ventana = tk.Toplevel(self.root)
            ventana.title("Instructivo de uso - DraftSender")
            ventana.geometry("700x500")

            texto = tk.Text(ventana, wrap="word", font=("Consolas", 10))
            texto.insert("1.0", contenido)
            texto.configure(state="disabled")
            texto.pack(expand=True, fill="both")

            scrollbar = tk.Scrollbar(ventana, command=texto.yview)
            texto.config(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")

        except Exception as e:
            logger.error(f"Error al mostrar README.txt: {e}")
            messagebox.showerror(
                "Error",
                f"No se pudo mostrar el instructivo:\n{e}"
            )

    def crear_encabezado(self):
        """
            Carga y muestra una imagen de portada en la parte superior de la ventana.
            La imagen se carga desde la carpeta 'config' y se ajusta al ancho de la ventana.
            """
        try:
            base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
            img_path = os.path.join(base_path,"config", "cover.jpg")
            img = Image.open(img_path).resize((600, 95))
            img_tk = ImageTk.PhotoImage(img)
            label_img = tk.Label(self.root, image=img_tk)
            label_img.image = img_tk
            label_img.pack(pady=10)
        except Exception as e:
            logger.error("No se pudo cargar la imagen de portada", exc_info=True)

    def crear_selector_metodo_envio(self):
        """
            Crea un selector (combobox) para elegir el metodo de envío (Ej: Envíos 1 o Seguimiento).
            Este selector depende del frame donde se cargan las cuentas Outlook.

            Habilita trazabilidad del valor seleccionado para activar campos relevantes.
            """
        try:
            frame_selector = self.root.nametowidget(self.combo_cuentas.winfo_parent())

            tk.Label(frame_selector,text="Método de envío:",font=("Arial", 10)).grid(row=0, column=2, sticky="w", padx=(10, 0))

            self.metodo_envio_var = tk.StringVar()
            self.metodo_envio_var.trace_add("write", self.log_metodo_seleccionado)  # Conecta trazabilidad

            opciones = ["Selecciona método", "Envíos 1", "Reenviados 2", "Reenviados 3"]
            self.combo_metodo = ttk.Combobox(frame_selector,textvariable=self.metodo_envio_var,values=opciones,font=("Arial", 10),state="disabled",width=20)
            self.combo_metodo.grid(row=1, column=2, sticky="w", padx=(10, 0))
            self.combo_metodo.current(0)

        except Exception as e:
            logger.error("Error al crear selector de método de envío", exc_info=True)

    def log_metodo_seleccionado(self, *_args):
        metodo = self.metodo_envio_var.get()

        if metodo != "Selecciona método":
            logger.info(f"Método de envío: {metodo}")

            if self.boton_cargar_excel and self.boton_cargar_excel.winfo_exists():
                self.boton_cargar_excel.config(state="normal")
            if self.boton_cargar_docx and self.boton_cargar_docx.winfo_exists():
                self.boton_cargar_docx.config(state="normal")
        else:
            if self.boton_cargar_excel and self.boton_cargar_excel.winfo_exists():
                self.boton_cargar_excel.config(state="disabled")
            if self.boton_cargar_docx and self.boton_cargar_docx.winfo_exists():
                self.boton_cargar_docx.config(state="disabled")

        self.validar_estado_para_boton_borradores()

    def crear_selector_cuenta(self):
        """
            Crea un panel con dos combobox:
            - Uno para seleccionar el perfil de Outlook.
            - Otro para seleccionar la cuenta vinculada a ese perfil.

            También agrega una etiqueta para mostrar el estado de la cuenta seleccionada.
            """
        frame_selector = tk.Frame(self.root)
        frame_selector.pack(padx=10, pady=5, fill="x")

        tk.Label(frame_selector, text="Perfil de Outlook:", font=("Arial", 10)).grid(row=0, column=0, sticky="w")
        perfiles = obtener_perfiles_outlook()
        opciones_perfil = ["Selecciona un perfil..."] + perfiles
        self.combo_perfiles = ttk.Combobox(frame_selector, values=opciones_perfil, state="readonly", font=("Arial", 10), width=22)
        self.combo_perfiles.grid(row=1, column=0, padx=(0, 10), sticky="w")
        self.combo_perfiles.current(0)
        self.combo_perfiles.bind("<<ComboboxSelected>>", self.actualizar_cuentas_por_perfil)

        tk.Label(frame_selector, text="Cuenta vinculada:", font=("Arial", 10)).grid(row=0, column=1, sticky="w")
        self.combo_cuentas = ttk.Combobox(frame_selector, values=["Selecciona una cuenta..."], font=("Arial", 10),width=28, state="disabled")
        self.combo_cuentas.grid(row=1, column=1, sticky="w")
        self.combo_cuentas.current(0)
        self.combo_cuentas.bind("<<ComboboxSelected>>", self.seleccionar_cuenta)

        self.label_cuenta_var.set("")
        tk.Label(self.root, textvariable=self.label_cuenta_var, font=("Arial", 10)).pack(anchor="w", padx=10, pady=(0, 5))

    def seleccionar_cuenta(self, _event=None):
        """
            Maneja la selección de una cuenta de correo desde el combobox.
            Valida el formato del correo, actualiza el estado de la interfaz y habilita el selector de metodo si es válido.

            Args:
                _event: Evento del sistema (no usado directamente).
            """
        seleccion = self.combo_cuentas.get()

        if not seleccion or seleccion == "Selecciona una cuenta...":
            self.cuenta_seleccionada = None
            self.label_cuenta_var.set("")
            self.combo_metodo.set("Selecciona método")
            self.combo_metodo.config(state="disabled")
            self.validar_estado_para_boton_borradores()
            return

        if not es_email_valido(seleccion):
            self.cuenta_seleccionada = None
            self.label_cuenta_var.set("El formato del correo seleccionado no es válido.")
            logger.warning(f"Cuenta con formato inválido seleccionada: {seleccion}")
            self.combo_metodo.set("Selecciona método")
            self.combo_metodo.config(state="disabled")
            self.validar_estado_para_boton_borradores()
            return

        self.cuenta_seleccionada = seleccion
        self.label_cuenta_var.set(f"Cuenta seleccionada: {seleccion}")
        logger.info(f"Cuenta seleccionada: {seleccion}")

        self.combo_metodo.config(state="readonly")
        self.combo_metodo.set("Selecciona método")
        self.validar_estado_para_boton_borradores()

    def crear_cargadores_archivos(self):
        """
        Crea los botones y etiquetas necesarias para cargar:
        - El archivo Excel con los destinatarios.
        - El archivo Word con el cuerpo del correo.

        Ambos botones están inicialmente deshabilitados y se activan según el metodo seleccionado.
        Se realiza validación inmediata de estructura de Excel y etiquetas del Word.
        """
        frame_excel = tk.Frame(self.root)
        frame_excel.pack(anchor="w", padx=10, pady=5)

        self.boton_cargar_excel = tk.Button(
            frame_excel,
            text="Cargar Excel",
            command=lambda: (validar_excel_gui(self.ruta_excel_path.get())
                                if cargar_excel(self.ruta_excel_var, self.ruta_excel_path)
                                else False) or self.actualizar_estado_post_carga(),
            font=("Arial", 10),
            state="disabled"
        )
        self.boton_cargar_excel.pack(side="left")
        tk.Label(frame_excel, textvariable=self.ruta_excel_var, fg="green").pack(side="left", padx=10)

        frame_docx = tk.Frame(self.root)
        frame_docx.pack(anchor="w", padx=10, pady=5)

        self.boton_cargar_docx = tk.Button(
            frame_docx,
            text="Cargar Texto Mail",
            command=lambda: cargar_docx(self.ruta_docx_var, self.ruta_docx_path) or self.actualizar_estado_post_carga(),
            font=("Arial", 10),
            state="disabled"
        )

        self.boton_cargar_docx.pack(side="left")
        tk.Label(frame_docx, textvariable=self.ruta_docx_var, fg="green").pack(side="left", padx=10)

    def crear_botonera(self):
        """
            Crea los botones principales de acción:
            - "Crear Borradores": genera los correos en Outlook, requiere Outlook activo.
            - "Enviar Borradores": abre la ventana de envío programado.
            """
        frame_boton = tk.Frame(self.root)
        frame_boton.pack(anchor="w", padx=10, pady=5)

        self.boton_crear_borradores = tk.Button(frame_boton,text="Crear Borradores",command=self.ejecutar_si_outlook_abierto,font=("Arial", 10),state="disabled")
        self.boton_crear_borradores.pack(side="left")

        frame_enviar = tk.Frame(self.root)
        frame_enviar.pack(anchor="center", pady=5)
        tk.Button(frame_enviar,text="Enviar Borradores",command=self.lanzar_envio_gui,font=("Arial", 12),bg="purple",fg="white").pack()

    def crear_barra_progreso(self):
        """
        Inicializa y configura la barra de progreso y etiqueta de estado en la interfaz.
        Se utiliza para mostrar el avance durante la creación de borradores.
        """
        self.frame_progreso = tk.Frame(self.root)

        self.barra_progreso = ttk.Progressbar(self.frame_progreso,length=300,mode='determinate',maximum=100)
        self.barra_progreso.pack(pady=(0, 2))

        self.frame_progreso.pack_forget()

        self.status_label = tk.Label(self.root, textvariable=self.status_var, font=("Arial", 10), fg="blue")
        self.status_label.pack(side="bottom", pady=(0, 5))
        self.status_label.pack_forget()

    def outlook_esta_abierto(self) -> bool:
        """
        Verifica si el proceso de Outlook está actualmente activo en el sistema operativo.

        Returns:
            bool: True si Outlook está abierto, False si no.
        """
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and "outlook.exe" in proc.info['name'].lower():
                return True
        return False

    def ejecutar_si_outlook_abierto(self):
        """
        Ejecuta la creación de borradores solo si Outlook está abierto.
        Si no lo está, muestra un mensaje de advertencia al usuario.
        """
        if not self.outlook_esta_abierto():
            messagebox.showwarning("Outlook cerrado", "Abre Outlook antes de generar los borradores.")
            return

        if not self.validar_entradas_para_crear_borradores():
            return

        threading.Thread(target=self._crear_borradores_en_hilo, daemon=True).start()

    def validar_entradas_para_crear_borradores(self) -> bool:
        """
        Verifica que todos los elementos necesarios estén definidos y correctamente configurados
        antes de permitir la creación de borradores.
        """
        cuenta = self.cuenta_seleccionada
        excel_path = self.ruta_excel_path.get()
        docx_path = self.ruta_docx_path.get()
        metodo = self.metodo_envio_var.get()

        if not cuenta or not es_email_valido(cuenta):
            messagebox.showerror("Cuenta inválida", "Debes seleccionar una cuenta de Outlook válida.")
            return False

        if metodo == "Selecciona método":
            messagebox.showerror("Método requerido", "Debes seleccionar el método de envío (Envíos 1 o Seguimiento).")
            return False

        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Archivo Excel faltante", "Debes cargar un archivo Excel válido.")
            return False

        if not docx_path or not os.path.exists(docx_path):
            messagebox.showerror("Archivo Word faltante", "Debes cargar un archivo Word válido.")
            return False

        if not validar_excel_gui(excel_path):
            return False

        return True

    def _crear_borradores_en_hilo(self):
        """
        gui.py > _crear_borradores_en_hilo
        Ejecuta el proceso de creación de borradores sin bloquear la interfaz.
        Usa el perfil real seleccionado, no la cuenta de correo.
        """
        from draftsender_app.borradores import generar_borradores

        try:
            metodo = self.metodo_envio_var.get()
            perfil_outlook = self.combo_perfiles.get().strip() if self.combo_perfiles else ""

            if metodo == "Selecciona método":
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "Método no seleccionado",
                        "Debes seleccionar un método de creación de borradores.",
                    ),
                )
                self.root.after(0, lambda: self.status_var.set(""))
                return

            self.root.after(0, lambda: self.status_var.set("Creando borradores..."))
            self.root.after(0, lambda: self.frame_progreso.pack(pady=(5, 5)))
            self.root.after(0, self.status_label.pack)

            enviados, errores, mensaje = generar_borradores(
                cuenta=self.cuenta_seleccionada,
                perfil=perfil_outlook,
                ruta_excel=self.ruta_excel_path.get(),
                ruta_docx=self.ruta_docx_path.get(),
                modo_envio=metodo,
                callback_progreso=self.wrapper_callback,
            )

            self.root.after(0, lambda: self.status_var.set("Proceso completado"))
            self.root.after(8000, self.ocultar_barra_progreso)

            if errores:
                self.root.after(
                    0,
                    lambda: messagebox.showwarning("Errores en el proceso", mensaje),
                )
            else:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo("Borradores creados", mensaje),
                )

        except Exception as e:
            logger.exception("Error al crear borradores")
            self.root.after(
                0,
                lambda: self.status_var.set("Ocurrió un error al crear los borradores."),
            )
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "Error",
                    f"Ocurrió un error inesperado:\n{e}",
                ),
            )

    def wrapper_callback(self, fila_actual, total_filas):
        """
        gui.py > wrapper_callback
        Actualiza progreso de forma segura desde el hilo principal de Tkinter.
        """

        def actualizar_ui():
            if total_filas <= 0:
                return

            porcentaje = int((fila_actual / total_filas) * 100)

            self.barra_progreso["value"] = porcentaje
            self.porcentaje_var.set(f"{porcentaje}%")
            self.status_var.set(
                f"{porcentaje}% - Procesando fila {fila_actual} de {total_filas}"
            )

            if fila_actual == 1:
                self.frame_progreso.pack(pady=(5, 5))
                self.status_label.pack()
                self.root.geometry("")

            if fila_actual == total_filas:
                self.status_var.set("¡Borradores creados con éxito!")
                self.root.after(8000, self.ocultar_barra_progreso)

        self.root.after(0, actualizar_ui)

    def lanzar_envio_gui(self):
        """
        Lanza una nueva ventana que permite enviar los borradores previamente creados.
        Utiliza el módulo envios_ui.py para gestionar esta interfaz secundaria.
        """
        try:
            if not self.outlook_esta_abierto():
                messagebox.showwarning("Outlook cerrado", "Debes tener Outlook abierto para enviar los correos.")
                return

            lanzar_envio_desde_gui(self.nombre_usuario)

        except Exception as e:
            logger.exception("Error al lanzar ventana de envíos automáticos")
            messagebox.showerror("Error", f"No se pudo abrir la ventana de envíos automático:\n{e}")

    def cargar_cuentas_outlook(self):
        """
        Carga las cuentas disponibles en Outlook para el perfil seleccionado.
        Llena el combobox de cuentas y habilita su selección si hay cuentas válidas.
        """
        try:
            cuentas = obtener_cuentas_activas()
            if cuentas:
                opciones = ["Selecciona una cuenta..."] + cuentas
                self.combo_cuentas["values"] = opciones
                self.combo_cuentas.current(0)
        except Exception as e:
            logger.error(f"No se pudieron cargar las cuentas de Outlook: {e}")

    def verificar_outlook_activo(self):
        """
        Monitorea en segundo plano si Outlook está activo.
        Habilita o deshabilita la creación de borradores según su disponibilidad.
        """
        was_open = False
        while True:
            outlook_abierto = self.outlook_esta_abierto()

            try:
                if not self.root or not self.root.winfo_exists():
                    break

                if outlook_abierto and not was_open:
                    self.root.after(0, self.cargar_cuentas_outlook)
                    was_open = True
                elif not outlook_abierto and was_open:
                    was_open = False

                if self.combo_cuentas:
                    nuevo_estado = "readonly" if outlook_abierto else "disabled"
                    self.root.after(0, lambda: self._configurar_combo_cuentas_seguro(nuevo_estado))

                self.root.after(0, self.validar_estado_para_boton_borradores)

                time.sleep(5)

            except Exception:
                break

    def _configurar_combo_cuentas_seguro(self, nuevo_estado):
        """
        Configura el combobox de cuentas en un entorno seguro (thread-safe).
        Se asegura de que el cambio de estado o contenido del widget no cause errores de concurrencia.
        """
        try:
            if self.root and self.root.winfo_exists() and self.combo_cuentas and self.combo_cuentas.winfo_exists():
                self.combo_cuentas.config(state=nuevo_estado)
        except Exception:
            pass

    def actualizar_aplicacion_intermedia(self):
        """
        gui.py > actualizar_aplicacion_intermedia
        Ejecuta la actualización manual desde el menú Archivo > Actualizar.
        """

        def tarea():
            try:
                self.root.after(
                    0,
                    lambda: self.status_var.set("Descargando actualización...")
                )

                actualizado = ejecutar_actualizacion(forzar=True)

                if not actualizado:
                    self.root.after(
                        0,
                        lambda: messagebox.showinfo(
                            "Sin actualizaciones",
                            "Ya tienes instalada la última versión disponible."
                        )
                    )
                    return

                def cerrar_app():
                    messagebox.showinfo(
                        "Actualización en proceso",
                        "La actualización fue descargada.\n\n"
                        "DraftSender se cerrará y se abrirá la nueva versión."
                    )

                    try:
                        self.root.quit()
                    finally:
                        self.root.destroy()

                self.root.after(0, cerrar_app)

            except Exception as e:
                logger.exception("Error al ejecutar actualización")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error de actualización",
                        f"No se pudo ejecutar la actualización:\n{e}"
                    )
                )

            finally:
                self.root.after(0, lambda: self.status_var.set(""))

        threading.Thread(target=tarea, daemon=True).start()

    def mostrar_acerca_de(self):
        """
        Muestra una ventana emergente con información sobre la aplicación y su versión.
        """
        messagebox.showinfo("Acerca de", "DraftSender - Automatización de borradores y envíos de correos.")

    def validar_estado_para_boton_borradores(self):
        """
        gui.py > validar_estado_para_boton_borradores
        Habilita Crear Borradores solo si hay cuenta, método y archivos válidos.
        """
        if not self.boton_crear_borradores:
            return

        cuenta_ok = bool(
            self.cuenta_seleccionada
            and es_email_valido(self.cuenta_seleccionada)
        )

        metodo = (
            self.metodo_envio_var.get()
            if hasattr(self, "metodo_envio_var")
            else "Selecciona método"
        )

        metodo_ok = bool(metodo and metodo != "Selecciona método")

        excel_ok = bool(
            self.ruta_excel_path.get()
            and os.path.exists(self.ruta_excel_path.get())
        )

        docx_ok = bool(
            self.ruta_docx_path.get()
            and os.path.exists(self.ruta_docx_path.get())
        )

        estado = "normal" if cuenta_ok and metodo_ok and excel_ok and docx_ok else "disabled"

        try:
            self.boton_crear_borradores.config(state=estado)

        except tk.TclError as e:
            logger.warning(f"No se pudo actualizar el botón de borradores: {e}")

    def deshabilitar_carga_archivos(self):
        """
        Deshabilita los botones de carga de Excel y Word.
        Usado al cambiar de cuenta o metodo de envío.
        """
        try:
            if self.boton_cargar_excel and self.boton_cargar_excel.winfo_exists():
                self.boton_cargar_excel.config(state="disabled")
            if self.boton_cargar_docx and self.boton_cargar_docx.winfo_exists():
                self.boton_cargar_docx.config(state="disabled")
        except Exception as e:
            logger.warning(f"Error al deshabilitar carga de archivos: {e}")

    def actualizar_estado_post_carga(self):
        """
        Valida el estado actual después de cargar archivos y ajusta la disponibilidad
        del botón 'Crear Borradores' si corresponde.
        """
        self.validar_estado_para_boton_borradores()

    def actualizar_cuentas_por_perfil(self, _event=None):
        """
        Carga las cuentas asociadas al perfil de Outlook seleccionado y actualiza el combobox de cuentas.
        """
        perfil = self.combo_perfiles.get()
        if not perfil or perfil == "Selecciona un perfil...":
            self.combo_cuentas["values"] = ["Selecciona una cuenta..."]
            self.combo_cuentas.set("Selecciona una cuenta...")
            self.combo_cuentas.config(state="disabled")
            self.label_cuenta_var.set("Debes seleccionar un perfil válido.")
            self.deshabilitar_carga_archivos()
            self.cuenta_seleccionada = None
            if self.combo_metodo and self.combo_metodo.winfo_exists():
                self.combo_metodo.set("Selecciona método")
                self.combo_metodo.config(state="disabled")
            return

        self.combo_cuentas["values"] = ["Cargando cuentas..."]
        self.combo_cuentas.set("Cargando cuentas...")
        self.combo_cuentas.config(state="disabled")
        self.root.update_idletasks()

        cerrar_outlook()
        time.sleep(1)

        iniciar_outlook_con_perfil(perfil)
        time.sleep(5)

        try:
            cuentas = obtener_cuentas_activas()
            if cuentas:
                opciones = ["Selecciona una cuenta..."] + cuentas
                self.combo_cuentas["values"] = opciones
                self.combo_cuentas.set("Selecciona una cuenta...")
                self.combo_cuentas.config(state="readonly")
                self.cuenta_seleccionada = None
                self.deshabilitar_carga_archivos()
                self.combo_metodo.set("Selecciona método")
                self.combo_metodo.config(state="disabled")
            else:
                raise ValueError("No se encontraron cuentas activas.")
        except Exception as e:
            import traceback
            logger.error(f"No se pudieron obtener cuentas del perfil '{perfil}': {e}")
            logger.debug("Detalles del error:\n" + traceback.format_exc())

            self.combo_cuentas["values"] = ["Selecciona una cuenta..."]
            self.combo_cuentas.set("Selecciona una cuenta...")
            self.combo_cuentas.config(state="disabled")
            self.label_cuenta_var.set("No fue posible acceder al perfil de Outlook.")
            self.cuenta_seleccionada = None
            self.deshabilitar_carga_archivos()
            self.combo_metodo.set("Selecciona método")
            self.combo_metodo.config(state="disabled")

def es_email_valido(correo: str) -> bool:
    """
       Verifica si una dirección de correo electrónico tiene un formato válido.

       Args:
           correo (str): Dirección de correo a validar.

       Returns:
           bool: True si el correo es válido, False en caso contrario.
       """
    patron = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return bool(re.match(patron, correo))

def obtener_nombre_usuario():
    """
        Obtiene el nombre del usuario desde el archivo de configuración local.

        Returns:
            str: Nombre del usuario, o 'usuario' si no se encuentra el archivo.
        """
    from tkinter import simpledialog, messagebox
    import pathlib

    ruta_base = pathlib.Path(getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))))
    ruta_data = os.path.abspath(os.path.join(ruta_base, "..", "data"))
    os.makedirs(ruta_data, exist_ok=True)

    ruta_usuario = os.path.join(ruta_data, "user.txt")

    if not os.path.exists(ruta_usuario):
        nombre = simpledialog.askstring("Identificación", "Por favor ingresa tu nombre completo:")
        if not nombre:
            messagebox.showerror("Error", "Debes ingresar un nombre para continuar.")
            sys.exit()
        with open(ruta_usuario, "w", encoding="utf-8") as f: f.write(nombre.strip())
    else:
        with open(ruta_usuario, "r", encoding="utf-8") as f: nombre = f.read().strip()

    return nombre

if __name__ == "__main__":
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()

    nombre_usuario = obtener_nombre_usuario()

    root.deiconify()
    app = DraftSenderApp(root)
    root.mainloop()