import tkinter as tk
import threading
import sys
import os
import time
import win32com.client
import logging
from tkinter import ttk
from draftsender_app.ui_utils import mostrar_error
from draftsender_app.outlook_utils import obtener_cuentas_activas
from draftsender_app.envios import contar_borradores
from draftsender_app.envios import enviar_siguiente_borrador
from draftsender_app.outlook_utils import obtener_carpeta_borradores
from draftsender_app.outlook_utils import obtener_cuenta_outlook

logger = logging.getLogger("DraftSender")

class EnvioCuenta:
    def __init__(self, root, cuenta_id, intervalos_usados, cuentas_usadas):
        """
        Inicializa la interfaz gráfica de programación de envíos, junto con sus variables y componentes de estado.

        Args:
            root (tk.Toplevel): Ventana padre de Tkinter.
        """
        self.root = root
        self.cuenta_id = cuenta_id

        # Conjuntos compartidos para validación global
        self.intervalos_usados = intervalos_usados
        self.cuentas_usadas = cuentas_usadas
        self.cuentas_envio_ref = []  # Referencia a los demás bloques

        # Outlook
        self.outlook = win32com.client.Dispatch("Outlook.Application")

        # Estado interno
        self.cuenta = ""
        self.intervalo = 30
        self.total_borradores = 0
        self.enviados = 0
        self.enviar_event = threading.Event()
        self.thread = None
        self.countdown_id = None
        self.reanudando = False

        # Rastreo de selección anterior (para liberar recursos correctamente)
        self._cuenta_anterior = None
        self._intervalo_anterior = None

        # Componentes UI
        self.combo_cuenta = None
        self.combo_intervalo = None
        self.boton_iniciar = None
        self.boton_finalizar = None
        self.etiqueta_estado = None
        self.etiqueta_countdown = None

        # Opciones de intervalo disponibles
        self.intervalos_disponibles = ["Selecciona intervalo...", "30", "60", "90"]

    def construir_ui(self, parent, obtener_cuentas_disponibles):
        """
        Construye todos los elementos gráficos de la ventana: combos, botones, labels y controles de envío.
        """
        self.obtener_cuentas_disponibles = obtener_cuentas_disponibles

        frame = tk.LabelFrame(parent, text=f"Cuenta #{self.cuenta_id + 1}", padx=10, pady=5)
        frame.pack(fill="x", padx=10, pady=5)

        fila = tk.Frame(frame)
        fila.pack(fill="x")

        self.combo_cuenta = ttk.Combobox(fila, state="readonly", width=30)
        self.combo_cuenta.grid(row=0, column=0, padx=(0, 5))
        self.combo_cuenta.bind("<<ComboboxSelected>>", self.actualizar_borradores_disponibles)

        self.intervalos_disponibles = ["Selecciona intervalo...", "30", "60", "90"]
        self.combo_intervalo = ttk.Combobox(fila, state="disabled", width=18, values=self.intervalos_disponibles)
        self.combo_intervalo.grid(row=0, column=1, padx=(0,5))
        self.combo_intervalo.current(0)
        self.combo_intervalo.bind("<<ComboboxSelected>>", self.seleccionar_intervalo_unico)

        botones = tk.Frame(fila)
        botones.grid(row=0, column=2, sticky="e")

        self.boton_iniciar = tk.Button(botones, text="Iniciar", command=self.iniciar_envio, bg="lightgreen", width=5, state="disabled")
        self.boton_iniciar.pack(side="left", padx=1)

        self.boton_finalizar = tk.Button(botones, text="Detener", command=self.finalizar_envio, width=7, state="disabled")
        self.boton_finalizar.pack(side="left", padx=1)

        fila_progreso = tk.Frame(frame)
        fila_progreso.pack(fill="x", pady=(5, 0))

        self.etiqueta_estado = tk.Label(fila_progreso, text="", font=("Arial", 12))
        self.etiqueta_estado.pack(side="left", padx=(5, 0))

        self.etiqueta_countdown = tk.Label(fila_progreso, text="", font=("Arial", 10), fg="green")
        self.etiqueta_countdown.pack(side="right", padx=(0, 5))

        self.actualizar_opciones_cuenta()
        self.validar_intervalo()

    def resetear_combo_intervalo(self):
        """
        Resetea el combobox de intervalos a su estado inicial por defecto.
        """
        if self.combo_intervalo:
            self.combo_intervalo.set("Selecciona intervalo...")
            self.combo_intervalo.config(state="disabled")

    def resetear_combo_cuenta(self):
        """
        Desactiva el combobox de cuentas para que no se pueda editar hasta que haya opciones disponibles.
        """
        if self.combo_cuenta:
            self.combo_cuenta.set("Selecciona cuenta...")
            self.combo_cuenta.config(state="readonly")

    def iniciar_countdown(self):
        """
        Inicia la cuenta regresiva para el envío, comenzando desde el intervalo seleccionado.
        """
        try:
            intervalo = int(self.combo_intervalo.get())
        except ValueError:
            return
        self.tiempo_restante = intervalo
        self.actualizar_countdown()

    def actualizar_countdown(self):
        """
        Actualiza el label de cuenta regresiva cada segundo. Ejecuta el envío cuando el contador llega a cero.
        """
        if not self.enviar_event.is_set():
            return

        try:
            if not self.root.winfo_exists():
                return
        except:
            return

        try:
            if self.tiempo_restante > 0:
                tiempo_str = self.formatear_tiempo(self.tiempo_restante)

                if self.etiqueta_countdown and self.etiqueta_countdown.winfo_exists():
                    self.etiqueta_countdown.config(text=f"⏳ Tiempo restante: {tiempo_str}")

                self.tiempo_restante -= 1
                self.countdown_id = self.root.after(1000, self.actualizar_countdown)
            else:
                if self.etiqueta_countdown and self.etiqueta_countdown.winfo_exists():
                    self.etiqueta_countdown.config(text="Esperando próximo envío...")
                self.countdown_id = None
        except:
            pass

    def actualizar_opciones_cuenta(self, _event=None):
        """
        Actualiza el listado de cuentas disponibles en el combobox según las detectadas en Outlook.
        """
        todas = self.obtener_cuentas_disponibles(self.cuenta_id)
        usadas = self.cuentas_usadas.copy()

        # Si esta instancia tiene ya una cuenta seleccionada, permitir mantenerla
        actual = self.combo_cuenta.get()
        if actual in usadas:
            usadas.discard(actual)

        # Filtrar las cuentas disponibles
        disponibles = ["Selecciona cuenta..."] + [c for c in todas if c not in usadas]
        self.combo_cuenta["values"] = disponibles

        if actual in disponibles:
            self.combo_cuenta.set(actual)
        else:
            self.combo_cuenta.current(0)
            actual = "Selecciona cuenta..."

        # Si no hay cuenta válida seleccionada, desactiva y limpia
        if not actual or actual.startswith("Selecciona"):
            self.combo_intervalo.config(state="disabled")
            self.etiqueta_countdown.config(text="")
            if self.etiqueta_estado:
                self.etiqueta_estado.config(text="")
            self.validar_intervalo()
            return

        # Solo si es válida, actualizar borradores e intervalo
        self.actualizar_borradores_disponibles()

    def validar_intervalo(self, _event=None):
        """
        Verifica que el intervalo ingresado sea un número válido y positivo.

        Returns:
            bool: True si es válido, False si no.
        """
        cuenta = self.combo_cuenta.get().strip()
        intervalo = self.combo_intervalo.get().strip()

        cuenta_valida = cuenta and not cuenta.startswith("Selecciona")
        intervalo_valido = intervalo.isdigit()
        borradores_validos = self.total_borradores > 0

        if cuenta_valida and intervalo_valido and borradores_validos:
            self.boton_iniciar.config(state="normal", bg="lightgreen")
        else:
            self.boton_iniciar.config(state="disabled", bg=self.root.cget("bg"))

    def actualizar_borradores_disponibles(self, _event=None):
        """
        Muestra la cantidad de borradores disponibles para envío en la cuenta seleccionada.
        """
        cuenta_actual = self.combo_cuenta.get()

        # Liberar cuenta anterior si existía
        if hasattr(self, "_cuenta_anterior") and self._cuenta_anterior in self.cuentas_usadas:
            self.cuentas_usadas.discard(self._cuenta_anterior)

        # Verificar si la cuenta ya está en uso en otro bloque
        if cuenta_actual and not cuenta_actual.startswith("Selecciona"):
            if cuenta_actual in self.cuentas_usadas:
                bloque_ocupante = next(
                    (f"Cuenta #{ce.cuenta_id + 1}"
                     for ce in self.cuentas_envio_ref
                     if ce is not self and ce.combo_cuenta.get() == cuenta_actual),
                    "otro bloque"
                )
                self.root.after(100, lambda: mostrar_error(
                    "Cuenta en uso",
                    f"La cuenta '{cuenta_actual}' ya está siendo usada en {bloque_ocupante}.\nPor favor, selecciona otra."))
                self.root.after(150, lambda: self.combo_cuenta.set("Selecciona cuenta..."))
                self.etiqueta_countdown.config(text="")
                if cuenta_actual and not cuenta_actual.startswith("Selecciona"):
                    self.combo_intervalo.config(state="readonly")
                else:
                    self.combo_intervalo.config(state="disabled")

                self.validar_intervalo()

                return
            self.cuentas_usadas.add(cuenta_actual)
            self._cuenta_anterior = cuenta_actual
        else:
            self._cuenta_anterior = None

        if not cuenta_actual or cuenta_actual.startswith("Selecciona"):
            self.combo_intervalo.config(state="disabled")
            self.etiqueta_countdown.config(text="")
            if self.etiqueta_estado:
                self.etiqueta_estado.config(text="")
            self.validar_intervalo()
            return

        self.total_borradores = contar_borradores(cuenta_actual)
        self.enviados = 0
        self.combo_intervalo.config(state="readonly")
        self.validar_intervalo()

        if self.etiqueta_estado:
            self.etiqueta_estado.config(text=f"Enviados: {self.enviados} | Restantes: {self.total_borradores}")

    def loggear_intervalo_seleccionado(self, _event=None):
        """
        Registra en el log el intervalo seleccionado por el usuario para el envío de correos.
        """
        intervalo_actual = self.combo_intervalo.get()
        if intervalo_actual == "Selecciona intervalo...":
            self.etiqueta_countdown.config(text="")
            self.validar_intervalo()
            return

        if not hasattr(self, "_intervalo_anterior") or self._intervalo_anterior != intervalo_actual:
            self._intervalo_anterior = intervalo_actual

        try:
            intervalo = int(intervalo_actual)
            if self.total_borradores > 0:
                total = intervalo * self.total_borradores
                tiempo_estimado = self.formatear_tiempo(total)
                self.etiqueta_countdown.config(text=f"Tiempo Estimado: {tiempo_estimado}")
            else:
                self.etiqueta_countdown.config(text="")

        except ValueError:
            self.etiqueta_countdown.config(text="")

        self.validar_intervalo()

    def formatear_tiempo(self, total_segundos: int) -> str:
        """
        Convierte segundos en formato `mm:ss` para mostrar en la interfaz.

        Args:
            segundos (int): Tiempo total en segundos.

        Returns:
            str: Cadena formateada tipo "04:15".
        """
        horas = total_segundos // 3600
        minutos = (total_segundos % 3600) // 60
        segundos = total_segundos % 60
        return f"{horas:02}:{minutos:02}:{segundos:02}"

    def ejecutar_envio_despues(self, intervalo_ms):
        """
        Ejecuta el envío de correos después de la cuenta regresiva, usando la cuenta y el intervalo configurados.
        """
        import time
        time.sleep(intervalo_ms / 1000)
        self.root.after(0, self.programar_envios)

    def iniciar_envio(self):
        """
        Prepara y ejecuta el primer ciclo de envío de borradores según configuración del usuario.
        """
        from draftsender_app.logger_utils import configurar_logger
        logger = configurar_logger()

        self.enviar_event.set()

        self.boton_iniciar.config(state="disabled", bg=self.root.cget("bg"))
        self.boton_finalizar.config(state="normal", bg="red")
        self.combo_intervalo.config(state="disabled")
        self.combo_cuenta.config(state="disabled")

        if not self.outlook:
            self.outlook = win32com.client.Dispatch("Outlook.Application")

        try:
            cuenta = self.combo_cuenta.get().strip()
            intervalo_str = self.combo_intervalo.get().strip()

            if not cuenta or cuenta.startswith("Selecciona") or not intervalo_str.isdigit():
                raise ValueError("Debe seleccionar una cuenta e intervalo válidos antes de iniciar.")

            intervalo = int(intervalo_str)

            if self.enviados == 0:
                self.total_borradores = contar_borradores(cuenta)
                self.tiempo_total = intervalo * self.total_borradores
                self.tiempo_restante = self.tiempo_total

                if self.total_borradores > 0:
                    estimado_min = round(self.tiempo_total / 60, 2)
                    logger.info(
                        f"Inicio de envío automatizado | Cuenta: {cuenta} | Bloque: Cuenta #{self.cuenta_id + 1}")
                    logger.info(
                        f"Intervalo: {intervalo} segundos | Borradores: {self.total_borradores} | Tiempo estimado: {estimado_min} minutos")
                else:
                    logger.warning(f"No hay borradores para enviar desde la cuenta {cuenta}")

            self.actualizar_countdown()

            self.envio_id = self.root.after(
                intervalo * 1000,
                lambda: threading.Thread(
                    target=self.programar_envios,
                    daemon=True
                ).start()
            )

            self.reanudando = False

        except Exception as e:
            if self.etiqueta_estado:
                self.etiqueta_estado.config(text=f"Error al iniciar: {e}")
            logger.exception(f"Error al iniciar el envío: {e}")

    def finalizar_envio(self):
        """
        Detiene el envío y restablece la interfaz a su estado inicial.
        """
        from draftsender_app.logger_utils import configurar_logger
        logger = configurar_logger()

        self.enviar_event.clear()
        self.reanudando = False

        cuenta = self.combo_cuenta.get().strip()
        logger.info(f"Envío detenido por el usuario | Bloque: Cuenta #{self.cuenta_id + 1} | Cuenta: {cuenta}")

        if self.etiqueta_estado:
            self.etiqueta_estado.config(text="")

        try:
            if hasattr(self, "countdown_id") and self.countdown_id:
                self.root.after_cancel(self.countdown_id)
                self.countdown_id = None
        except:
            pass

        try:
            if hasattr(self, "envio_id") and self.envio_id:
                self.root.after_cancel(self.envio_id)
                self.envio_id = None
        except:
            pass

        self.enviados = 0
        self.tiempo_restante = 0

        if self.etiqueta_countdown:
            self.etiqueta_countdown.config(text="00:00:00")

        self.boton_finalizar.config(state="disabled", bg=self.root.cget("bg"))
        self.boton_iniciar.config(state="disabled", bg=self.root.cget("bg"))  # desactivar

        # liberar intervalo y cuenta
        if hasattr(self, "_intervalo_anterior") and self._intervalo_anterior in self.intervalos_usados:
            self.intervalos_usados.discard(self._intervalo_anterior)
            self._intervalo_anterior = None

        if hasattr(self, "_cuenta_anterior") and self._cuenta_anterior in self.cuentas_usadas:
            self.cuentas_usadas.discard(self._cuenta_anterior)
            self._cuenta_anterior = None

        self.etiqueta_countdown.config(text="")
        self.etiqueta_estado.config(text="")

        self.resetear_combo_intervalo()
        self.resetear_combo_cuenta()

        self.validar_intervalo()

    def programar_envios(self):
        """
        Permite configurar múltiples cuentas con sus propios intervalos de envío y botones independientes.
        """
        import pythoncom
        pythoncom.CoInitialize()

        if not self.enviar_event.is_set():
            return

        try:
            if not self.root.winfo_exists():
                return
        except:
            return

        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            cuenta = self.combo_cuenta.get()
            cuenta_outlook = obtener_cuenta_outlook(None, cuenta)
            carpeta = obtener_carpeta_borradores(namespace, cuenta)

            if cuenta_outlook and carpeta:
                borradores_disponibles = carpeta.Items.Count
                contador = {"enviados": 0, "restantes": borradores_disponibles}

                intervalo_segundos = int(self.combo_intervalo.get())

                while self.enviar_event.is_set() and contador["restantes"] > 0:
                    enviado = enviar_siguiente_borrador(cuenta, outlook, intervalo_segundos, self.etiqueta_estado, contador)

                    if self.etiqueta_estado:
                        self.etiqueta_estado.config(text=f"Enviados: {contador['enviados']} | Restantes: {contador['restantes']}")

                    self.root.update_idletasks()

                    if not enviado:
                        break

                    time.sleep(intervalo_segundos)

                self.finalizar_envio()
            else:
                self.etiqueta_estado.config(text="Cuenta o carpeta inválida")

        except Exception as e:
            import traceback
            if self.etiqueta_estado:
                self.etiqueta_estado.config(text=f"Error: {str(e)}")
            print(traceback.format_exc())

    def actualizar_intervalos_disponibles(self):
        """
        Actualiza dinámicamente los intervalos posibles para cada cuenta, evitando duplicados o conflictos.
        """
        actual = self.combo_intervalo.get()
        if not isinstance(self.intervalos_usados, set):
            return

        opciones_disponibles = [i for i in self.intervalos_disponibles
                                if i not in self.intervalos_usados or i == actual]
        self.combo_intervalo["values"] = ["Selecciona intervalo..."] + opciones_disponibles

    def seleccionar_intervalo_unico(self, event):
        """
        Marca el intervalo como seleccionado y bloquea cambios si ya está en uso por otra cuenta.
        """
        seleccionado = self.combo_intervalo.get()

        if not isinstance(self.intervalos_usados, set):
            return

        if hasattr(self, "_intervalo_anterior") and self._intervalo_anterior in self.intervalos_usados:
            self.intervalos_usados.discard(self._intervalo_anterior)

        if seleccionado in self.intervalos_usados:
            # Buscar qué bloque lo está usando
            bloque_ocupante = next(
                (f"Cuenta #{ce.cuenta_id + 1}"
                 for ce in getattr(self, "cuentas_envio_ref", [])
                 if ce is not self and ce.combo_intervalo.get() == seleccionado),
                "otro bloque")

            self.root.after(100, lambda: mostrar_error(
                "Intervalo duplicado",
                f"Este intervalo ya está siendo usado en {bloque_ocupante}. Por favor, elige otro."))
            self.root.after(150, lambda: self.combo_intervalo.set("Selecciona intervalo..."))
            self.etiqueta_countdown.config(text="")
            self.validar_intervalo(event)
            return

        self.intervalos_usados.add(seleccionado)
        self._intervalo_anterior = seleccionado
        self.validar_intervalo(event)

        try:
            intervalo = int(seleccionado)
            if self.total_borradores > 0:
                tiempo_estimado = intervalo * self.total_borradores
                tiempo_str = self.formatear_tiempo(tiempo_estimado)
                self.etiqueta_countdown.config(text=f"Tiempo Estimado: {tiempo_str}")
        except ValueError:
            self.etiqueta_countdown.config(text="")

def lanzar_envio_desde_gui(_=None):
    """
    Abre la ventana gráfica que permite seleccionar intervalos y cuentas para enviar los borradores programados.
    """
    from draftsender_app.logger_utils import configurar_logger
    logger = configurar_logger()

    cuentas_disponibles = obtener_cuentas_activas()
    if not cuentas_disponibles:
        mostrar_error("Sin cuentas", "No se encontraron cuentas activas en Outlook.")
        return

    root = tk.Tk()
    root.title("DraftSender - Envíos Automatizados")
    root.geometry("500x350")

    try:
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        icon_path = os.path.join(base_path, "config", "icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception as e:
        logger.warning(f"No se pudo cargar ícono: {e}")

    tk.Label(root, text="Puedes seleccionar hasta 3 cuentas para envío automatizado por intervalos:",).pack(pady=10)

    container = tk.Frame(root)
    container.pack(fill="both", expand=True)

    intervalos_usados = set()
    cuentas_usadas = set()
    cuentas_envio = []

    def obtener_opciones_disponibles(_):
        """
        Devuelve los intervalos de envío que aún no han sido asignados a ninguna cuenta.

        Returns:
            list[int]: Lista de intervalos disponibles.
        """
        return cuentas_disponibles

    for i in range(3):
        cuenta_envio = EnvioCuenta(
            root,
            cuenta_id=i,
            intervalos_usados=intervalos_usados,
            cuentas_usadas=cuentas_usadas)
        cuenta_envio.construir_ui(container, obtener_opciones_disponibles)
        cuentas_envio.append(cuenta_envio)

    # Asignar referencia cruzada a todos los bloques
    for ce in cuentas_envio:
        ce.cuentas_envio_ref = cuentas_envio

    logger.info("Envíos automatizados: Se ejecutó correctamente.")

    def al_cerrar_ventana():
        """
        Limpia hilos activos y cierra correctamente la ventana de programación de envíos.
        """
        logger.info("Envíos automatizados: Se cerró por el usuario.")

        for cuenta_envio in cuentas_envio:
            try:
                if hasattr(cuenta_envio, "countdown_id") and cuenta_envio.countdown_id:
                    cuenta_envio.root.after_cancel(cuenta_envio.countdown_id)
            except:
                pass
            try:
                if hasattr(cuenta_envio, "envio_id") and cuenta_envio.envio_id:
                    cuenta_envio.root.after_cancel(cuenta_envio.envio_id)
            except:
                pass

            # Liberar intervalos y cuentas
            if hasattr(cuenta_envio, "_intervalo_anterior") and cuenta_envio._intervalo_anterior in intervalos_usados:
                intervalos_usados.discard(cuenta_envio._intervalo_anterior)
            if hasattr(cuenta_envio, "_cuenta_anterior") and cuenta_envio._cuenta_anterior in cuentas_usadas:
                cuentas_usadas.discard(cuenta_envio._cuenta_anterior)

            # Resetear combos visualmente
            try:
                cuenta_envio.resetear_combo_intervalo()
                cuenta_envio.resetear_combo_cuenta()

            except:
                pass

        root.destroy()

    root.protocol("WM_DELETE_WINDOW", al_cerrar_ventana)
    root.mainloop()