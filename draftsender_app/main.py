import os
import re
import sys
import atexit
import subprocess
from tkinter import Tk, Toplevel, Label, Entry, Button, messagebox
from draftsender_app.logger_utils import configurar_logger
from draftsender_app.gui import DraftSenderApp
from draftsender_app.ui_utils import get_data_path
from draftsender_app.version import obtener_version_local
from draftsender_app.elimina_v_anterior import eliminar_ejecutable_anterior_si_es_necesario

def abrir_outlook_clasico():
    """
        Intenta abrir el acceso directo de Outlook clásico desde el menú de inicio.
        Imprime un mensaje indicando si se encontró o no el acceso directo.
        """
    ruta_acceso_directo = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook (classic).lnk"
    if os.path.exists(ruta_acceso_directo):
        subprocess.Popen(['cmd', '/c', 'start', '', ruta_acceso_directo], shell=True)
        print("Se abrió Outlook clásico desde acceso directo.")
    else:
        print("No se encontró el acceso directo de Outlook clásico.")

def solicitar_nombre_usuario(root: Tk) -> str:
    nombre_usuario = None

    def guardar_nombre(event=None):
        nonlocal nombre_usuario
        valor = entry_nombre.get().strip()
        if not valor:
            messagebox.showwarning("Campo vacío", "Por favor, ingresa tu correo institucional.")
            return
        if not re.match(r"^[\w\.-]+@mejoreferido\.cl$", valor):
            messagebox.showerror("Correo inválido", "Ingresar tu correo institucional.")
            return
        nombre_usuario = valor
        ventana.destroy()

    def cerrar_app():
        nonlocal nombre_usuario
        nombre_usuario = None
        if ventana and ventana.winfo_exists():
            ventana.destroy()
        if root and root.winfo_exists():
            root.quit()

    ventana = Toplevel(root)
    ventana.title("¡Bienvenido a DraftSender!")
    ventana.geometry("400x180")
    ventana.resizable(False, False)
    ventana.grab_set()

    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    icon_path = os.path.join(base_path, "config", "icon.ico")
    if os.path.exists(icon_path):
        try:
            ventana.iconbitmap(icon_path)
        except Exception:
            pass

    Label(ventana, text="¡Bienvenido a DraftSender!", font=("Arial", 11)).pack(pady=(20, 5))
    Label(ventana, text="Escribe tu correo para continuar:", font=("Arial", 10)).pack(pady=(5, 5))

    entry_nombre = Entry(ventana, font=("Arial", 12), width=30)
    entry_nombre.pack(pady=5)
    entry_nombre.focus()

    Button(ventana, text="Entrar", command=guardar_nombre, font=("Arial", 10), width=15).pack(pady=10)

    ventana.bind("<Return>", guardar_nombre)
    ventana.protocol("WM_DELETE_WINDOW", cerrar_app)
    root.wait_window(ventana)

    if not nombre_usuario:
        return None

    # Solo si se ingresó correctamente, guardar en archivo
    data_dir = get_data_path()
    os.makedirs(data_dir, exist_ok=True)
    user_file = os.path.join(data_dir, "user.txt")
    with open(user_file, "w", encoding="utf-8") as f:
        f.write(nombre_usuario)

    return nombre_usuario.strip()

def cargar_icono_ventana(root: Tk) -> None:
    """
        Intenta cargar un ícono personalizado para la ventana principal de la app.

        Args:
            root (Tk): Ventana principal de la aplicación.
        """
    try:
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "config", "icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass

logger = None  # declaración global segura para usar en except

def correo_valido(correo: str) -> bool:
    return bool(re.match(r"^[\w\.-]+@mejoreferido\.cl$", correo))

import os
import sys
import time
import glob

def eliminar_version_anterior_si_es_necesario():
    """
    Si este ejecutable es 'DraftSender vX.Y.Z.exe', elimina cualquier otra 'DraftSender v*.exe'
    excepto a sí mismo. No toca la carpeta 'data'.
    """
    exe_actual = os.path.abspath(sys.argv[0])
    nombre_actual = os.path.basename(exe_actual)
    carpeta = os.path.dirname(exe_actual)

    if not nombre_actual.lower().startswith("draftsender v") or not nombre_actual.lower().endswith(".exe"):
        return  # No está en formato de versión numerada

    for exe in glob.glob(os.path.join(carpeta, "DraftSender v*.exe")):
        exe_path = os.path.abspath(exe)
        if exe_path != exe_actual:
            try:
                os.remove(exe_path)
                print(f"Versión anterior eliminada: {exe}")
            except Exception as e:
                print(f"Error al eliminar {exe}: {e}")

eliminar_version_anterior_si_es_necesario()

def main() -> None:
    """
    Función principal que inicializa la aplicación DraftSender con manejo robusto de errores.
    """
    eliminar_ejecutable_anterior_si_es_necesario()

    global logger

    try:
        root = Tk()
        root.withdraw()

        # Paso 1: Leer correo si ya existe y es válido
        nombre_usuario = None
        user_file = os.path.join(get_data_path(), "user.txt")
        if os.path.exists(user_file):
            with open(user_file, "r", encoding="utf-8") as f:
                posible_correo = f.read().strip()
                if correo_valido(posible_correo):
                    nombre_usuario = posible_correo

        # Paso 2: Si no hay correo válido, solicitarlo
        if not nombre_usuario:
            try:
                nombre_usuario = solicitar_nombre_usuario(root)
            except Exception as e:
                root.destroy()
                return

        if not nombre_usuario or not root.winfo_exists():
            root.destroy()
            return

        # Paso 3: Configurar logger
        try:
            logger = configurar_logger(usuario=nombre_usuario)
            logger.info("Iniciando aplicación")
        except Exception as e:
            print(f"Fallo crítico al configurar el logger: {e}")
            messagebox.showerror("Logger", "No se pudo inicializar el sistema de registro.")
            return

        # Paso 4: Configurar GUI
        version = obtener_version_local()
        titulo_app = f"DraftSender {version} - {nombre_usuario}"

        root.title(titulo_app)

        try:
            cargar_icono_ventana(root)
            root.deiconify()
        except Exception as e:
            logger.warning(f"No se pudo cargar el ícono de la aplicación: {e}")

        # Paso 5: Iniciar interfaz principal
        try:
            DraftSenderApp(root, titulo_ventana=titulo_app)
            atexit.register(lambda: logger.info("Aplicación finalizada."))
            root.mainloop()
        except Exception as e:
            logger.exception("Fallo al cargar la interfaz principal")
            messagebox.showerror("Error crítico", f"No se pudo iniciar la interfaz principal:\n{e}")
            if root and root.winfo_exists():
                try:
                    root.destroy()
                except Exception:
                    pass

    except Exception as e:
        if logger:
            logger.exception("Fallo inesperado en el arranque general")
        else:
            print(f"Fallo crítico antes de iniciar el logger: {e}")
        messagebox.showerror("Error grave", f"Ocurrió un error inesperado:\n{e}")

    finally:
        try:
            if root and root.winfo_exists():
                root.destroy()
        except Exception:
            pass

if __name__ == "__main__":
    main()