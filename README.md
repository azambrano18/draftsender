
# Proyecto Creador de Borradores

Este proyecto permite la automatización del proceso de creación y envío de borradores de correos electrónicos utilizando Microsoft Outlook. Los usuarios pueden cargar archivos Excel con la lista de destinatarios y un archivo Word como plantilla para generar los borradores de manera eficiente.

## Características

- **Carga de archivos**: Permite cargar archivos Excel (.xlsx) para obtener destinatarios y asunto, y archivos Word (.docx) como plantilla para el contenido del correo.
- **Creación de borradores**: Genera borradores de correos electrónicos en Outlook con el contenido proporcionado.
- **Envío automatizado**: Envía los borradores de manera programada según el intervalo seleccionado.
- **Actualización automática**: Verifica si hay nuevas versiones del software y permite actualizarlo automáticamente desde GitHub.
- **Interfaz gráfica**: Utiliza Tkinter para una interfaz sencilla y visualmente amigable.

## Requisitos

- **Python 3.x**
- **Bibliotecas de Python**:
  - `win32com.client` (para interactuar con Outlook)
  - `pythoncom` (para inicializar el modelo de objetos COM de Outlook)
  - `pandas` (para manejar archivos Excel)
  - `mammoth` (para convertir archivos DOCX a HTML)
  - `tkinter` (para la interfaz gráfica)
  - `urllib` (para manejar actualizaciones desde GitHub)
  
- **Microsoft Outlook** debe estar instalado en el sistema para utilizar la funcionalidad de correos electrónicos.

## Instalación

1. **Clonar el repositorio**:
   ```bash
   git clone https://github.com//azambrano18/draftsender.git
   cd draftsender
   ```

2. **Instalar dependencias**:
   Asegúrate de tener todas las bibliotecas necesarias instaladas. Puedes hacerlo ejecutando:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configurar el entorno**:
   No se necesita ninguna configuración adicional para el funcionamiento básico, pero asegúrate de que Outlook esté instalado y configurado correctamente en tu sistema.

## Uso

1. **Cargar archivos**:
   - Usa el botón "Cargar Excel" para seleccionar el archivo `.xlsx` que contiene la lista de destinatarios, asunto y nombre.
   - Usa el botón "Cargar Texto Mail" para seleccionar el archivo `.docx` que contiene la plantilla para el cuerpo del correo.

2. **Generar borradores**:
   - Selecciona un perfil de Outlook.
   - Selecciona la cuenta asociada.
   - Haz clic en el botón "Crear Borradores" para generar los borradores en Outlook.

3. **Enviar borradores**:
   - Elige el intervalo de envío de los borradores en segundos.
   - Haz clic en el botón "Iniciar Envío" para comenzar a enviar los borradores de forma automatizada.

4. **Verificación de actualizaciones**:
   - El programa verifica automáticamente si hay nuevas versiones disponibles y te pide confirmación para descargarla e instalarla.

## Estructura del Proyecto

```
draftsender/
├── actualizacion.py          # Lógica para verificar y descargar actualizaciones del software.
├── archivos.py               # Funciones para cargar los archivos Excel y DOCX.
├── borradores.py             # Funciones para crear los borradores de correos en Outlook.
├── envios.py                 # Funciones para enviar los borradores automáticamente.
├── ejecutores.py             # Lógica para ejecutar scripts y validar datos.
├── estado.py                 # Variables de estado global para la aplicación.
├── logger_utils.py           # Configuración del logger para registrar actividades y errores.
└── __init__.py               # Marca el directorio como un paquete Python.
```

## Contribuciones

1. Haz un fork de este repositorio.
2. Crea una rama para tu contribución (`git checkout -b feature-nueva`).
3. Haz los cambios necesarios y confirma tus cambios (`git commit -am 'Agrega nueva característica'`).
4. Sube tus cambios (`git push origin feature-nueva`).
5. Abre un Pull Request.

## Licencia
