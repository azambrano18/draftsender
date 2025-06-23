class EstadoApp:
    """
    Clase que encapsula el estado de la aplicación de forma controlada.
    Uso recomendado: pasar una instancia donde se requiera, en lugar de usar variables globales.
    """
    def __init__(self):
        """
        Inicializa el objeto de estado para gestionar variables compartidas o temporales.

        Este constructor puede ser extendido en el futuro si se agregan más propiedades.

        Args:
            sin argumentos explícitos distintos de self.
        """

        self.cuenta_seleccionada = None
        self.ruta_excel = None
        self.ruta_docx = None

estado_app = EstadoApp()