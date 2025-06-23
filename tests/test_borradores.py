import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from unittest.mock import patch
from borradores import generar_borradores

@patch("borradores.crear_borrador")
def test_generar_borradores_mock(mock_crear):
    ruta_excel = os.path.join("tests", "archivos", "formato_test.xlsx")
    ruta_docx = os.path.join("tests", "archivos", "cuerpo_mail_test.docx")

    cantidad = generar_borradores(
        cuenta="test@correo.com",
        perfil="perfil_mock",
        ruta_excel=ruta_excel,
        ruta_docx=ruta_docx,
        callback_progreso=None
    )

    assert cantidad == 1
    mock_crear.assert_called_once()

    args = mock_crear.call_args[0]

    # Cuenta y destinatario
    assert args[0] == "test@correo.com"
    assert args[1] and "@" in args[1]

    # Asunto no vacío
    assert args[2] and isinstance(args[2], str)

    cuerpo = args[3]

    # Validar formato HTML
    assert cuerpo.startswith('<div style="font-family: Calibri, sans-serif; font-size: 11pt;">')
    assert cuerpo.endswith('</div>')

    # No deben quedar etiquetas sin reemplazar
    for etiqueta in ["[Nombre]", "{{Nombre}}", "[Empresa]", "{{Empresa}}"]:
        assert etiqueta not in cuerpo