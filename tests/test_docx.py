import os
from borradores import cargar_cuerpo_desde_docx

def test_reemplazo_variables_docx():
    ruta_docx = os.path.join("tests", "archivos", "cuerpo_mail_test.docx")
    variables = {"Nombre": "Fernando"}

    html = cargar_cuerpo_desde_docx(ruta_docx, variables)

    assert "Fernando" in html
    assert "[Nombre]" not in html
    assert "{{Nombre}}" not in html
    assert html.startswith('<div style="font-family: Calibri')