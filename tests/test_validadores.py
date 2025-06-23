import pytest
from borradores import es_email_valido

@pytest.mark.parametrize("email, esperado", [
    ("usuario@dominio.com", True),
    ("nombre.apellido@empresa.cl", True),
    ("malformato@", False),
    ("@incompleto.com", False),
    ("usuario@.com", False),
    ("usuario@dominio", False),
    ("usuario@dominio.c", True),  # dominio corto válido
])
def test_es_email_valido(email, esperado):
    assert es_email_valido(email) == esperado