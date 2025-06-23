import pytest
import os
import pandas as pd

def test_excel_faltante_columna():
    ruta_excel = os.path.join("tests", "archivos", "formato_test.xlsx")
    df = pd.read_excel(ruta_excel)