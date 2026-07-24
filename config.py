#!/usr/bin/env python3
"""
Configuración compartida entre los scripts.
Centraliza la ruta del archivo Excel para no repetirla en cada script.

Para cambiar el archivo, editá EXCEL_FILE aquí (o definí EXCEL_FILE en el .env).
"""

import os
from dotenv import load_dotenv

load_dotenv()

# Ruta del archivo Excel a procesar.
# Prioridad: variable de entorno EXCEL_FILE > valor por defecto.
EXCEL_FILE = os.getenv("EXCEL_FILE", "docs/New Anode Grades 4 Helder.xlsx")
