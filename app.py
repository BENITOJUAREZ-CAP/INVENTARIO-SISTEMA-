from datetime import datetime
import pandas as pd
from google.oauth2.service_account import Credentials
import gspread
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Control de Inventario por Maletines",
    page_icon="🧰",
    layout="wide",
)

# IDs de Google Sheets
EXCEL_MAESTRO_ID = "1Chjc0zz3T0qF6TaydjQxLa7bI12sT2ZS11gYl6aeun0"
EXCEL_MALETINES_ID = "1StjBMVIkueBy9sVy5dNh1yIpG5he-51iCYeeaj_mixM"

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ENCABEZADOS = [
    "Folio de producto",
    "Clave de Producto",
    "Producto",
    "Tipo",
    "Uso",
    "Planillas",
    "Fecha",
    "No. de Serie",
    "Caja",
    "Estatus",
    "Folio del Maletín",
    "Número de Maletín",
    "Entidad de Envío",
    "Registrado Por",
    "Cantidad",
    "Servidor de la Salud",
    "Observaciones",
]


@st.cache_resource
def get_gspread_client():
  try:
    if "gcp_service_
