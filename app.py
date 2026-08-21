import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Configuración de página
st.set_page_config(
    page_title="Inventario General y Carpetas Personales",
    page_icon="📦",
    layout="wide"
)

# IDs de Google Drive / Sheets extraídos de tus enlaces actuales
EXCEL_MAESTRO_ID = "1Chjc0zz3T0qF6TaydjQxLa7bI12sT2ZS11gYl6aeun0"
CARPETA_PERSONAL_ID = "1QVW-qYDtNGYX9CjFIjzD0isYTRFj2F-u"
CARPETA_PRINCIPAL_ID = "14nPwqk129lZn5ACi12GN4RoAAvIvQIJh"

# Autenticación con Google APIs
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_gspread_client():
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
            return gspread.authorize(creds)
        else:
            st.error("No se encontraron las credenciales en los Secrets de Streamlit.")
            return None
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return None

gc = get_gspread_client()

st.title("📦 Inventario General y Carpetas Personales")

if gc:
    try:
        # Abrir el Excel Maestro
        sh_maestro = gc.open_by_key(EXCEL_MAESTRO_ID)
        worksheet = sh_maestro.get_worksheet(0)
        
        # Obtener datos
        data = worksheet.get_all_records()
        if data:
            df = pd.DataFrame(data)
            st.success("¡Conexión establecida correctamente con Google Drive y Sheets!")
            st.dataframe(df, use_container_width=True)
        else:
            st.info("Conexión exitosa, pero la hoja del Excel Maestro está vacía actualmente.")
            
    except Exception as e:
        st.error(f"Error de conexión o permisos:")
        st.warning("Verifica que la Service Account `buscador-benito-juarez@buscador-de-base-general.iam.gserviceaccount.com` sea Editora tanto del archivo Excel Maestro como de las carpetas.")
        st.exception(e)
