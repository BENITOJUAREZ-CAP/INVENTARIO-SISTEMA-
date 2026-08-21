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

# IDs de Google Drive / Sheets
EXCEL_MAESTRO_ID = "1Chjc0zz3T0qF6TaydjQxLa7bI12sT2ZS11gYl6aeun0"
CARPETA_PERSONAL_ID = "1QVW-qYDtNGYX9CjFIjzD0isYTRFj2F-u"
CARPETA_PRINCIPAL_ID = "14nPwqk129lZn5ACi12GN4RoAAvIvQIJh"

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
            st.error("No se encontraron las credenciales en los Secrets.")
            return None
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return None

gc = get_gspread_client()

st.title("📦 Inventario General y Carpetas Personales")

if gc:
    try:
        sh_maestro = gc.open_by_key(EXCEL_MAESTRO_ID)
        worksheet = sh_maestro.get_worksheet(0)
        data = worksheet.get_all_records()
        
        if data:
            df = pd.DataFrame(data)
            
            # --- SECCIÓN 1: SELECCIONAR SERVIDOR Y CONSULTAR ---
            st.subheader("🔍 Consulta por Servidor de la Salud")
            
            columna_nombre = "SERVIDOR DE LA SALUD"
            if columna_nombre in df.columns:
                lista_servidores = df[columna_nombre].dropna().unique().tolist()
                
                # Desplegable para seleccionar a una persona
                servidor_seleccionado = st.selectbox(
                    "Selecciona o busca un Servidor de la Salud:", 
                    options=lista_servidores
                )
                
                # Filtrar datos del servidor seleccionado
                df_filtrado = df[df[columna_nombre] == servidor_seleccionado]
                st.write(f"### Datos de: **{servidor_seleccionado}**")
                st.dataframe(df_filtrado, use_container_width=True)
                
                st.divider()
                
                # --- SECCIÓN 2: REGISTRO DE NUEVOS DATOS / INVENTARIO ---
                st.subheader("📝 Registrar / Actualizar Elemento de Inventario")
                
                with st.form("form_inventario", clear_on_submit=True):
                    st.write(f"Agregar/actualizar registro para: **{servidor_seleccionado}**")
                    
                    # Campos de entrada interactivos
                    concepto = st.text_input("Concepto / Artículo:")
                    folio = st.text_input("Folio / Serie:")
                    observaciones = st.text_area("Observaciones:")
                    
                    btn_guardar = st.form_submit_button("💾 Guardar Registro")
                    
                    if btn_guardar:
                        if concepto and folio:
                            # Añadir nueva fila al Excel Maestro en Google Drive
                            nueva_fila = [servidor_seleccionado, concepto, folio, observaciones]
                            worksheet.append_row(nueva_fila)
                            st.success("¡Registro guardado correctamente en Google Sheets!")
                            st.cache_data.clear()
                        else:
                            st.warning("Por favor completa al menos el Concepto y el Folio.")
                            
            else:
                st.warning(f"No se encontró la columna '{columna_nombre}' en la hoja.")
                st.dataframe(df, use_container_width=True)
                
    except Exception as e:
        st.error("Error al procesar la solicitud:")
        st.exception(e)
