import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

# Configuración de página
st.set_page_config(
    page_title="Inventario General y Carpetas Personales",
    page_icon="🧰",
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
def get_services():
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
            
            gc = gspread.authorize(creds)
            drive_service = build('drive', 'v3', credentials=creds)
            return gc, drive_service
        else:
            st.error("No se encontraron las credenciales en los Secrets.")
            return None, None
    except Exception as e:
        st.error(f"Error al conectar con Google APIs: {e}")
        return None, None

gc, drive_service = get_services()

def obtener_o_crear_carpeta(nombre_servidor, parent_folder_id):
    """Busca si existe la carpeta del servidor; si no, la crea."""
    query = f"'{parent_folder_id}' in parents and name = '{nombre_servidor}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name, webViewLink)").execute()
    items = results.get('files', [])
    
    if items:
        return items[0]['id'], items[0]['webViewLink']
    else:
        folder_metadata = {
            'name': nombre_servidor,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = drive_service.files().create(body=folder_metadata, fields='id, webViewLink').execute()
        return folder.get('id'), folder.get('webViewLink')

st.title("🧰 Control de Inventario por Maletines")

if gc and drive_service:
    try:
        sh_maestro = gc.open_by_key(EXCEL_MAESTRO_ID)
        worksheet = sh_maestro.get_worksheet(0)
        data = worksheet.get_all_records()
        
        if data:
            df = pd.DataFrame(data)
            
            # --- SECCIÓN 1: SELECCIONAR SERVIDOR ---
            st.subheader("🔍 Seleccionar Servidor de la Salud")
            columna_nombre = "SERVIDOR DE LA SALUD"
            
            if columna_nombre in df.columns:
                lista_servidores = df[columna_nombre].dropna().unique().tolist()
                servidor_seleccionado = st.selectbox(
                    "Selecciona o busca un Servidor de la Salud:", 
                    options=lista_servidores
                )
                
                # Mostrar registros actuales del servidor seleccionado
                df_filtrado = df[df[columna_nombre] == servidor_seleccionado]
                st.write(f"### Inventario actual de: **{servidor_seleccionado}**")
                st.dataframe(df_filtrado, use_container_width=True)
                
                # Gestión de Carpeta en Google Drive
                folder_id, folder_link = obtener_o_crear_carpeta(servidor_seleccionado, CARPETA_PERSONAL_ID)
                st.markdown(f"📂 **Carpeta en Google Drive:** [Abrir carpeta personal de {servidor_seleccionado}]({folder_link})")
                
                st.divider()
                
                # --- SECCIÓN 2: FORMULARIO DE CAPTURA POR MALETÍN ---
                st.subheader("📝 Capturar / Asignar Producto a Maletín")
                
                with st.form("form_maletines", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Selección de Maletín (Máximo 4)
                        maletin = st.selectbox(
                            "Selecciona el Maletín:",
                            options=["Maletín 1", "Maletín 2", "Maletín 3", "Maletín 4"]
                        )
                        
                        # Nombre / Descripción del Producto
                        producto = st.text_input("Nombre / Descripción del Producto:")
                        
                    with col2:
                        # Cantidad de este producto
                        cantidad = st.number_input("Cantidad (¿Cuántos son?):", min_value=1, value=1, step=1)
                        
                        # Folio o Número de Serie del producto
                        folio = st.text_input("Folio / Número de Serie:")
                    
                    observaciones = st.text_area("Observaciones del estado del producto/maletín:")
                    
                    # Cargar archivo de resguardo si aplica
                    archivo_subido = st.file_uploader(
                        "Adjuntar archivo/resguardo en PDF o imagen (Opcional):", 
                        type=["pdf", "png", "jpg", "jpeg"]
                    )
                    
                    btn_guardar = st.form_submit_button("💾 Guardar Registro de Producto")
                    
                    if btn_guardar:
                        if producto and folio:
                            enlace_archivo = "Sin archivo"
                            
                            # Subir archivo si fue adjuntado
                            if archivo_subido is not None:
                                file_metadata = {
                                    'name': f"{folio}_{archivo_subido.name}",
                                    'parents': [folder_id]
                                }
                                media = MediaIoBaseUpload(
                                    io.BytesIO(archivo_subido.read()), 
                                    mimetype=archivo_subido.type, 
                                    resumable=True
                                )
                                archivo_drive = drive_service.files().create(
                                    body=file_metadata, 
                                    media_body=media, 
                                    fields='webViewLink'
                                ).execute()
                                enlace_archivo = archivo_drive.get('webViewLink')
                            
                            # Estructura de fila a insertar en Excel Maestro
                            nueva_fila = [
                                servidor_seleccionado, 
                                maletin, 
                                producto, 
                                cantidad, 
                                folio, 
                                observaciones, 
                                enlace_archivo
                            ]
                            
                            worksheet.append_row(nueva_fila)
                            st.success(f"¡Se registró correctamente el producto '{producto}' en {maletin}!")
                            st.cache_data.clear()
                        else:
                            st.warning("Por favor ingresa al menos el Nombre del Producto y el Folio/Serie.")
                            
            else:
                st.warning(f"No se encontró la columna '{columna_nombre}' en la hoja.")
                st.dataframe(df, use_container_width=True)
                
    except Exception as e:
        st.error("Error al procesar la solicitud:")
        st.exception(e)
