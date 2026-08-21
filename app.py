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
            columna_nombre = "SERVIDOR DE LA SALUD"
            
            if columna_nombre in df.columns:
                lista_servidores = df[columna_nombre].dropna().unique().tolist()
                
                # --- PESTAÑAS DE NAVEGACIÓN ---
                tab_consultar, tab_agregar_persona, tab_eliminar_persona = st.tabs([
                    "🔍 Consultar y Registrar Productos", 
                    "➕ Agregar Nuevo Servidor", 
                    "🗑️ Eliminar Servidor"
                ])
                
                # === PESTAÑA 1: CONSULTA Y REGISTRO DE MALETINES ===
                with tab_consultar:
                    servidor_seleccionado = st.selectbox(
                        "Selecciona un Servidor de la Salud:", 
                        options=lista_servidores
                    )
                    
                    df_filtrado = df[df[columna_nombre] == servidor_seleccionado]
                    st.write(f"### Inventario actual de: **{servidor_seleccionado}**")
                    st.dataframe(df_filtrado, use_container_width=True)
                    
                    folder_id, folder_link = obtener_o_crear_carpeta(servidor_seleccionado, CARPETA_PERSONAL_ID)
                    st.markdown(f"📂 **Carpeta en Google Drive:** [Abrir carpeta personal de {servidor_seleccionado}]({folder_link})")
                    
                    st.divider()
                    
                    st.subheader("📝 Capturar / Asignar Producto a Maletín")
                    with st.form("form_maletines", clear_on_submit=True):
                        col1, col2 = st.columns(2)
                        with col1:
                            maletin = st.selectbox("Selecciona el Maletín:", ["Maletín 1", "Maletín 2", "Maletín 3", "Maletín 4"])
                            producto = st.text_input("Nombre / Descripción del Producto:")
                        with col2:
                            cantidad = st.number_input("Cantidad (¿Cuántos son?):", min_value=1, value=1, step=1)
                            folio = st.text_input("Folio / Número de Serie:")
                        
                        observaciones = st.text_area("Observaciones:")
                        archivo_subido = st.file_uploader("Adjuntar archivo (PDF/Imagen):", type=["pdf", "png", "jpg", "jpeg"])
                        
                        btn_guardar = st.form_submit_button("💾 Guardar Registro de Producto")
                        
                        if btn_guardar:
                            if producto and folio:
                                enlace_archivo = "Sin archivo"
                                if archivo_subido is not None:
                                    file_metadata = {'name': f"{folio}_{archivo_subido.name}", 'parents': [folder_id]}
                                    media = MediaIoBaseUpload(io.BytesIO(archivo_subido.read()), mimetype=archivo_subido.type, resumable=True)
                                    archivo_drive = drive_service.files().create(body=file_metadata, media_body=media, fields='webViewLink').execute()
                                    enlace_archivo = archivo_drive.get('webViewLink')
                                
                                nueva_fila = [servidor_seleccionado, maletin, producto, cantidad, folio, observaciones, enlace_archivo]
                                worksheet.append_row(nueva_fila)
                                st.success(f"¡Producto '{producto}' guardado exitosamente!")
                                st.rerun()
                            else:
                                st.warning("Ingresa el Nombre del Producto y el Folio/Serie.")

                # === PESTAÑA 2: AGREGAR NUEVO SERVIDOR ===
                with tab_agregar_persona:
                    st.subheader("➕ Agregar Nuevo Servidor de la Salud")
                    with st.form("form_nuevo_servidor", clear_on_submit=True):
                        nuevo_nombre = st.text_input("Nombre completo del Servidor de la Salud:")
                        btn_agregar = st.form_submit_button("➕ Registrar Servidor")
                        
                        if btn_agregar:
                            if nuevo_nombre.strip():
                                if nuevo_nombre.strip() in lista_servidores:
                                    st.warning("Este Servidor ya existe en la lista.")
                                else:
                                    # Fila inicial de registro
                                    nueva_fila = [nuevo_nombre.strip(), "Sin Maletín", "Sin Producto", 0, "N/A", "Registro inicial", "Sin archivo"]
                                    worksheet.append_row(nueva_fila)
                                    
                                    # Crear su carpeta en Drive
                                    obtener_o_crear_carpeta(nuevo_nombre.strip(), CARPETA_PERSONAL_ID)
                                    
                                    st.success(f"¡Servidor **{nuevo_nombre.strip()}** agregado exitosamente!")
                                    st.rerun()
                            else:
                                st.warning("Escribe un nombre válido.")

                # === PESTAÑA 3: ELIMINAR SERVIDOR ===
                with tab_eliminar_persona:
                    st.subheader("🗑️ Eliminar Servidor de la Salud")
                    st.error("⚠️ Esta acción eliminará todas las filas vinculadas a este Servidor en el Excel Maestro.")
                    
                    servidor_a_eliminar = st.selectbox(
                        "Selecciona el Servidor que deseas eliminar:", 
                        options=lista_servidores,
                        key="select_eliminar"
                    )
                    
                    confirmar = st.checkbox(f"Confirmo que deseo borrar a **{servidor_a_eliminar}**")
                    btn_eliminar = st.button("🗑️ Eliminar Definativamente", type="primary")
                    
                    if btn_eliminar:
                        if confirmar:
                            # Buscar y borrar todas las filas que coincidan
                            cell_list = worksheet.findall(servidor_a_eliminar)
                            rows_to_delete = sorted(list(set([cell.row for cell in cell_list])), reverse=True)
                            
                            for row_idx in rows_to_delete:
                                worksheet.delete_rows(row_idx)
                                
                            st.success(f"Se ha eliminado a **{servidor_a_eliminar}** y sus registros.")
                            st.rerun()
                        else:
                            st.warning("Marca la casilla de confirmación para poder proceder.")

    except Exception as e:
        st.error("Error al procesar la solicitud:")
        st.exception(e)
