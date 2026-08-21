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
        
        # --- OBTENER LISTA DE SERVIDORES (De las pestañas personales o del listado maestro) ---
        todas_las_hojas = [ws.title for ws.title in sh_maestro.worksheets()]
        hojas_reservadas = ["PRODUCTOS", "PERSONAL DE SALUDA", "Hoja 1"]
        lista_servidores = [h for h in todas_las_hojas if h not in hojas_reservadas]
        
        # Si existe la pestaña "PERSONAL DE SALUDA", extraer de ahí también
        try:
            ws_personal = sh_maestro.worksheet("PERSONAL DE SALUDA")
            datos_personal = ws_personal.get_all_values()
            nombres_personal = [fila[0] for fila in datos_personal if fila and fila[0].strip()]
            lista_servidores = sorted(list(set(lista_servidores + nombres_personal)))
        except:
            pass

        # --- OBTENER CATÁLOGO DESDE LA PESTAÑA "PRODUCTOS" ---
        lista_productos = []
        try:
            ws_productos = sh_maestro.worksheet("PRODUCTOS")
            valores_productos = ws_productos.get_all_values()
            # Lee todos los elementos no vacíos de la columna A
            lista_productos = [fila[0].strip() for fila in valores_productos if fila and fila[0].strip()]
        except Exception as e:
            st.warning("No se pudo cargar la pestaña 'PRODUCTOS'. Verifica el nombre exacto de la pestaña en Google Sheets.")

        # --- PESTAÑAS DE LA APLICACIÓN ---
        tab_consultar, tab_agregar_persona, tab_eliminar_persona = st.tabs([
            "🔍 Consultar y Asignar Productos", 
            "➕ Agregar Nuevo Servidor", 
            "🗑️ Eliminar Servidor"
        ])
        
        # === PESTAÑA 1: CONSULTA Y ASIGNACIÓN ===
        with tab_consultar:
            if lista_servidores:
                servidor_seleccionado = st.selectbox(
                    "Selecciona un Servidor de la Salud:", 
                    options=lista_servidores
                )
                
                # Cargar o crear la hoja del servidor seleccionado
                try:
                    worksheet_servidor = sh_maestro.worksheet(servidor_seleccionado)
                except:
                    worksheet_servidor = sh_maestro.add_worksheet(title=servidor_seleccionado, rows="100", cols="10")
                    worksheet_servidor.append_row(["SERVIDOR DE LA SALUD", "MALETÍN", "PRODUCTO", "CANTIDAD", "FOLIO", "OBSERVACIONES", "COMPROBANTE"])

                data_servidor = worksheet_servidor.get_all_records()
                df_filtrado = pd.DataFrame(data_servidor)
                
                st.write(f"### Inventario actual de: **{servidor_seleccionado}**")
                st.dataframe(df_filtrado, use_container_width=True)
                
                folder_id, folder_link = obtener_o_crear_carpeta(servidor_seleccionado, CARPETA_PERSONAL_ID)
                st.markdown(f"📂 **Carpeta en Google Drive:** [Abrir carpeta personal de {servidor_seleccionado}]({folder_link})")
                
                st.divider()
                st.subheader("📋 Verificación y Asignación de Productos")
                
                col1, col2 = st.columns(2)
                with col1:
                    maletin_seleccionado = st.selectbox(
                        "Selecciona el Maletín a registrar/revisar:", 
                        ["Maletín 1", "Maletín 2", "Maletín 3", "Maletín 4"]
                    )
                with col2:
                    if lista_productos:
                        producto_seleccionado = st.selectbox("Selecciona un producto del catálogo:", options=lista_productos)
                    else:
                        producto_seleccionado = st.text_input("Nombre del producto:")
                
                tiene_producto = st.radio(
                    f"¿El **{maletin_seleccionado}** de **{servidor_seleccionado}** contiene el producto **'{producto_seleccionado}'**?",
                    ["Sí", "No"],
                    horizontal=True
                )
                
                if tiene_producto == "Sí":
                    with st.form("form_confirmar_producto", clear_on_submit=True):
                        st.write(f"### Detalle para: {producto_seleccionado} en {maletin_seleccionado}")
                        c1, c2 = st.columns(2)
                        with c1:
                            cantidad = st.number_input("¿Cuántos son? (Cantidad):", min_value=1, value=1, step=1)
                        with c2:
                            folio = st.text_input("Folio / Número de Serie:")
                        
                        observaciones = st.text_area("Observaciones del estado del producto/maletín:")
                        archivo_subido = st.file_uploader("Adjuntar resguardo/comprobante (Opcional):", type=["pdf", "png", "jpg", "jpeg"])
                        
                        btn_guardar = st.form_submit_button("💾 Guardar en Inventario")
                        
                        if btn_guardar:
                            if folio.strip():
                                enlace_archivo = "Sin archivo"
                                if archivo_subido is not None:
                                    file_metadata = {'name': f"{folio}_{archivo_subido.name}", 'parents': [folder_id]}
                                    media = MediaIoBaseUpload(io.BytesIO(archivo_subido.read()), mimetype=archivo_subido.type, resumable=True)
                                    archivo_drive = drive_service.files().create(body=file_metadata, media_body=media, fields='webViewLink').execute()
                                    enlace_archivo = archivo_drive.get('webViewLink')
                                
                                nueva_fila = [servidor_seleccionado, maletin_seleccionado, producto_seleccionado, cantidad, folio, observaciones, enlace_archivo]
                                worksheet_servidor.append_row(nueva_fila)
                                st.success(f"¡Se agregó **{cantidad}x {producto_seleccionado}** (Folio: {folio}) al **{maletin_seleccionado}**!")
                                st.rerun()
                            else:
                                st.warning("Por favor ingresa el Folio / Número de Serie del producto.")
            else:
                st.info("No hay Servidores de la Salud registrados. Agrega uno en la siguiente pestaña.")

        # === PESTAÑA 2: AGREGAR NUEVO SERVIDOR ===
        with tab_agregar_persona:
            st.subheader("➕ Agregar Nuevo Servidor de la Salud")
            with st.form("form_nuevo_servidor", clear_on_submit=True):
                nuevo_nombre = st.text_input("Nombre completo del Servidor de la Salud:")
                btn_agregar = st.form_submit_button("➕ Registrar Servidor")
                
                if btn_agregar:
                    if nuevo_nombre.strip():
                        if nuevo_nombre.strip() in lista_servidores:
                            st.warning("Este Servidor ya existe.")
                        else:
                            # Crear hoja nueva para esta persona
                            nueva_ws = sh_maestro.add_worksheet(title=nuevo_nombre.strip(), rows="100", cols="10")
                            nueva_ws.append_row(["SERVIDOR DE LA SALUD", "MALETÍN", "PRODUCTO", "CANTIDAD", "FOLIO", "OBSERVACIONES", "COMPROBANTE"])
                            
                            # Crear su carpeta en Drive
                            obtener_o_crear_carpeta(nuevo_nombre.strip(), CARPETA_PERSONAL_ID)
                            st.success(f"¡Servidor **{nuevo_nombre.strip()}** agregado exitosamente!")
                            st.rerun()
                    else:
                        st.warning("Escribe un nombre válido.")

        # === PESTAÑA 3: ELIMINAR SERVIDOR ===
        with tab_eliminar_persona:
            st.subheader("🗑️ Eliminar Servidor de la Salud")
            st.error("⚠️ Esta acción borrará la pestaña y registros vinculados a este Servidor.")
            
            if lista_servidores:
                servidor_a_eliminar = st.selectbox("Selecciona el Servidor a eliminar:", options=lista_servidores, key="select_eliminar")
                confirmar = st.checkbox(f"Confirmo que deseo borrar a **{servidor_a_eliminar}**")
                btn_eliminar = st.button("🗑️ Eliminar Definitivamente", type="primary")
                
                if btn_eliminar:
                    if confirmar:
                        try:
                            ws_del = sh_maestro.worksheet(servidor_a_eliminar)
                            sh_maestro.del_worksheet(ws_del)
                            st.success(f"Se ha eliminado a **{servidor_a_eliminar}**.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error al eliminar la hoja: {e}")
                    else:
                        st.warning("Marca la casilla de confirmación para proceder.")

    except Exception as e:
        st.error("Error al procesar la solicitud:")
        st.exception(e)
