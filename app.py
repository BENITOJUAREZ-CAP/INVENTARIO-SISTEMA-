import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io
from datetime import datetime

# Configuración de página
st.set_page_config(
    page_title="Control de Inventario por Maletines",
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

# Encabezados exactos según la imagen de A hasta P
ENCABEZADOS = [
    "Folio de producto", "Clave de Producto", "Producto", "Tipo", "Uso",
    "Planillas", "Fecha", "No. de Serie", "Caja", "Estatus",
    "Folio del Maletín", "Entidad de Envío", "Registrado Por",
    "Cantidad", "Servidor de la Salud", "Observaciones"
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
        
        # --- OBTENER LISTA DE SERVIDORES ---
        hojas = sh_maestro.worksheets()
        todas_las_hojas = [hoja.title for hoja in hojas]
        hojas_reservadas = ["PRODUCTOS", "PERSONAL DE SALUDA", "PERSONAL DE SALUD", "Hoja 1"]
        lista_servidores = [h for h in todas_las_hojas if h not in hojas_reservadas]
        
        # --- OBTENER CATÁLOGO DESDE LA PESTAÑA "PRODUCTOS" ---
        lista_productos = []
        try:
            ws_productos = sh_maestro.worksheet("PRODUCTOS")
            valores_productos = ws_productos.get_all_values()
            lista_productos = [fila[0].strip() for fila in valores_productos if fila and fila[0].strip()]
            lista_productos = [p for p in lista_productos if p.upper() not in ["PRODUCTO", "PRODUCTOS", "CONCEPTO"]]
        except Exception as e:
            st.warning("No se pudo cargar la pestaña 'PRODUCTOS'.")

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
                    worksheet_servidor = sh_maestro.add_worksheet(title=servidor_seleccionado, rows="100", cols="20")
                    worksheet_servidor.append_row(ENCABEZADOS)

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
                        st.write(f"### Capturar datos de: {producto_seleccionado}")
                        
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            folio_producto = st.text_input("Folio de producto:")
                            clave_producto = st.text_input("Clave de Producto:")
                            tipo_prod = st.selectbox("Tipo:", ["EQUIPO", "DESECHABLE", "MALETIN", "OTRO"])
                            uso_prod = st.selectbox("Uso:", ["ENFERMERA", "DERECHOHABIENTE", "GENERAL"])
                        with c2:
                            planillas = st.number_input("Planillas:", min_value=1, value=1000, step=100)
                            no_serie = st.text_input("No. de Serie (Opcional):")
                            caja = st.text_input("Caja (Opcional):")
                            estatus = st.selectbox("Estatus:", ["EN ALMACEN", "EN TRANSITO", "ENTREGADO"])
                        with c3:
                            folio_maletin = st.text_input("Folio del Maletín:")
                            entidad_envio = st.text_input("Entidad de Envío:", value="CIUDAD DE MEXICO")
                            registrado_por = st.text_input("Registrado Por:", value="INVENTARIO DE SALUD 20")
                            cantidad = st.number_input("Cantidad:", min_value=1, value=1, step=1)

                        observaciones = st.text_area("Observaciones:")
                        archivo_subido = st.file_uploader("Adjuntar comprobante (Opcional):", type=["pdf", "png", "jpg", "jpeg"])
                        
                        btn_guardar = st.form_submit_button("💾 Guardar en Inventario")
                        
                        if btn_guardar:
                            if folio_producto.strip():
                                fecha_actual = datetime.now().strftime("%m/%d/%y %H:%M")
                                
                                # Guardar comprobante en Drive si existe
                                if archivo_subido is not None:
                                    file_metadata = {'name': f"{folio_producto}_{archivo_subido.name}", 'parents': [folder_id]}
                                    media = MediaIoBaseUpload(io.BytesIO(archivo_subido.read()), mimetype=archivo_subido.type, resumable=True)
                                    drive_service.files().create(body=file_metadata, media_body=media, fields='webViewLink').execute()

                                # Fila ordenada exactamente de Columna A a P
                                nueva_fila = [
                                    folio_producto,        # A: Folio de producto
                                    clave_producto,        # B: Clave de Producto
                                    producto_seleccionado, # C: Producto
                                    tipo_prod,             # D: Tipo
                                    uso_prod,              # E: Uso
                                    planillas,             # F: Planillas
                                    fecha_actual,          # G: Fecha
                                    no_serie,              # H: No. de Serie
                                    caja,                  # I: Caja
                                    estatus,               # J: Estatus
                                    folio_maletin,         # K: Folio del Maletín
                                    entidad_envio,         # L: Entidad de Envío
                                    registrado_por,        # M: Registrado Por
                                    cantidad,              # N: Cantidad
                                    servidor_seleccionado, # O: Servidor de la Salud
                                    observaciones          # P: Observaciones
                                ]
                                
                                worksheet_servidor.append_row(nueva_fila)
                                st.success(f"¡Se agregó **{producto_seleccionado}** (Folio: {folio_producto}) correctamente!")
                                st.rerun()
                            else:
                                st.warning("Por favor ingresa al menos el Folio de producto.")
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
                            nueva_ws = sh_maestro.add_worksheet(title=nuevo_nombre.strip(), rows="100", cols="20")
                            nueva_ws.append_row(ENCABEZADOS)
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
