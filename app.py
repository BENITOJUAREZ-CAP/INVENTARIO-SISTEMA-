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
        
        # --- OBTENER HOJA PRINCIPAL DE DATOS ---
        hojas = sh_maestro.worksheets()
        nombres_hojas = [h.title for h in hojas]
        
        if "PERSONAL DE SALUD" in nombres_hojas:
            ws_inventario = sh_maestro.worksheet("PERSONAL DE SALUD")
        elif "Hoja 1" in nombres_hojas:
            ws_inventario = sh_maestro.worksheet("Hoja 1")
        else:
            ws_inventario = hojas[0]

        data_raw = ws_inventario.get_all_records()
        if data_raw:
            df_general = pd.DataFrame(data_raw).astype(str)
        else:
            df_general = pd.DataFrame(columns=ENCABEZADOS)

        col_servidor_nombre = "Servidor de la Salud"
        if col_servidor_nombre in df_general.columns:
            lista_servidores = sorted(list(set(df_general[col_servidor_nombre].dropna().unique())))
            lista_servidores = [s.strip() for s in lista_servidores if s.strip() and s.strip().upper() != "SERVIDOR DE LA SALUD"]
        else:
            lista_servidores = []

        # --- OBTENER CATÁLOGO DESDE LA PESTAÑA "PRODUCTOS" ---
        lista_productos = []
        try:
            ws_productos = sh_maestro.worksheet("PRODUCTOS")
            valores_productos = ws_productos.get_all_values()
            lista_productos = [fila[0].strip() for fila in valores_productos if fila and fila[0].strip()]
            lista_productos = [p for p in lista_productos if p.upper() not in ["PRODUCTO", "PRODUCTOS", "CONCEPTO"]]
        except Exception as e:
            st.warning("No se pudo cargar la pestaña 'PRODUCTOS'.")

        tab_consultar, tab_agregar_persona = st.tabs([
            "🔍 Consultar y Asignar Productos", 
            "➕ Agregar Nuevo Servidor de la Salud"
        ])
        
        # === PESTAÑA 1: CONSULTA Y ASIGNACIÓN ===
        with tab_consultar:
            if lista_servidores:
                servidor_seleccionado = st.selectbox(
                    "Selecciona un Servidor de la Salud:", 
                    options=lista_servidores
                )
                
                if col_servidor_nombre in df_general.columns:
                    df_filtrado = df_general[df_general[col_servidor_nombre].str.strip() == servidor_seleccionado]
                else:
                    df_filtrado = pd.DataFrame(columns=ENCABEZADOS)
                
                st.write(f"### Inventario actual registrado para: **{servidor_seleccionado}**")
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
                    # Detectar si el producto a capturar es de tipo Maletín
                    es_maletin = "MALETIN" in producto_seleccionado.upper()
                    
                    with st.form("form_confirmar_producto", clear_on_submit=True):
                        st.write(f"### Capturar datos de: {producto_seleccionado}")
                        
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            folio_producto = st.text_input("Folio de producto:")
                            clave_producto = st.text_input("Clave de Producto:")
                            tipo_defecto = "MALETIN" if es_maletin else "EQUIPO"
                            tipo_prod = st.selectbox("Tipo:", ["EQUIPO", "DESECHABLE", "MALETIN", "OTRO"], index=2 if es_maletin else 0)
                            uso_prod = st.selectbox("Uso:", ["ENFERMERA", "DERECHOHABIENTE", "GENERAL"])
                        with c2:
                            planillas = st.number_input("Planillas:", min_value=1, value=1000, step=100)
                            no_serie = st.text_input("No. de Serie (Opcional):")
                            caja = st.text_input("Caja (Opcional):")
                            estatus = st.selectbox("Estatus:", ["EN ALMACEN", "EN TRANSITO", "ENTREGADO"])
                        with c3:
                            if es_maletin or tipo_prod == "MALETIN":
                                st.info("ℹ️ Al ser un Maletín, el **Folio del Maletín** (Columna K) se asignará igual al **Folio de producto**.")
                                folio_maletin = ""  # Se asignará automáticamante al folio_producto
                            else:
                                folio_maletin = st.text_input("Folio del Maletín al que pertenece (Columna K):")
                                
                            entidad_envio = st.text_input("Entidad de Envío:", value="CIUDAD DE MEXICO")
                            registrado_por = st.text_input("Registrado Por:", value="INVENTARIO DE SALUD 20")
                            cantidad = st.number_input("Cantidad:", min_value=1, value=1, step=1)

                        observaciones = st.text_area("Observaciones:")
                        archivo_subido = st.file_uploader("Adjuntar comprobante (Opcional):", type=["pdf", "png", "jpg", "jpeg"])
                        
                        btn_guardar = st.form_submit_button("💾 Guardar en Inventario")
                        
                        if btn_guardar:
                            if folio_producto.strip():
                                fecha_actual = datetime.now().strftime("%m/%d/%y %H:%M")
                                
                                # Si el producto es un maletín, la Columna K toma el mismo valor que el Folio de producto
                                if es_maletin or tipo_prod == "MALETIN":
                                    val_folio_maletin = folio_producto.strip()
                                else:
                                    val_folio_maletin = folio_maletin.strip()

                                # Guardar comprobante en Drive si existe
                                if archivo_subido is not None:
                                    file_metadata = {'name': f"{folio_producto}_{archivo_subido.name}", 'parents': [folder_id]}
                                    media = MediaIoBaseUpload(io.BytesIO(archivo_subido.read()), mimetype=archivo_subido.type, resumable=True)
                                    drive_service.files().create(body=file_metadata, media_body=media, fields='webViewLink').execute()

                                # Fila ordenada exactamente de Columna A a P
                                nueva_fila = [
                                    str(folio_producto),        # A: Folio de producto
                                    str(clave_producto),        # B: Clave de Producto
                                    str(producto_seleccionado), # C: Producto
                                    str(tipo_prod),             # D: Tipo
                                    str(uso_prod),              # E: Uso
                                    str(planillas),             # F: Planillas
                                    str(fecha_actual),          # G: Fecha
                                    str(no_serie),              # H: No. de Serie
                                    str(caja),                  # I: Caja
                                    str(estatus),               # J: Estatus
                                    str(val_folio_maletin),     # K: Folio del Maletín
                                    str(entidad_envio),         # L: Entidad de Envío
                                    str(registrado_por),        # M: Registrado Por
                                    str(cantidad),              # N: Cantidad
                                    str(servidor_seleccionado), # O: Servidor de la Salud
                                    str(observaciones)          # P: Observaciones
                                ]
                                
                                ws_inventario.append_row(nueva_fila)
                                st.success(f"¡Se agregó **{producto_seleccionado}** para **{servidor_seleccionado}** correctamente!")
                                st.rerun()
                            else:
                                st.warning("Por favor ingresa al menos el Folio de producto.")
            else:
                st.info("No hay Servidores de la Salud con registros. Puedes registrar uno en la siguiente pestaña.")

        # === PESTAÑA 2: AGREGAR NUEVO SERVIDOR DE LA SALUD ===
        with tab_agregar_persona:
            st.subheader("➕ Registrar Nuevo Servidor de la Salud")
            with st.form("form_nuevo_servidor", clear_on_submit=True):
                nuevo_nombre = st.text_input("Nombre completo del Servidor de la Salud:")
                btn_agregar = st.form_submit_button("➕ Registrar Servidor y Crear Carpeta")
                
                if btn_agregar:
                    nombre_limpio = nuevo_nombre.strip().upper()
                    if nombre_limpio:
                        if nombre_limpio in [s.upper() for s in lista_servidores]:
                            st.warning("Este Servidor ya existe en la lista.")
                        else:
                            folder_id, folder_link = obtener_o_crear_carpeta(nombre_limpio, CARPETA_PERSONAL_ID)
                            st.success(f"¡Servidor **{nombre_limpio}** preparado correctamente! Ya puedes seleccionarlo en la lista para registrar sus productos.")
                            st.rerun()
                    else:
                        st.warning("Escribe un nombre válido.")

    except Exception as e:
        st.error("Error al procesar la solicitud:")
        st.exception(e)
