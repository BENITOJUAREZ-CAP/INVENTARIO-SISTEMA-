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

# Encabezados exactos de Columna A a Columna P
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

        # Lectura segura de datos
        valores_raw = ws_inventario.get_all_values()
        
        if len(valores_raw) > 1:
            encabezados_hoja = [str(col).strip() for col in valores_raw[0]]
            filas_datos = valores_raw[1:]
            df_general = pd.DataFrame(filas_datos, columns=encabezados_hoja).astype(str)
        else:
            df_general = pd.DataFrame(columns=ENCABEZADOS)

        # Normalizar nombres de columnas
        columnas_normalizadas = {col: col.strip().upper() for col in df_general.columns}
        df_general.rename(columns=columnas_normalizadas, inplace=True)

        # Buscar la columna del Servidor de la Salud
        col_servidor = None
        for col in df_general.columns:
            if "SERVIDOR" in col and "SALUD" in col:
                col_servidor = col
                break

        if col_servidor:
            lista_servidores = sorted(list(set(df_general[col_servidor].str.strip().unique())))
            lista_servidores = [
                s for s in lista_servidores 
                if s and s.upper() not in ["NONE", "NAN", "", "SERVIDOR DE LA SALUD"]
            ]
        else:
            lista_servidores = []

        # --- OBTENER CATÁLOGO DESDE LA PESTAÑA "PRODUCTOS" ---
        lista_productos = []
        try:
            ws_productos = sh_maestro.worksheet("PRODUCTOS")
            valores_productos = ws_productos.get_all_values()
            lista_productos = [fila[0].strip() for fila in valores_productos if fila and fila[0].strip()]
            lista_productos = [p for p in lista_productos if p.upper() not in ["PRODUCTO", "PRODUCTOS", "CONCEPTO"]]
        except Exception:
            pass

        # --- PESTAÑAS DE LA APLICACIÓN ---
        tab_consultar, tab_agregar_persona = st.tabs([
            "🔍 Consultar y Asignar Productos", 
            "➕ Agregar Nuevo Servidor de la Salud"
        ])
        
        # === PESTAÑA 1: CONSULTA Y ASIGNACIÓN ===
        with tab_consultar:
            if lista_servidores:
                # PASO 1: Seleccionar Servidor de la Salud
                servidor_seleccionado = st.selectbox(
                    "1️⃣ Selecciona un Servidor de la Salud:", 
                    options=lista_servidores
                )
                
                # Filtrar registros del servidor seleccionado
                if col_servidor:
                    df_filtrado = df_general[df_general[col_servidor].str.strip() == servidor_seleccionado]
                else:
                    df_filtrado = pd.DataFrame(columns=ENCABEZADOS)
                
                st.write(f"### Inventario actual registrado para: **{servidor_seleccionado}**")
                st.dataframe(df_filtrado, use_container_width=True)
                
                folder_id, folder_link = obtener_o_crear_carpeta(servidor_seleccionado, CARPETA_PERSONAL_ID)
                st.markdown(f"📂 **Carpeta en Google Drive:** [Abrir carpeta personal de {servidor_seleccionado}]({folder_link})")
                
                st.divider()
                st.subheader("📋 Configuración del Maletín")
                
                # PASO 2 Y 3: Seleccionar Número de Maletín e Ingresar su Folio
                col1, col2 = st.columns(2)
                with col1:
                    opciones_maletin = [f"Maletín {i}" for i in range(1, 21)]  # Genera Maletín 1 hasta Maletín 20
                    maletin_num_seleccionado = st.selectbox(
                        "2️⃣ Selecciona el número de Maletín:", 
                        options=opciones_maletin
                    )
                with col2:
                    folio_maletin_input = st.text_input(
                        f"3️⃣ Ingresa el Folio/Clave para **{maletin_num_seleccionado}** (Columna K):",
                        placeholder="Ej. MAL-2024-001"
                    ).strip()

                st.divider()

                # PASO 4: Desplegar Productos solo cuando se haya ingresado el folio del maletín
                if folio_maletin_input:
                    st.success(f"✅ Configuración lista: **{maletin_num_seleccionado}** con Folio **{folio_maletin_input}** para **{servidor_seleccionado}**")
                    
                    st.subheader("📦 Selección e Integración del Producto")
                    
                    if lista_productos:
                        producto_seleccionado = st.selectbox(
                            "4️⃣ Selecciona un producto del catálogo:", 
                            options=lista_productos
                        )
                    else:
                        producto_seleccionado = st.text_input("4️⃣ Nombre del producto del catálogo:")
                    
                    if producto_seleccionado:
                        es_maletin = "MALETIN" in producto_seleccionado.upper()
                        
                        with st.form("form_confirmar_producto", clear_on_submit=True):
                            st.write(f"### Capturar datos de: **{producto_seleccionado}**")
                            
                            cantidad = st.number_input("Cantidad de este producto a registrar:", min_value=1, value=1, step=1)
                            
                            st.markdown("#### 🏷️ Folios de Producto")
                            folios_ingresados = []
                            cols_folios = st.columns(min(cantidad, 4))
                            for i in range(cantidad):
                                col_idx = i % 4
                                with cols_folios[col_idx]:
                                    f_val = st.text_input(f"Folio de producto #{i+1}:", key=f"folio_prod_{i}")
                                    folios_ingresados.append(f_val)
                            
                            st.divider()
                            
                            c1, c2, c3 = st.columns(3)
                            with c1:
                                clave_producto = st.text_input("Clave de Producto:")
                                tipo_prod = st.selectbox(
                                    "Tipo:", 
                                    ["EQUIPO", "DESECHABLE", "MALETIN", "OTRO"], 
                                    index=2 if es_maletin else 0
                                )
                                uso_prod = st.selectbox("Uso:", ["ENFERMERA", "DERECHOHABIENTE", "GENERAL"])
                            with c2:
                                planillas = st.number_input("Planillas:", min_value=1, value=1000, step=100)
                                no_serie = st.text_input("No. de Serie (Opcional):")
                                caja = st.text_input("Caja (Opcional):")
                                estatus = st.selectbox("Estatus:", ["EN ALMACEN", "EN TRANSITO", "ENTREGADO"])
                            with c3:
                                entidad_envio = st.text_input("Entidad de Envío:", value="CIUDAD DE MEXICO")
                                registrado_por = st.text_input("Registrado Por:", value="INVENTARIO DE SALUD 20")

                            observaciones = st.text_area("Observaciones:")
                            archivo_subido = st.file_uploader("Adjuntar comprobante (Opcional):", type=["pdf", "png", "jpg", "jpeg"])
                            
                            btn_guardar = st.form_submit_button("💾 Guardar en Inventario")
                            
                            if btn_guardar:
                                folios_limpios = [f.strip() for f in folios_ingresados if f.strip()]
                                if len(folios_limpios) == cantidad:
                                    fecha_actual = datetime.now().strftime("%m/%d/%y %H:%M")
                                    
                                    # Subir comprobante a Google Drive
                                    if archivo_subido is not None:
                                        nombre_archivo = f"{folios_limpios[0]}_{archivo_subido.name}"
                                        file_metadata = {'name': nombre_archivo, 'parents': [folder_id]}
                                        media = MediaIoBaseUpload(io.BytesIO(archivo_subido.read()), mimetype=archivo_subido.type, resumable=True)
                                        drive_service.files().create(body=file_metadata, media_body=media, fields='webViewLink').execute()

                                    # Agregar una fila por cada folio ingresado
                                    for folio in folios_limpios:
                                        val_folio_maletin = folio if (es_maletin or tipo_prod == "MALETIN") else folio_maletin_input

                                        nueva_fila = [
                                            str(folio),                 # A: Folio de producto
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
                                            "1",                        # N: Cantidad
                                            str(servidor_seleccionado), # O: Servidor de la Salud
                                            str(observaciones)          # P: Observaciones
                                        ]
                                        ws_inventario.append_row(nueva_fila)

                                    st.success(f"¡Se registraron exitosamente {len(folios_limpios)} unidad(es) de **{producto_seleccionado}** para {maletin_num_seleccionado} (Folio: {folio_maletin_input})!")
                                    st.rerun()
                                else:
                                    st.warning(f"Ingresa los {cantidad} folios requeridos (llevas {len(folios_limpios)} de {cantidad}).")
                else:
                    st.info("👈 Ingresa el **Folio/Clave del Maletín** en el paso 3 para desbloquear y desplegar la lista de productos del catálogo.")

            else:
                st.info("No se encontraron registros de Servidores de la Salud en la hoja. Agrega uno nuevo en la siguiente pestaña.")

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
                            
                            fecha_actual = datetime.now().strftime("%m/%d/%y %H:%M")
                            fila_inicial = [
                                "N/A", "N/A", "REGISTRO INICIAL", "SISTEMA", "N/A",
                                "0", fecha_actual, "N/A", "N/A", "ALTA",
                                "N/A", "CIUDAD DE MEXICO", "SISTEMA", "0",
                                nombre_limpio, "Registro de Servidor"
                            ]
                            ws_inventario.append_row(fila_inicial)
                            
                            st.success(f"¡Servidor **{nombre_limpio}** registrado con éxito!")
                            st.rerun()
                    else:
                        st.warning("Escribe un nombre válido.")

    except Exception as e:
        st.error("Error al procesar los datos de la hoja:")
        st.exception(e)
