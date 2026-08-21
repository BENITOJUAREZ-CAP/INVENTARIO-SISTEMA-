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

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

ENCABEZADOS = [
    "Folio de producto", "Clave de Producto", "Producto", "Tipo", "Uso",
    "Planillas", "Fecha", "No. de Serie", "Caja", "Estatus",
    "Folio del Maletín", "Número de Maletín", "Entidad de Envío", "Registrado Por",
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

def formatear_encabezado_hoja(worksheet):
    fmt = {
        "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "textFormat": {
            "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
            "fontSize": 12,
            "bold": True
        }
    }
    worksheet.format("A1:Q1", fmt)

def obtener_o_crear_carpeta_y_excel(nombre_servidor, parent_folder_id):
    """
    Crea la carpeta personal en Drive y copia un Excel limpio dentro si no existe.
    """
    # 1. Crear u obtener la carpeta personal
    query_folder = f"'{parent_folder_id}' in parents and name = '{nombre_servidor}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results_folder = drive_service.files().list(q=query_folder, fields="files(id, name, webViewLink)").execute()
    items_folder = results_folder.get('files', [])
    
    if items_folder:
        folder_id = items_folder[0]['id']
        folder_link = items_folder[0]['webViewLink']
    else:
        folder_metadata = {
            'name': nombre_servidor,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = drive_service.files().create(body=folder_metadata, fields='id, webViewLink').execute()
        folder_id = folder.get('id')
        folder_link = folder.get('webViewLink')

    # 2. Verificar o copiar el Excel individual dentro de la carpeta
    nombre_excel_personal = f"INVENTARIO_{nombre_servidor}"
    query_sheet = f"'{folder_id}' in parents and name = '{nombre_excel_personal}' and trashed = false"
    results_sheet = drive_service.files().list(q=query_sheet, fields="files(id, name)").execute()
    items_sheet = results_sheet.get('files', [])

    if not items_sheet:
        try:
            # Duplicar el archivo maestro como plantilla evita el límite de cuota
            copied_file_body = {
                'name': nombre_excel_personal,
                'parents': [folder_id]
            }
            new_file = drive_service.files().copy(
                fileId=EXCEL_MAESTRO_ID,
                body=copied_file_body,
                fields='id'
            ).execute()
            
            # Limpiar el contenido copiado y dejar solo los encabezados
            nuevo_id = new_file.get('id')
            sh_nuevo = gc.open_by_key(nuevo_id)
            ws_nuevo = sh_nuevo.sheet1
            ws_nuevo.clear()
            ws_nuevo.append_row(ENCABEZADOS)
            formatear_encabezado_hoja(ws_nuevo)
        except Exception as e:
            st.warning(f"No se pudo duplicar el archivo Excel individual: {e}")

    return folder_id, folder_link

st.title("🧰 Control de Inventario por Maletines")

if gc and drive_service:
    try:
        sh_maestro = gc.open_by_key(EXCEL_MAESTRO_ID)
        
        hojas_objetos = sh_maestro.worksheets()
        nombres_hojas = [h.title for h in hojas_objetos]
        hojas_reservadas = ["PERSONAL DE SALUD", "PERSONAL DE SALUDA", "PRODUCTOS", "HOJA 1", "PLANTILLA"]
        
        lista_servidores = [
            h for h in nombres_hojas 
            if h.strip().upper() not in hojas_reservadas
        ]
        
        lista_productos = []
        if "PRODUCTOS" in [h.upper() for h in nombres_hojas]:
            try:
                ws_productos = sh_maestro.worksheet("PRODUCTOS")
                valores_productos = ws_productos.get_all_values()
                lista_productos = [fila[0].strip() for fila in valores_productos if fila and fila[0].strip()]
                lista_productos = [p for p in lista_productos if p.upper() not in ["PRODUCTO", "PRODUCTOS", "CONCEPTO"]]
            except Exception:
                pass

        tab_consultar, tab_agregar_persona, tab_eliminar_persona = st.tabs([
            "🔍 Consultar y Asignar Productos", 
            "➕ Agregar Nuevo Servidor",
            "❌ Eliminar Servidor"
        ])
        
        # === PESTAÑA 1: CONSULTA Y ASIGNACIÓN ===
        with tab_consultar:
            if lista_servidores:
                servidor_seleccionado = st.selectbox(
                    "1️⃣ Selecciona un Servidor de la Salud:", 
                    options=sorted(lista_servidores)
                )
                
                ws_servidor = sh_maestro.worksheet(servidor_seleccionado)
                valores_raw = ws_servidor.get_all_values()
                
                if len(valores_raw) > 1:
                    raw_cols = [str(col).strip().upper() for col in valores_raw[0]]
                    cleaned_cols = [c if c else f"COLUMNA_{i+1}" for i, c in enumerate(raw_cols)]
                    df_filtrado = pd.DataFrame(valores_raw[1:], columns=cleaned_cols).astype(str)
                    df_filtrado = df_filtrado.loc[:, ~df_filtrado.columns.duplicated()]
                else:
                    df_filtrado = pd.DataFrame(columns=ENCABEZADOS)
                
                st.write(f"### Inventario actual en la hoja de: **{servidor_seleccionado}**")
                st.dataframe(df_filtrado, use_container_width=True)
                
                col_folder, col_fmt = st.columns([2, 1])
                with col_folder:
                    folder_id, folder_link = obtener_o_crear_carpeta_y_excel(servidor_seleccionado, CARPETA_PERSONAL_ID)
                    st.markdown(f"📂 **Carpeta en Google Drive:** [Abrir carpeta personal de {servidor_seleccionado}]({folder_link})")
                with col_fmt:
                    if st.button("🎨 Aplicar formato a encabezados en Google Sheets"):
                        formatear_encabezado_hoja(ws_servidor)
                        st.success("¡Encabezados formateados con éxito!")

                st.divider()
                st.subheader("📋 Configuración del Maletín")
                
                col1, col2 = st.columns(2)
                with col1:
                    opciones_maletin = [f"Maletín {i}" for i in range(1, 21)]
                    maletin_num_seleccionado = st.selectbox(
                        "2️⃣ Selecciona el número de Maletín:", 
                        options=opciones_maletin
                    )
                with col2:
                    folio_maletin_input = st.text_input(
                        f"3️⃣ Ingresa el Folio/Clave para **{maletin_num_seleccionado}** (Columna K):",
                        placeholder="Ej. MAL-2024-001"
                    ).strip()

                col3, col4 = st.columns(2)
                with col3:
                    entidad_envio_gen = st.text_input(
                        "4️⃣ Entidad de Envío (Columna M):", 
                        value="CIUDAD DE MEXICO"
                    ).strip()
                with col4:
                    registrado_por_gen = st.text_input(
                        "5️⃣ Registrado Por (Columna N):", 
                        value="INVENTARIO DE SALUD 20"
                    ).strip()

                st.divider()

                if folio_maletin_input:
                    st.success(f"✅ Configuración lista: **{maletin_num_seleccionado}** | Folio: **{folio_maletin_input}** | Entidad: **{entidad_envio_gen}** | Registrado por: **{registrado_por_gen}**")
                    st.subheader("📦 Selección e Integración del Producto")
                    
                    if lista_productos:
                        producto_seleccionado = st.selectbox(
                            "Selecciona un producto del catálogo:", 
                            options=lista_productos
                        )
                    else:
                        producto_seleccionado = st.text_input("Nombre del producto del catálogo:")
                    
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
                            
                            c1, c2 = st.columns(2)
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

                            observaciones = st.text_area("Observaciones:")
                            archivo_subido = st.file_uploader("Adjuntar comprobante (Opcional):", type=["pdf", "png", "jpg", "jpeg"])
                            
                            btn_guardar = st.form_submit_button("💾 Guardar en Hoja del Servidor")
                            
                            if btn_guardar:
                                folios_limpios = [f.strip() for f in folios_ingresados if f.strip()]
                                if len(folios_limpios) == cantidad:
                                    fecha_actual = datetime.now().strftime("%m/%d/%y %H:%M")
                                    
                                    if archivo_subido is not None:
                                        try:
                                            nombre_archivo = f"{folios_limpios[0]}_{archivo_subido.name}"
                                            file_metadata = {'name': nombre_archivo, 'parents': [folder_id]}
                                            media = MediaIoBaseUpload(io.BytesIO(archivo_subido.read()), mimetype=archivo_subido.type, resumable=True)
                                            drive_service.files().create(body=file_metadata, media_body=media, fields='webViewLink').execute()
                                        except Exception as e_upload:
                                            st.warning(f"No se pudo guardar el archivo adjunto: {e_upload}")

                                    for folio in folios_limpios:
                                        val_folio_maletin = folio if (es_maletin or tipo_prod == "MALETIN") else folio_maletin_input

                                        nueva_fila = [
                                            str(folio),
                                            str(clave_producto),
                                            str(producto_seleccionado),
                                            str(tipo_prod),
                                            str(uso_prod),
                                            str(planillas),
                                            str(fecha_actual),
                                            str(no_serie),
                                            str(caja),
                                            str(estatus),
                                            str(val_folio_maletin),
                                            str(maletin_num_seleccionado),
                                            str(entidad_envio_gen),
                                            str(registrado_por_gen),
                                            "1",
                                            str(servidor_seleccionado),
                                            str(observaciones)
                                        ]
                                        ws_servidor.append_row(nueva_fila)

                                    st.success(f"¡Se guardaron {len(folios_limpios)} registro(s)!")
                                    st.rerun()
                                else:
                                    st.warning(f"Ingresa los {cantidad} folios requeridos.")
                else:
                    st.info("👈 Ingresa el **Folio/Clave del Maletín** para desbloquear los productos.")

            else:
                st.info("No se encontraron pestañas de Servidores de la Salud.")

        # === PESTAÑA 2: AGREGAR NUEVO SERVIDOR DE LA SALUD ===
        with tab_agregar_persona:
            st.subheader("➕ Crear Pestaña y Registrar Servidor")
            with st.form("form_nuevo_servidor", clear_on_submit=True):
                nuevo_nombre = st.text_input("Nombre completo del Servidor de la Salud:")
                btn_agregar = st.form_submit_button("➕ Registrar y Crear Hoja")
                
                if btn_agregar:
                    nombre_limpio = nuevo_nombre.strip().upper()
                    if nombre_limpio:
                        if nombre_limpio in [n.strip().upper() for n in nombres_hojas]:
                            st.warning("Ya existe un Servidor con este nombre.")
                        else:
                            try:
                                ws_indice = None
                                for sheet in sh_maestro.worksheets():
                                    title_clean = sheet.title.strip().upper()
                                    if "PERSONAL DE SALUD" in title_clean or "PERSONAL DE SALUDA" in title_clean:
                                        ws_indice = sheet
                                        break
                                
                                if ws_indice is None:
                                    ws_indice = sh_maestro.add_worksheet(title="PERSONAL DE SALUD", rows="100", cols="5")
                                    ws_indice.append_row(["PERSONAL DE SALUD"])
                                
                                col_a_values = ws_indice.col_values(1)
                                primera_fila_vacia = len(col_a_values) + 1
                                ws_indice.update_cell(primera_fila_vacia, 1, nombre_limpio)
                            except Exception as e_ind:
                                st.error(f"Error al escribir en la lista general: {e_ind}")

                            # Crear Carpeta + Excel duplicado en Drive
                            folder_id, folder_link = obtener_o_crear_carpeta_y_excel(nombre_limpio, CARPETA_PERSONAL_ID)
                            
                            # Crear Pestaña en Libro Maestro
                            nueva_hoja = sh_maestro.add_worksheet(title=nombre_limpio, rows="100", cols="20")
                            nueva_hoja.append_row(ENCABEZADOS)
                            formatear_encabezado_hoja(nueva_hoja)
                            
                            st.success(f"¡Se creó la carpeta y el Excel individual para **{nombre_limpio}** exitosamente!")
                            st.rerun()
                    else:
                        st.warning("Escribe un nombre válido.")

        # === PESTAÑA 3: ELIMINAR SERVIDOR DE LA SALUD ===
        with tab_eliminar_persona:
            st.subheader("❌ Eliminar Servidor de la Salud del Sistema")
            st.warning("⚠️ **Atención:** Esto borrará su nombre de 'PERSONAL DE SALUD' y eliminará su pestaña del Excel.")
            
            if lista_servidores:
                servidor_a_eliminar = st.selectbox(
                    "Selecciona el Servidor a eliminar:",
                    options=sorted(lista_servidores),
                    key="select_eliminar_servidor"
                )
                
                confirmacion = st.checkbox(
                    f"Confirmo que deseo eliminar a **{servidor_a_eliminar}**.",
                    key="chk_confirmar_eliminar"
                )
                
                if st.button("🗑️ Eliminar Definitivamente", type="primary", disabled=not confirmacion):
                    try:
                        try:
                            ws_indice = None
                            for sheet in sh_maestro.worksheets():
                                title_clean = sheet.title.strip().upper()
                                if "PERSONAL DE SALUD" in title_clean or "PERSONAL DE SALUDA" in title_clean:
                                    ws_indice = sheet
                                    break
                            
                            if ws_indice:
                                celda = ws_indice.find(servidor_a_eliminar)
                                if celda:
                                    ws_indice.delete_rows(celda.row)
                        except Exception:
                            pass

                        hoja_borrar = sh_maestro.worksheet(servidor_a_eliminar)
                        sh_maestro.del_worksheet(hoja_borrar)
                        
                        st.success(f"¡**{servidor_a_eliminar}** eliminado con éxito!")
                        st.rerun()
                    except Exception as err:
                        st.error(f"Error al eliminar: {err}")
            else:
                st.info("No hay Servidores registrados para eliminar.")

    except Exception as e:
        st.error("Error al procesar el libro de Google Sheets:")
        st.exception(e)
