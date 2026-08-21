import datetime
from google.oauth2 import service_account
import gspread
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Inventario Centralizado",
    page_icon="📦",
    layout="wide",
)

# ---------------------------------------------------------
# CONFIGURACIÓN DE IDS DE GOOGLE DRIVE Y SHEETS
# ---------------------------------------------------------
EXCEL_MAESTRO_ID = "1Chjc0zz3T0qF6TaydjQxLa7bI12sT2ZS11gYl6aeun0"
SUBCARPETA_PERSONALES_ID = "1QVW-qYDtNGYX9CjFIjzD0isYTRFj2F-u"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COLUMNAS_PLANTILLA = [
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
    "Entidad de Envío",
    "Registrado Por",
    "Cantidad",
    "Servidor de la Salud",
    "Observaciones",
]


@st.cache_resource
def conectar_gspread():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)


st.title("📦 Inventario General y Carpetas Personales")

try:
    client = conectar_gspread()
    sh_maestro = client.open_by_key(EXCEL_MAESTRO_ID)

    todas_hojas = [ws.title for ws in sh_maestro.worksheets()]
    hoja_personal_nombre = todas_hojas[0] if todas_hojas else "Personal"
    hojas_inventario = [h for h in todas_hojas if h != hoja_personal_nombre]
except Exception as e:
    st.error(
        f"Error de conexión: {e}\n\nVerifica que la Service Account sea Editora de la carpeta principal."
    )
    st.stop()

tab1, tab2, tab3 = st.tabs(
    [
        "👥 Personal y Archivos Individuales",
        "🔍 Consultar Inventario",
        "➕ Registrar Producto",
    ]
)

# ---------------------------------------------------------
# TAB 1: CREAR SERVIDOR, EXCEL MAESTRO Y EXCEL INDIVIDUAL
# ---------------------------------------------------------
with tab1:
    st.subheader("👥 Control de Personal")

    ws_personal = sh_maestro.worksheet(hoja_personal_nombre)
    df_personal = pd.DataFrame(ws_personal.get_all_records())

    c1, c2 = st.columns([2, 1])

    with c1:
        st.markdown("### Servidores Registrados")
        if not df_personal.empty:
            st.dataframe(df_personal, use_container_width=True, hide_index=True)
        else:
            st.info("No hay registros en la lista general de personal.")

    with c2:
        st.markdown("### ➕ Registrar Nuevo Servidor")
        nuevo_nombre = st.text_input("Nombre Completo:")
        nuevo_cargo = st.text_input("Cargo / Función:")

        if st.button("Crear Registro + Excel Individual"):
            nombre_limpio = nuevo_nombre.strip().upper()
            if nombre_limpio:
                try:
                    # 1. Crear el Excel INDIVIDUAL dentro de la SUBCARPETA 'Inventario_personal'
                    nuevo_file = client.create(
                        nombre_limpio, folder_id=SUBCARPETA_PERSONALES_ID
                    )
                    sheet_indiv = nuevo_file.sheet1
                    sheet_indiv.append_row(COLUMNAS_PLANTILLA)
                    url_individual = f"https://docs.google.com/spreadsheets/d/{nuevo_file.id}"

                    # 2. Agregar a la Hoja 1 del Excel Maestro
                    ws_personal.append_row(
                        [nombre_limpio, nuevo_cargo.strip(), url_individual]
                    )

                    # 3. Crear pestaña individual en el Excel Maestro
                    if nombre_limpio not in todas_hojas:
                        nueva_ws = sh_maestro.add_worksheet(
                            title=nombre_limpio, rows="100", cols="20"
                        )
                        nueva_ws.append_row(COLUMNAS_PLANTILLA)

                    st.success(
                        f"✅ ¡**{nombre_limpio}** creado en la subcarpeta 'Inventario_personal' y sincronizado en el Maestro!"
                    )
                    st.rerun()
                except Exception as ex:
                    st.error(f"Error al crear los archivos: {ex}")
            else:
                st.error("Ingresa un nombre válido.")

# ---------------------------------------------------------
# TAB 2: CONSULTAR INVENTARIO
# ---------------------------------------------------------
with tab2:
    st.subheader("🔍 Consultar Registros")
    if hojas_inventario:
        persona_sel = st.selectbox(
            "Selecciona la persona:", hojas_inventario, key="sel_cons"
        )
        ws = sh_maestro.worksheet(persona_sel)
        df = pd.DataFrame(ws.get_all_records())

        if not df.empty:
            busqueda = st.text_input("Buscar por Serie, Folio, Producto:")
            if busqueda:
                b_str = str(busqueda).strip().lower()
                df = df[
                    df.apply(
                        lambda c: c.astype(str)
                        .str.lower()
                        .str.contains(b_str, na=False)
                    ).any(axis=1)
                ]

            st.dataframe(
                df.astype(str), use_container_width=True, hide_index=True
            )
        else:
            st.info("Sin registros de inventario asignados.")
    else:
        st.warning("No hay personal en el sistema.")

# ---------------------------------------------------------
# TAB 3: REGISTRAR PRODUCTO EN AMBOS ARCHIVOS
# ---------------------------------------------------------
with tab3:
    st.subheader("➕ Agregar Entrada de Inventario")

    if hojas_inventario:
        persona_destino = st.selectbox(
            "Asignar a Servidor de la Salud:",
            hojas_inventario,
            key="sel_reg",
        )

        with st.form("form_registro", clear_on_submit=True):
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                clave_prod = st.text_input("Clave de Producto*")
                producto = st.text_input("Producto*")
                tipo = st.text_input("Tipo")
                uso = st.text_input("Uso")
                planillas = st.text_input("Planillas")

            with col_b:
                fecha_reg = st.date_input("Fecha", datetime.date.today())
                no_serie = st.text_input("No. de Serie*")
                caja = st.text_input("Caja")
                estatus = st.selectbox(
                    "Estatus",
                    [
                        "Disponible",
                        "Enviado",
                        "En Mantenimiento",
                        "Baja",
                        "Otro",
                    ],
                )
                folio_maletin = st.text_input("Folio del Maletín")

            with col_c:
                entidad_envio = st.text_input("Entidad de Envío")
                registrado_por = st.text_input("Registrado Por")
                cantidad = st.number_input(
                    "Cantidad", min_value=1, value=1, step=1
                )
                observaciones = st.text_area("Observaciones")

            btn_guardar = st.form_submit_button("💾 Guardar Inventario")

            if btn_guardar:
                if not clave_prod or not producto or not no_serie:
                    st.error("Campos obligatorios (*): Clave, Producto y Serie.")
                else:
                    nueva_fila = [
                        str(clave_prod),
                        str(producto),
                        str(tipo),
                        str(uso),
                        str(planillas),
                        str(fecha_reg),
                        str(no_serie),
                        str(caja),
                        str(estatus),
                        str(folio_maletin),
                        str(entidad_envio),
                        str(registrado_por),
                        int(cantidad),
                        str(persona_destino),
                        str(observaciones),
                    ]

                    try:
                        # 1. Guardar en el Excel Maestro
                        ws_m = sh_maestro.worksheet(persona_destino)
                        ws_m.append_row(nueva_fila)

                        # 2. Guardar en su Excel Individual dentro de la subcarpeta
                        file_indiv = client.open(persona_destino)
                        file_indiv.sheet1.append_row(nueva_fila)

                        st.success(
                            f"✅ Guardado en el **Excel Maestro** y en el **Excel Individual de {persona_destino}**."
                        )
                    except Exception as ex:
                        st.error(f"Error al guardar: {ex}")
    else:
        st.warning("Registra al menos un servidor primero.")
