from datetime import datetime
import pandas as pd
from google.oauth2.service_account import Credentials
import gspread
import streamlit as st

# Configuración inicial de la página
st.set_page_config(
    page_title="Control de Inventario por Maletines",
    page_icon="🧰",
    layout="wide",
)

# IDs de Google Sheets
EXCEL_MAESTRO_ID = "1Chjc0zz3T0qF6TaydjQxLa7bI12sT2ZS11gYl6aeun0"
EXCEL_MALETINES_ID = "1StjBMVIkueBy9sVy5dNh1yIpG5he-51iCYeeaj_mixM"

SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ENCABEZADOS = [
    "Folio de producto",
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
    "Número de Maletín",
    "Entidad de Envío",
    "Registrado Por",
    "Cantidad",
    "Servidor de la Salud",
    "Observaciones",
]


@st.cache_resource
def get_gspread_client():
  try:
    if "gcp_service_account" in st.secrets:
      creds_dict = dict(st.secrets["gcp_service_account"])
      creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
      gc = gspread.authorize(creds)
      return gc
    else:
      st.error("No se encontraron las credenciales en st.secrets.")
      return None
  except Exception as e:
    st.error(f"Error al conectar con Google Sheets: {e}")
    return None


gc = get_gspread_client()


def formatear_encabezado_hoja(worksheet):
  fmt = {
      "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
      "horizontalAlignment": "CENTER",
      "verticalAlignment": "MIDDLE",
      "textFormat": {
          "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
          "fontSize": 12,
          "bold": True,
      },
  }
  worksheet.format("A1:Q1", fmt)


st.title("🧰 Control de Inventario por Maletines")

if gc:
  try:
    # 1. Abrir libro de Maletines y Maestro
    sh_maletines = gc.open_by_key(EXCEL_MALETINES_ID)
    sh_maestro = gc.open_by_key(EXCEL_MAESTRO_ID)

    # 2. Cargar Catálogo de Productos desde la pestaña "PRODUCTOS"
    lista_productos = []
    try:
      ws_prod = sh_maletines.worksheet("PRODUCTOS")
      vals_prod = ws_prod.get_all_values()
      lista_productos = [
          f[0].strip()
          for f in vals_prod
          if f and f[0].strip() and f[0].strip().upper() != "PRODUCTO"
      ]
    except Exception:
      pass

    # 3. Cargar lista de Servidores de la Salud
    hojas_maestro = [h.title for h in sh_maestro.worksheets()]
    hojas_reservadas = [
        "PERSONAL DE SALUD",
        "PERSONAL DE SALUDA",
        "PRODUCTOS",
        "HOJA 1",
        "PLANTILLA",
    ]
    lista_servidores = [
        h for h in hojas_maestro if h.strip().upper() not in hojas_reservadas
    ]

    # Pestañas principales
    tab_registro_maletines, tab_asignacion, tab_gestion = st.tabs([
        "📦 1. Registrar Insumos en Maletín",
        "🔗 2. Asignar Maletín a Servidor",
        "👥 3. Gestión de Personal",
    ])

    # =========================================================
    # PESTAÑA 1: REGISTRAR INSUMOS EN MALETÍN
    # =========================================================
    with tab_registro_maletines:
      st.subheader("📝 Captura de Contenido del Maletín")
      opciones_maletines = [f"MALETIN {i}" for i in range(1, 31)]
      maletin_sel = st.selectbox(
          "Selecciona el Maletín a armar:", options=opciones_maletines
      )

      # Obtener o crear la hoja del maletín
      try:
        ws_maletin = sh_maletines.worksheet(maletin_sel)
      except Exception:
        ws_maletin = sh_maletines.add_worksheet(
            title=maletin_sel, rows="100", cols="20"
        )
        ws_maletin.append_row(ENCABEZADOS)
        formatear_encabezado_hoja(ws_maletin)

      # Mostrar tabla actual del maletín
      vals_m = ws_maletin.get_all_values()
      if len(vals_m) > 1:
        cols_m = [
            c if c else f"COL_{i+1}"
            for i, c in enumerate([str(x).strip().upper() for x in vals_m[0]])
        ]
        df_maletin = pd.DataFrame(vals_m[1:], columns=cols_m).astype(str)
        st.write(
            f"### Contenido actual registrado en: **{maletin_sel}**"
        )
        st.dataframe(df_maletin, use_container_width=True)
      else:
        st.info(
            f"El **{maletin_sel}** aún no tiene insumos o productos registrados."
        )

      st.divider()

      with st.form("form_agregar_insumo", clear_on_submit=True):
        st.markdown(f"#### ➕ Agregar Producto a **{maletin_sel}**")

        if lista_productos:
          prod_sel = st.selectbox(
              "Producto (Catálogo):", options=lista_productos
          )
        else:
          prod_sel = st.text_input("Nombre del Producto:")

        col1, col2, col3 = st.columns(3)
        with col1:
          folio_prod = st.text_input("Folio del Producto / Insumo:")
          clave_prod = st.text_input("Clave de Producto:")
          cantidad = st.number_input("Cantidad:", min_value=1, value=1, step=1)
        with col2:
          tipo_prod = st.selectbox(
              "Tipo:", ["EQUIPO", "DESECHABLE", "MALETIN", "OTRO"]
          )
          uso_prod = st.selectbox(
              "Uso:", ["ENFERMERA", "DERECHOHABIENTE", "GENERAL"]
          )
          estatus = st.selectbox(
              "Estatus:", ["EN ALMACEN", "EN TRANSITO", "ENTREGADO"]
          )
        with col3:
          folio_maletin_general = st.text_input(
              "Folio del Maletín:", placeholder="Ej. MAL-2026-001"
          )
          no_serie = st.text_input("No. de Serie (Opcional):")
          planillas = st.number_input(
              "Planillas:", min_value=0, value=1000, step=100
          )

        obs = st.text_area("Observaciones:")
        btn_guardar_insumo = st.form_submit_button(
            "💾 Guardar Producto en Maletín"
        )

        if btn_guardar_insumo:
          if folio_prod and prod_sel:
            fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M")
            nueva_fila = [
                str(folio_prod),
                str(clave_prod),
                str(prod_sel),
                str(tipo_prod),
                str(uso_prod),
                str(planillas),
                str(fecha_actual),
                str(no_serie),
                "",  # Caja
                str(estatus),
                str(folio_maletin_general),
                str(maletin_sel),
                "CIUDAD DE MEXICO",
                "INVENTARIO SALUD",
                str(cantidad),
                "SIN ASIGNAR",
                str(obs),
            ]
            ws_maletin.append_row(nueva_fila)
            st.success(
                f"¡Producto **{prod_sel}** agregado exitosamente a"
                f" **{maletin_sel}**!"
            )
            st.rerun()
          else:
            st.warning("Escribe el Folio y Nombre del Producto.")

    # =========================================================
    # PESTAÑA 2: ASIGNAR MALETÍN COMPLETO A SERVIDOR
    # =========================================================
    with tab_asignacion:
      st.subheader("🔗 Transferir Maletín Completo a un Servidor de la Salud")

      if lista_servidores:
        col_s, col_m = st.columns(2)
        with col_s:
          servidor_destino = st.selectbox(
              "1️⃣ Selecciona el Servidor de la Salud:",
              options=sorted(lista_servidores),
          )
        with col_m:
          hojas_maletines_existentes = [
              h.title
              for h in sh_maletines.worksheets()
              if h.title.upper() != "PRODUCTOS"
          ]
          if hojas_maletines_existentes:
            maletin_a_asignar = st.selectbox(
                "2️⃣ Selecciona el Maletín a Entregar:",
                options=sorted(hojas_maletines_existentes),
            )
          else:
            maletin_a_asignar = None

        if maletin_a_asignar:
          ws_origen = sh_maletines.worksheet(maletin_a_asignar)
          datos_origen = ws_origen.get_all_values()

          if len(datos_origen) > 1:
            df_m_preview = pd.DataFrame(
                datos_origen[1:], columns=datos_origen[0]
            )
            st.write(
                f"### Vista previa
