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
EXCEL_MALETINES_ID = "1NCxdzjkr49t4pB_AzL5fhQ7ZfVmIVSHoUvDkpjK5flY"

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

      # Cargar maletines existentes
      hojas_maletines_existentes = [
          h.title
          for h in sh_maletines.worksheets()
          if h.title.upper() != "PRODUCTOS"
      ]

      col_m1, col_m2 = st.columns(2)
      with col_m1:
        # Permite elegir un maletín existente o ingresar un número/nombre nuevo sin límite
        opcion_maletin_input = st.text_input(
            "1️⃣ Número / Nombre del Maletín:",
            value=hojas_maletines_existentes[0] if hojas_maletines_existentes else "MALETIN 1",
            placeholder="Ej. MALETIN 1, MALETIN 50, etc.",
        ).strip().upper()

        if hojas_maletines_existentes:
          maletin_sel_dropdown = st.selectbox(
              "O selecciona de los maletines ya creados:",
              options=["-- Seleccionar existente --"] + sorted(hojas_maletines_existentes),
          )
          if maletin_sel_dropdown != "-- Seleccionar existente --":
            opcion_maletin_input = maletin_sel_dropdown

      with col_m2:
        folio_maletin_input = st.text_input(
            "2️⃣ Folio del Maletín (Columna K):",
            placeholder="Ej. MAL-2026-001",
        ).strip()

      maletin_sel = opcion_maletin_input

      if maletin_sel:
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
          st.write(f"### Contenido actual registrado en: **{maletin_sel}**")
          st.dataframe(df_maletin, use_container_width=True)
        else:
          st.info(f"El **{maletin_sel}** aún no tiene insumos o productos registrados.")

        st.divider()

        with st.form("form_agregar_insumo", clear_on_submit=True):
          st.markdown(f"#### ➕ Agregar Producto a **{maletin_sel}**")

          if lista_productos:
            prod_sel = st.selectbox("Producto (Catálogo):", options=lista_productos)
          else:
            prod_sel = st.text_input("Nombre del Producto:")

          col1, col2, col3 = st.columns(3)
          with col1:
            folio_prod = st.text_input("Folio del Producto / Insumo:")
            clave_prod = st.text_input("Clave de Producto:")
            cantidad = st.number_input("Cantidad:", min_value=1, value=1, step=1)
          with col2:
            tipo_prod = st.selectbox("Tipo:", ["EQUIPO", "DESECHABLE", "MALETIN", "OTRO"])
            uso_prod = st.selectbox("Uso:", ["ENFERMERA", "DERECHOHABIENTE", "GENERAL"])
            estatus = st.selectbox("Estatus:", ["EN ALMACEN", "EN TRANSITO", "ENTREGADO"])
          with col3:
            no_serie = st.text_input("No. de Serie (Opcional):")
            planillas = st.number_input("Planillas:", min_value=0, value=1000, step=100)
            caja = st.text_input("Caja (Opcional):")

          obs = st.text_area("Observaciones:")
          btn_guardar_insumo = st.form_submit_button("💾 Guardar Producto en Maletín")

          if btn_guardar_insumo:
            if folio_prod and prod_sel and folio_maletin_input:
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
                  str(caja),
                  str(estatus),
                  str(folio_maletin_input),
                  str(maletin_sel),
                  "CIUDAD DE MEXICO",
                  "INVENTARIO SALUD",
                  str(cantidad),
                  "SIN ASIGNAR",
                  str(obs),
              ]
              ws_maletin.append_row(nueva_fila)
              st.success(
                  f"¡Producto **{prod_sel}** agregado exitosamente a **{maletin_sel}** con Folio **{folio_maletin_input}**!"
              )
              st.rerun()
            else:
              st.warning("Por favor completa el Folio del Producto, Nombre del Producto y el Folio del Maletín.")

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
            df_m_preview = pd.DataFrame(datos_origen[1:], columns=datos_origen[0])
            st.write(f"### Vista previa del contenido de **{maletin_a_asignar}**")
            st.dataframe(df_m_preview, use_container_width=True)

            if st.button(f"🚀 Transferir {maletin_a_asignar} a {servidor_destino}", type="primary"):
              try:
                ws_destino = sh_maestro.worksheet(servidor_destino)

                filas_a_copiar = []
                for fila in datos_origen[1:]:
                  if len(fila) >= 16:
                    fila[15] = servidor_destino
                  filas_a_copiar.append(fila)

                for f in filas_a_copiar:
                  ws_destino.append_row(f)

                st.success(f"🎉 ¡El **{maletin_a_asignar}** fue asignado exitosamente a **{servidor_destino}**!")
              except Exception as ex_asig:
                st.error(f"Error al realizar la asignación: {ex_asig}")
          else:
            st.info(f"El **{maletin_a_asignar}** está vacío. Registra insumos en la Pestaña 1 primero.")
      else:
        st.info("No hay Servidores registrados en la base general.")

    # =========================================================
    # PESTAÑA 3: GESTIÓN DE PERSONAL
    # =========================================================
    with tab_gestion:
      st.subheader("➕ Agregar o Eliminar Servidores de la Salud")

      col_add, col_del = st.columns(2)

      with col_add:
        st.markdown("#### Registrar Nuevo Servidor")
        with st.form("form_nuevo_servidor", clear_on_submit=True):
          nuevo_nombre = st.text_input("Nombre completo:")
          btn_add = st.form_submit_button("➕ Crear Servidor")

          if btn_add:
            n_limpio = nuevo_nombre.strip().upper()
            if n_limpio:
              if n_limpio in [h.upper() for h in hojas_maestro]:
                st.warning("Ya existe un Servidor con este nombre.")
              else:
                nueva_h = sh_maestro.add_worksheet(title=n_limpio, rows="100", cols="20")
                nueva_h.append_row(ENCABEZADOS)
                formatear_encabezado_hoja(nueva_h)
                st.success(f"¡Servidor **{n_limpio}** creado exitosamente!")
                st.rerun()

      with col_del:
        st.markdown("#### Eliminar Servidor")
        if lista_servidores:
          servidor_del = st.selectbox(
              "Selecciona Servidor a eliminar:",
              options=sorted(lista_servidores),
          )
          chk_confirm = st.checkbox(f"Confirmar eliminación de {servidor_del}")

          if st.button("🗑️ Eliminar Definitivamente", disabled=not chk_confirm):
            hoja_b = sh_maestro.worksheet(servidor_del)
            sh_maestro.del_worksheet(hoja_b)
            st.success(f"¡**{servidor_del}** eliminado correctamente!")
            st.rerun()

  except Exception as e:
    st.error("Error general en el procesamiento de datos:")
    st.exception(e)
