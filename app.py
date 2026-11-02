from datetime import datetime, timedelta
import os
import openpyxl
import streamlit as st
from supabase import create_client, Client

# --- CREDENCIALES DE SUPABASE (Pega aquí tus datos del bloc de notas) ---
SUPABASE_URL = "PEGAR_AQUÍ_TU_PROJECT_URL"
SUPABASE_KEY = "PEGAR_AQUÍ_TU_PUBLISHABLE_KEY"

# Inicializar conexión con Supabase
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

EXCEL_FILE = os.path.abspath("datos.xlsx")

st.set_page_config(page_title="QA/QC Geosintéticos HDPE", layout="wide")

st.title("🛡️ Sistema de Control de Calidad - Geosintéticos HDPE (Cloud Sinc)")
st.markdown(
    "Registro oficial estructurado para protocolos QA/QC (Minera Las Cenizas)."
)

# ----------------- MENÚ LATERAL Y DESCARGA -----------------
protocolo = st.sidebar.selectbox(
    "Selecciona el Protocolo / Módulo",
    [
        "Despliegue geosintéticos",
        "Despliegue geotextil",
        "Puesta en servicio cuña",
        "Puesta en servicio extrusora",
        "Soldadura por cuña",
        "Soldadura por extrusión",
        "Ensayos destructivos",
        "Aceptación Superficie",
    ],
)

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Exportar y Gestionar")


# Función para generar el Excel descargable desde la base de datos en la nube
def generar_excel_desde_nube():
  try:
    # Consultar todos los registros guardados en Supabase
    response = supabase.table("registros_qaqc").select("*").execute()
    registros = response.data

    if os.path.exists(EXCEL_FILE):
      wb = openpyxl.load_workbook(EXCEL_FILE)
    else:
      return None

    # Rellenar las hojas dinámicamente con los datos de la nube
    # (El sistema mantiene tu estructura oficial intacta)
    for reg in registros:
      proto = reg.get("protocolo")
      datos = reg.get("datos", {})

      if proto in wb.sheetnames:
        ws = wb[proto]
        # Buscamos la siguiente fila vacía
        fila = 17
        while ws.cell(row=fila, column=2).value is not None:
          fila += 1

        # Volcar las llaves del diccionario en las columnas correspondientes
        col = 2
        for k, v in datos.items():
          ws.cell(row=fila, column=col, value=v)
          col += 1

    wb.save(EXCEL_FILE)
    with open(EXCEL_FILE, "rb") as f:
      return f.read()
  except Exception as e:
    st.error(f"Error sincronizando con la nube para descarga: {e}")
    if os.path.exists(EXCEL_FILE):
      with open(EXCEL_FILE, "rb") as f:
        return f.read()
    return None


excel_bytes = generar_excel_desde_nube()
if excel_bytes:
  st.sidebar.download_button(
      label="📥 Descargar Excel Actualizado (Cloud)",
      data=excel_bytes,
      file_name="datos_actualizados_QAQC_cloud.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      key="btn_download_cloud",
  )
else:
  st.sidebar.warning("No se pudo generar el archivo de descarga.")

# ----------------- BOTÓN DE REINICIO / LIMPIEZA DIARIA TOTAL -----------------
st.sidebar.markdown("---")
with st.sidebar.expander("⚙️ Opciones Avanzadas"):
  st.warning(
      "⚠️ Este botón limpia por completo los registros de la base de datos en"
      " la nube para comenzar un nuevo día."
  )
  if st.button("🧹 Limpiar Base de Datos (Nuevo Día)"):
    try:
      # Borrar registros en Supabase
      supabase.table("registros_qaqc").delete().neq("id", 0).execute()
      st.sidebar.success("✅ ¡Base de datos en la nube limpiada al 100%!")
      st.rerun()
    except Exception as e:
      st.sidebar.error(f"Error al limpiar la nube: {e}")


def guardar_en_nube(nombre_protocolo, fecha_val, diccionario_datos):
  """Función estándar para enviar cualquier registro directamente a Supabase"""
  try:
    payload = {
        "protocolo": nombre_protocolo,
        "fecha": str(fecha_val),
        "datos": diccionario_datos,
    }
    supabase.table("registros_qaqc").insert(payload).execute()
    return True
  except Exception as e:
    st.error(f"Error al guardar en Supabase: {e}")
    return False


# ----------------- 1. DESPLIEGUE GEOSINTÉTICOS -----------------
if protocolo == "Despliegue geosintéticos":
  st.subheader("📋 Protocolo: Registro de Instalación de Geomembrana HDPE")
  with st.form("form_despliegue"):
    c1, c2, c3 = st.columns(3)
    with c1:
      num_panel = st.text_input("N° de Panel", "01")
      fecha = st.date_input("Fecha de Instalación", datetime.now())
      rollo_n = st.text_input("N° de Rollo", "R-001")
    with c2:
      hora_inst = st.str(datetime.now().time())
      largo = st.number_input("Largo (m)", min_value=0.0, value=25.0, step=0.1)
      ancho = st.number_input("Ancho (m)", min_value=0.0, value=7.01, step=0.01)
    with c3:
      espesor = st.text_input("Espesor (mm)", "1.5")
      textura = st.selectbox(
          "Textura", ["Doble Texturado", "Liso / Liso", "Texturado / Liso"]
      )
      observaciones = st.text_input(
          "Observaciones / Ubicación", "HDPE (Talud Colector principal)"
      )
      inspector = st.text_input("Técnico Q.C.", "J.A.C.")

    submitted = st.form_submit_button("💾 Guardar Despliegue en la Nube")
    if submitted:
      area_panel = round(largo * ancho, 3)
      datos_registro = {
          "num_panel": num_panel,
          "fecha": fecha.strftime("%Y-%m-%d"),
          "rollo_n": rollo_n,
          "hora_inst": str(datetime.now().time()),
          "largo": largo,
          "ancho": ancho,
          "area_panel": area_panel,
          "espesor": espesor,
          "textura": textura,
          "observaciones": observaciones,
          "inspector": inspector,
      }
      if guardar_en_nube(
          "Despliegue geosintéticos", fecha.strftime("%Y-%m-%d"), datos_registro
      ):
        st.success(
            f"✅ Panel N° {num_panel} guardado y sincronizado en la nube con"
            " éxito."
        )
        st.rerun()

# [Los demás protocolos pueden seguir la misma estructura de guardado en nube]
st.markdown("---")
st.info(
    "💡 Tu aplicación ahora está conectada a Supabase. Realiza una prueba"
    " guardando un registro desde tu celular y verás cómo aparece de"
    " inmediato en tu computador."
)
