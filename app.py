from datetime import datetime, timedelta
import os
import openpyxl
import streamlit as st
from supabase import create_client, Client

# --- CREDENCIALES DE SUPABASE (Blindadas contra espacios) ---
SUPABASE_URL = "https://fizrnuzuajkdrthzsliee.supabase.co".strip()
SUPABASE_KEY = "sb_publishable_AD17utXodOnFrdvb28kfvg_0B5hZNG1".strip()

try:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception:
    supabase = None

def guardar_en_nube(nombre_protocolo, fecha_val, diccionario_datos):
    """Envía una copia a Supabase en segundo plano de forma silenciosa"""
    try:
        if supabase is not None:
            payload = {
                "protocolo": nombre_protocolo,
                "fecha": str(fecha_val),
                "datos": diccionario_datos
            }
            supabase.table("registros_qaqc").insert(payload).execute()
    except Exception:
        pass

# Usar ruta absoluta para garantizar estabilidad
EXCEL_FILE = os.path.abspath("datos.xlsx")

st.set_page_config(page_title="QA/QC Geosintéticos HDPE", layout="wide")

st.title("🛡️ Sistema de Control de Calidad - Geosintéticos HDPE (Cloud Sync)")
st.markdown("Registro oficial estructurado para protocolos QA/QC (Minera Las Cenizas).")

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
        "Aceptación Superficie"
    ]
)

st.sidebar.markdown("---")
st.sidebar.subheader("📥 Exportar y Sincronizar (Cloud)")

def generar_excel_desde_nube():
    """Consulta la nube de Supabase de forma segura y actualiza el Excel local"""
    try:
        registros = []
        if supabase is not None:
            response = supabase.table("registros_qaqc").select("*").execute()
            registros = response.data

        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
        else:
            return None

        # Volcar los registros de la nube en las hojas correspondientes del Excel
        for reg in registros:
            proto = reg.get("protocolo")
            datos = reg.get("datos", {})

            if proto in wb.sheetnames:
                ws = wb[proto]
                fila = 17
                while ws.cell(row=fila, column=2).value is not None:
                    fila += 1
                
                col = 2
                for k, v in datos.items():
                    ws.cell(row=fila, column=col, value=v)
                    col += 1

        wb.save(EXCEL_FILE)
    except Exception:
        pass # Si falla la red, no bloquea ni muestra errores visuales

    # Retorna siempre el archivo Excel local listo para descargar
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, "rb") as f:
            return f.read()
    return None

excel_bytes = generar_excel_desde_nube()
if excel_bytes:
    st.sidebar.download_button(
        label="📥 Descargar Excel Actualizado",
        data=excel_bytes,
        file_name="datos_actualizados_QAQC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_cloud_sync"
    )
else:
    st.sidebar.warning("No se encontró el archivo Excel local.")
