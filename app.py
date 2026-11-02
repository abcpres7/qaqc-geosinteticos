from datetime import datetime, timedelta
import os
import openpyxl
import streamlit as st
from supabase import create_client, Client

# --- CREDENCIALES DE SUPABASE ---
SUPABASE_URL = "https://fizrnuzuajkdrthzsliee.supabase.co".strip()
SUPABASE_KEY = "sb_publishable_AD17utXodOnFrdvb28kfvg_0B5hZNG1".strip()

# Inicializar conexión con Supabase de forma segura
try:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception:
    supabase = None

EXCEL_FILE = os.path.abspath("datos.xlsx")

st.set_page_config(page_title="QA/QC Geosintéticos HDPE", layout="wide")

st.title("🛡️ Sistema de Control de Calidad - Geosintéticos HDPE (Cloud Sinc)")
st.markdown("Registro oficial estructurado para protocolos QA/QC (Minera Las Cenizas).")

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
st.sidebar.subheader("📥 Exportar y Gestionar")

def generar_excel_desde_nube():
    try:
        if supabase is not None:
            response = supabase.table("registros_qaqc").select("*").execute()
            registros = response.data
        else:
            registros = []

        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE)
        else:
            return None

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
        with open(EXCEL_FILE, "rb") as f:
            return f.read()
    except Exception:
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                return f.read()
        return None

excel_bytes = generar_excel_desde_nube()
if excel_bytes:
    st.sidebar.download_button(
        label="📥 Descargar Excel Actualizado",
        data=excel_bytes,
        file_name="datos_actualizados_QAQC_cloud.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_cloud"
    )
else:
    st.sidebar.warning("No se pudo generar el archivo de descarga.")

st.sidebar.markdown("---")
with st.sidebar.expander("⚙️ Opciones Avanzadas"):
    st.warning("⚠️ Este botón limpia por completo los registros para comenzar un nuevo día.")
    if st.button("🧹 Limpiar Registros (Nuevo Día)"):
        try:
            if supabase is not None:
                supabase.table("registros_qaqc").delete().neq("id", 0).execute()
            st.sidebar.success("✅ ¡Registros limpiados al 100%!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al limpiar: {e}")

def guardar_en_nube(nombre_protocolo, fecha_val, diccionario_datos):
    try:
        if supabase is not None:
            payload = {
                "protocolo": nombre_protocolo,
                "fecha": str(fecha_val),
                "datos": diccionario_datos
            }
            supabase.table("registros_qaqc").insert(payload).execute()
        return True
    except Exception as e:
        st.error(f"Error al guardar en la nube: {e}")
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
            hora_inst = str(datetime.now().time())
            largo = st.number_input("Largo (m)", min_value=0.0, value=25.0, step=0.1)
            ancho = st.number_input("Ancho (m)", min_value=0.0, value=7.01, step=0.01)
        with c3:
            espesor = st.text_input("Espesor (mm)", "1.5")
            textura = st.selectbox("Textura", ["Doble Texturado", "Liso / Liso", "Texturado / Liso"])
            observaciones = st.text_input("Observaciones / Ubicación", "HDPE (Talud Colector principal)")
            inspector = st.text_input("Técnico Q.C.", "J.A.C.")

        submitted = st.form_submit_button("💾 Guardar Despliegue en la Nube")
        if submitted:
            area_panel = round(largo * ancho, 3)
            datos_registro = {
                "num_panel": num_panel, "fecha": fecha.strftime("%Y-%m-%d"), "rollo_n": rollo_n,
                "hora_inst": hora_inst, "largo": largo, "ancho": ancho, "area_panel": area_panel,
                "espesor": espesor, "textura": textura, "observaciones": observaciones, "inspector": inspector
            }
            if guardar_en_nube("Despliegue geosintéticos", fecha.strftime("%Y-%m-%d"), datos_registro):
                st.success(f"✅ Panel N° {num_panel} guardado con éxito.")
                st.rerun()

# ----------------- 2. DESPLIEGUE GEOTEXTIL -----------------
elif protocolo == "Despliegue geotextil":
    st.subheader("📋 Protocolo: Registro de Instalación de Geotextil")
    with st.form("form_geotextil"):
        c1, c2, c3 = st.columns(3)
        with c1:
            num_panel = st.text_input("N° de Panel Geotextil", "1")
            fecha = st.date_input("Fecha Instalación", datetime.now())
            rollo_n = st.text_input("N° de Rollo", "GT-01")
        with c2:
            hora_inst = str(datetime.now().time())
            largo = st.number_input("Largo (m)", min_value=0.0, value=20.0, step=0.1)
            ancho = st.number_input("Ancho (m)", min_value=0.0, value=4.6, step=0.1)
        with c3:
            observaciones = st.text_input("Observaciones", "Geotextil (Colector principal)")
            inspector = st.text_input("Técnico Q.C.", "J.A.C.")

        submitted = st.form_submit_button("💾 Guardar Geotextil en la Nube")
        if submitted:
            area_panel = round(largo * ancho, 3)
            datos_registro = {
                "num_panel": num_panel, "fecha": fecha.strftime("%Y-%m-%d"), "rollo_n": rollo_n,
                "hora_inst": hora_inst, "largo": largo, "ancho": ancho, "area_panel": area_panel,
                "observaciones": observaciones, "inspector": inspector
            }
            if guardar_en_nube("Despliegue geotextil", fecha.strftime("%Y-%m-%d"), datos_registro):
                st.success(f"✅ Geotextil N° {num_panel} guardado.")
                st.rerun()

# ----------------- 3. PUESTA EN SERVICIO CUÑA -----------------
elif protocolo == "Puesta en servicio cuña":
    st.subheader("⚙️ Protocolo: Pruebas Iniciales de Cuña (Tensiometro Examo)")
    with st.form("form_puesta_cuna"):
        c1, c2, c3 = st.columns(3)
        with c1:
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            fecha = st.date_input("Fecha Prueba", datetime.now())
            hora = str(datetime.now().time())
        with c2:
            maquina_id = st.text_input("Máquina N° / ID", "0872")
            temp = st.text_input("Temperatura (ºC)", "450°")
            vel = st.text_input("Velocidad Máquina", "3.0")
        with c3:
            ensayado_por = st.text_input("Ensayado Por (QC)", "Monica Lupa")

        submitted = st.form_submit_button("💾 Guardar Cuña en la Nube")
        if submitted:
            datos_registro = {
                "operador": operador, "fecha": fecha.strftime("%Y-%m-%d"), "hora": hora,
                "maquina_id": maquina_id, "temp": temp, "vel": vel, "ensayado_por": ensayado_por
            }
            if guardar_en_nube("Puesta en servicio cuña", fecha.strftime("%Y-%m-%d"), datos_registro):
                st.success("✅ Prueba inicial de cuña guardada.")
                st.rerun()

# ----------------- 4. PUESTA EN SERVICIO EXTRUSORA -----------------
elif protocolo == "Puesta en servicio extrusora":
    st.subheader("🔥 Protocolo: Pruebas Iniciales Extrusora")
    with st.form("form_puesta_ext"):
        c1, c2, c3 = st.columns(3)
        with c1:
            operador = st.text_input("Nombre Operador", "Rene Espinoza")
            fecha = st.date_input("Fecha Prueba", datetime.now())
            hora = str(datetime.now().time())
        with c2:
            maquina_id = st.text_input("Máquina N° (Leister)", "0532")
            temp_pre = st.text_input("Temp. Precalentamiento", "180°")
            temp_ext = st.text_input("Temperatura Extr. (ºC)", "230°")
        with c3:
            ensayado_por = st.text_input("Ensayado Por (QC)", "Brayan Flores")

        submitted = st.form_submit_button("💾 Guardar Extrusora en la Nube")
        if submitted:
            datos_registro = {
                "operador": operador, "fecha": fecha.strftime("%Y-%m-%d"), "hora": hora,
                "maquina_id": maquina_id, "temp_pre": temp_pre, "temp_ext": temp_ext, "ensayado_por": ensayado_por
            }
            if guardar_en_nube("Puesta en servicio extrusora", fecha.strftime("%Y-%m-%d"), datos_registro):
                st.success("✅ Prueba inicial de extrusora guardada.")
                st.rerun()

# ----------------- 5. SOLDADURA POR CUÑA -----------------
elif protocolo == "Soldadura por cuña":
    st.subheader("📋 Protocolo: Registro de Uniones Cuña y Pruebas de Aire")
    with st.form("form_cuna"):
        c1, c2, c3 = st.columns(3)
        with c1:
            union_n = st.text_input("Unión Nº", "1")
            ubicacion = st.text_input("Ubicación / Paneles", "P1-P2")
            fecha_sold = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            maquina = st.text_input("Máquina Nº", "0872")
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            temp = st.text_input("Temp. Cº Cuña", "460°")
        with c3:
            velocidad = st.text_input("Velocidad Cuña", "3,0")
            qc_asistente = st.text_input("Probado Por (QC)", "Edgar Cruz")
            prueba_ok = st.selectbox("Prueba OK (Si / No)", ["SI", "NO"])

        submitted = st.form_submit_button("💾 Guardar Soldadura Cuña en la Nube")
        if submitted:
            datos_registro = {
                "union_n": union_n, "ubicacion": ubicacion, "fecha_sold": fecha_sold.strftime("%Y-%m-%d"),
                "maquina": maquina, "operador": operador, "temp": temp, "velocidad": velocidad,
                "qc_asistente": qc_asistente, "prueba_ok": prueba_ok
            }
            if guardar_en_nube("Soldadura por cuña", fecha_sold.strftime("%Y-%m-%d"), datos_registro):
                st.success(f"✅ Costura {union_n} guardada.")
                st.rerun()

# ----------------- 6. SOLDADURA POR EXTRUSIÓN -----------------
elif protocolo == "Soldadura por extrusión":
    st.subheader("🔥 Protocolo: Soldadura por Extrusión y Pruebas de Vacío")
    with st.form("form_extrusion"):
        c1, c2, c3 = st.columns(3)
        with c1:
            parche_n = st.text_input("Parche Nº", "1")
            ubicacion = st.text_input("Ubicación / Paneles", "P5-P6")
            fecha_sold = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            maquina = st.text_input("Máquina Nº (Leister)", "0532")
            operador = st.text_input("Nombre Operador", "Rene Espinoza")
            temp_ext = st.text_input("Temp. Extrusora °C", "230°")
        with c3:
            temp_pre = st.text_input("Temp. Precalentamiento", "180°")
            probado_por = st.text_input("Probado Por (QC)", "Diana Valdés")
            resultado_ok = st.selectbox("Prueba No Destructiva OK", ["OK", "NO OK"])

        submitted = st.form_submit_button("💾 Guardar Extrusión en la Nube")
        if submitted:
            datos_registro = {
                "parche_n": parche_n, "ubicacion": ubicacion, "fecha_sold": fecha_sold.strftime("%Y-%m-%d"),
                "maquina": maquina, "operador": operador, "temp_ext": temp_ext, "temp_pre": temp_pre,
                "probado_por": probado_por, "resultado_ok": resultado_ok
            }
            if guardar_en_nube("Soldadura por extrusión", fecha_sold.strftime("%Y-%m-%d"), datos_registro):
                st.success(f"✅ Extrusión {parche_n} guardada.")
                st.rerun()

# ----------------- 7. ENSAYOS DESTRUCTIVOS -----------------
elif protocolo == "Ensayos destructivos":
    st.subheader("🧪 Protocolo: Registro de Ensayos Destructivos (ASTM D4437 / D6392)")
    with st.form("form_destructivos"):
        c1, c2, c3 = st.columns(3)
        with c1:
            destructivo_n = st.text_input("Destructivo N°", "1")
            union_paneles = st.text_input("Unión de paneles", "P1-P2")
            fecha_sold = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            maquina_id = st.text_input("Máquina N°", "0872")
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            temp = st.text_input("Temp (ºC) Cuña", "460°")
        with c3:
            vel = st.text_input("Velocidad Máquina", "3,0")
            ensayado_por = st.text_input("Ensayado Por (QC)", "Johnny Arias")
            ensayo_ok = st.selectbox("Ensayo OK (Si / No)", ["SI", "NO"])

        submitted = st.form_submit_button("💾 Guardar Destructivo en la Nube")
        if submitted:
            datos_registro = {
                "destructivo_n": destructivo_n, "union_paneles": union_paneles, "fecha_sold": fecha_sold.strftime("%Y-%m-%d"),
                "maquina_id": maquina_id, "operador": operador, "temp": temp, "vel": vel,
                "ensayado_por": ensayado_por, "ensayo_ok": ensayo_ok
            }
            if guardar_en_nube("Ensayos destructivos", fecha_sold.strftime("%Y-%m-%d"), datos_registro):
                st.success(f"✅ Ensayo destructivo {destructivo_n} guardado.")
                st.rerun()

# ----------------- 8. ACEPTACIÓN SUPERFICIE -----------------
elif protocolo == "Aceptación Superficie":
    st.subheader("✅ Protocolo: Aceptación de Superficie de Subrasante")
    with st.form("form_superficie"):
        c1, c2 = st.columns(2)
        with c1:
            sector = st.text_input("Ítem / Área Evaluada", "Canaletas colectoras principal y secundarias")
            estado = st.selectbox("Condición", ["OK", "Rechazado"])
        with c2:
            desc = st.text_input("Observación / Descripción", "Presenta una excavación uniforme; fondo preparado.")

        submitted = st.form_submit_button("💾 Guardar Aceptación en la Nube")
        if submitted:
            datos_registro = {
                "sector": sector, "estado": estado, "desc": desc
            }
            if guardar_en_nube("Aceptación Superficie", datetime.now().strftime("%Y-%m-%d"), datos_registro):
                st.success("✅ Aceptación de superficie guardada.")
                st.rerun()
