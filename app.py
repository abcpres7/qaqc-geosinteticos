from datetime import datetime, timedelta
import os
import openpyxl
import streamlit as st
from supabase import create_client, Client

# --- CREDENCIALES DE SUPABASE (Silenciosas y seguras) ---
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

# ----------------- BOTÓN DE REINICIO / LIMPIEZA DIARIA TOTAL -----------------
st.sidebar.markdown("---")
with st.sidebar.expander("⚙️ Opciones Avanzadas"):
    st.warning("⚠️ Este botón limpia por completo los registros de **todas** las planillas para comenzar un nuevo día.")
    if st.button("🧹 Limpiar Todas las Planillas (Nuevo Día)"):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            rangos_limpieza = {
                "Despliegue geosintéticos": (18, 100),
                "Despliegue geotextil": (18, 150),
                "Puesta en servicio cuña": (17, 100),
                "Puesta en servicio extrusora": (17, 100),
                "Soldadura por cuña": (17, 150),
                "Soldadura por extrusión": (17, 150),
                "Ensayos destructivos": (17, 100),
                "Aceptación Superficie": (15, 50)
            }
            
            for sheet_name, (fila_ini, fila_fin) in rangos_limpieza.items():
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    celdas_combinadas = ws.merged_cells.ranges
                    for r in range(fila_ini, fila_fin + 1):
                        for c in range(1, 25):
                            es_combinada = False
                            for merged_range in celdas_combinadas:
                                if (r >= merged_range.min_row and r <= merged_range.max_row and 
                                    c >= merged_range.min_col and c <= merged_range.max_col):
                                    es_combinada = True
                                    break
                            if not es_combinada:
                                ws.cell(row=r, column=c).value = None
            
            wb.save(EXCEL_FILE)
            
            if supabase is not None:
                try:
                    supabase.table("registros_qaqc").delete().neq("id", 0).execute()
                except Exception:
                    pass

            st.sidebar.success("✅ ¡Planillas y nube reiniciadas al 100%!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al limpiar: {e}")

if not os.path.exists(EXCEL_FILE):
    st.error(f"No se encontró el archivo base '{EXCEL_FILE}'. Asegúrate de tenerlo en la misma carpeta.")
    st.stop()

def encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1):
    fila = fila_inicio
    while ws.cell(row=fila, column=columna_clave).value is not None:
        fila += salto
    return fila

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
            hora_inst = st.time_input("Hora de Instalado", datetime.now().time())
            largo = st.number_input("Largo (m)", min_value=0.0, value=25.0, step=0.1)
            ancho = st.number_input("Ancho (m)", min_value=0.0, value=7.01, step=0.01)
        with c3:
            espesor = st.text_input("Espesor (mm)", "1.5")
            textura = st.selectbox("Textura", ["Doble Texturado", "Liso / Liso", "Texturado / Liso"])
            observaciones = st.text_input("Observaciones / Ubicación", "HDPE (Talud Colector principal)")
            inspector = st.text_input("Técnico Q.C.", "J.A.C.")

        submitted = st.form_submit_button("💾 Guardar Despliegue Geomembrana")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Despliegue geosintéticos"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=18, salto=1)
                area_panel = round(largo * ancho, 3)

                ws.cell(row=r, column=2, value=num_panel)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=rollo_n)
                ws.cell(row=r, column=5, value=str(hora_inst))
                ws.cell(row=r, column=6, value=largo)
                ws.cell(row=r, column=7, value=ancho)
                ws.cell(row=r, column=8, value=area_panel)
                ws.cell(row=r, column=9, value=espesor)
                ws.cell(row=r, column=10, value=textura)
                ws.cell(row=r, column=11, value=observaciones)
                ws.cell(row=r, column=12, value=inspector)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Despliegue geosintéticos", fecha.strftime("%Y-%m-%d"), {"num_panel": num_panel})
                st.success(f"✅ Panel N° {num_panel} registrado correctamente en la fila {r}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

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
            hora_inst = st.time_input("Hora de Instalado", datetime.now().time())
            largo = st.number_input("Largo (m)", min_value=0.0, value=20.0, step=0.1)
            ancho = st.number_input("Ancho (m)", min_value=0.0, value=4.6, step=0.1)
        with c3:
            observaciones = st.text_input("Observaciones", "Geotextil (Colector principal)")
            inspector = st.text_input("Técnico Q.C.", "J.A.C.")

        submitted = st.form_submit_button("💾 Guardar Despliegue Geotextil")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Despliegue geotextil"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=18, salto=1)
                area_panel = round(largo * ancho, 3)

                ws.cell(row=r, column=2, value=num_panel)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=rollo_n)
                ws.cell(row=r, column=5, value=str(hora_inst))
                ws.cell(row=r, column=6, value=largo)
                ws.cell(row=r, column=7, value=ancho)
                ws.cell(row=r, column=8, value=area_panel)
                ws.cell(row=r, column=11, value=observaciones)
                ws.cell(row=r, column=12, value=inspector)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Despliegue geotextil", fecha.strftime("%Y-%m-%d"), {"num_panel": num_panel})
                st.success(f"✅ Geotextil N° {num_panel} registrado en la fila {r}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 3. PUESTA EN SERVICIO CUÑA -----------------
elif protocolo == "Puesta en servicio cuña":
    st.subheader("⚙️ Protocolo: Pruebas Iniciales de Cuña (Tensiometro Examo)")
    with st.form("form_puesta_cuna"):
        c1, c2, c3 = st.columns(3)
        with c1:
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            fecha = st.date_input("Fecha Prueba", datetime.now())
            hora = st.time_input("Hora Prueba", datetime.now().time())
        with c2:
            maquina_id = st.text_input("Máquina N° / ID", "0872")
            temp = st.text_input("Temperatura (ºC)", "450°")
            vel = st.text_input("Velocidad Máquina", "3.0")
        with c3:
            ensayado_por = st.text_input("Ensayado Por (QC)", "Monica Lupa")

        st.markdown("---")
        st.markdown("### 🧪 Ensayos de Desgarro (N) - Probetas 1 a 5 (Valores A y B)")
        
        col_a1, col_b1 = st.columns(2)
        with col_a1: p1_sup = st.number_input("Probeta 1 - [A] (Superior)", value=497.0, step=1.0)
        with col_b1: p1_inf = st.number_input("Probeta 1 - [B] (Inferior)", value=526.0, step=1.0)

        col_a2, col_b2 = st.columns(2)
        with col_a2: p2_sup = st.number_input("Probeta 2 - [A] (Superior)", value=532.0, step=1.0)
        with col_b2: p2_inf = st.number_input("Probeta 2 - [B] (Inferior)", value=503.0, step=1.0)

        col_a3, col_b3 = st.columns(2)
        with col_a3: p3_sup = st.number_input("Probeta 3 - [A] (Superior)", value=509.0, step=1.0)
        with col_b3: p3_inf = st.number_input("Probeta 3 - [B] (Inferior)", value=503.0, step=1.0)

        col_a4, col_b4 = st.columns(2)
        with col_a4: p4_sup = st.number_input("Probeta 4 - [A] (Superior)", value=428.0, step=1.0)
        with col_b4: p4_inf = st.number_input("Probeta 4 - [B] (Inferior)", value=481.0, step=1.0)

        col_a5, col_b5 = st.columns(2)
        with col_a5: p5_sup = st.number_input("Probeta 5 - [A] (Superior)", value=439.0, step=1.0)
        with col_b5: p5_inf = st.number_input("Probeta 5 - [B] (Inferior)", value=497.0, step=1.0)

        st.markdown("---")
        st.markdown("### 🧪 Ensayos de Cizalle (N) - Probetas 1 a 5")
        cc1, cc2, cc3, cc4, cc5 = st.columns(5)
        with cc1: c1_v = st.number_input("Cizalle 1", value=658.0)
        with cc2: c2_v = st.number_input("Cizalle 2", value=595.0)
        with cc3: c3_v = st.number_input("Cizalle 3", value=555.0)
        with cc4: c4_v = st.number_input("Cizalle 4", value=595.0)
        with cc5: c5_v = st.number_input("Cizalle 5", value=612.0)

        st.markdown("---")
        cf1, cf2 = st.columns(2)
        with cf1: ftb = st.selectbox("FTB (Si / No)", ["SI", "NO"])
        with cf2: ensayo_ok = st.selectbox("Ensayo OK (Si / No)", ["SI", "NO"])

        submitted = st.form_submit_button("💾 Guardar Prueba Inicial Cuña")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Puesta en servicio cuña"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=2)

                ws.cell(row=r, column=2, value=operador)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=str(hora))
                ws.cell(row=r, column=5, value=maquina_id)
                ws.cell(row=r, column=6, value=temp)
                ws.cell(row=r, column=7, value=vel)
                ws.cell(row=r, column=8, value=ensayado_por)

                ws.cell(row=r, column=9, value=p1_sup)
                ws.cell(row=r, column=10, value=p2_sup)
                ws.cell(row=r, column=11, value=p3_sup)
                ws.cell(row=r, column=12, value=p4_sup)
                ws.cell(row=r, column=13, value=p5_sup)

                ws.cell(row=r+1, column=9, value=p1_inf)
                ws.cell(row=r+1, column=10, value=p2_inf)
                ws.cell(row=r+1, column=11, value=p3_inf)
                ws.cell(row=r+1, column=12, value=p4_inf)
                ws.cell(row=r+1, column=13, value=p5_inf)

                ws.cell(row=r, column=14, value=c1_v)
                ws.cell(row=r, column=15, value=c2_v)
                ws.cell(row=r, column=16, value=c3_v)
                ws.cell(row=r, column=17, value=c4_v)
                ws.cell(row=r, column=18, value=c5_v)

                ws.cell(row=r, column=19, value=ftb)
                ws.cell(row=r, column=20, value=ensayo_ok)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Puesta en servicio cuña", fecha.strftime("%Y-%m-%d"), {"operador": operador})
                st.success(f"✅ Prueba inicial de cuña registrada en las filas {r} y {r+1}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 4. PUESTA EN SERVICIO EXTRUSORA -----------------
elif protocolo == "Puesta en servicio extrusora":
    st.subheader("🔥 Protocolo: Pruebas Iniciales Extrusora")
    with st.form("form_puesta_ext"):
        c1, c2, c3 = st.columns(3)
        with c1:
            operador = st.text_input("Nombre Operador", "Rene Espinoza")
            fecha = st.date_input("Fecha Prueba", datetime.now())
            hora = st.time_input("Hora Prueba", datetime.now().time())
        with c2:
            maquina_id = st.text_input("Máquina N° (Leister)", "0532")
            temp_pre = st.text_input("Temp. Precalentamiento", "180°")
            temp_ext = st.text_input("Temperatura Extr. (ºC)", "230°")
        with c3:
            ensayado_por = st.text_input("Ensayado Por (QC)", "Brayan Flores")

        st.markdown("---")
        st.markdown("### 🧪 Ensayos de Desgarro (N) - Probetas 1 a 5 (Valor A)")
        
        de1, de2, de3, de4, de5 = st.columns(5)
        with de1: p1 = st.number_input("Probeta 1", value=446.0, step=1.0)
        with de2: p2 = st.number_input("Probeta 2", value=389.0, step=1.0)
        with de3: p3 = st.number_input("Probeta 3", value=395.0, step=1.0)
        with de4: p4 = st.number_input("Probeta 4", value=481.0, step=1.0)
        with de5: p5 = st.number_input("Probeta 5", value=332.0, step=1.0)

        st.markdown("---")
        st.markdown("### 🧪 Ensayos de Cizalle (N) - Probetas 1 a 5")
        cc_ex1, cc_ex2, cc_ex3, cc_ex4, cc_ex5 = st.columns(5)
        with cc_ex1: c1_v = st.number_input("Cizalle Ext. 1", value=492.0)
        with cc_ex2: c2_v = st.number_input("Cizalle Ext. 2", value=475.0)
        with cc_ex3: c3_v = st.number_input("Cizalle Ext. 3", value=544.0)
        with cc_ex4: c4_v = st.number_input("Cizalle Ext. 4", value=526.0)
        with cc_ex5: c5_v = st.number_input("Cizalle Ext. 5", value=520.0)

        st.markdown("---")
        cf_ex1, cf_ex2 = st.columns(2)
        with cf_ex1: ftb = st.selectbox("FTB (Si / No) - Extrusora", ["SI", "NO"])
        with cf_ex2: ensayo_ok = st.selectbox("Ensayo OK (Si / No) - Extrusora", ["SI", "NO"])

        submitted = st.form_submit_button("💾 Guardar Prueba Inicial Extrusora")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Puesta en servicio extrusora"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1)

                ws.cell(row=r, column=2, value=operador)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=str(hora))
                ws.cell(row=r, column=5, value=maquina_id)
                ws.cell(row=r, column=6, value=temp_pre)
                ws.cell(row=r, column=7, value=temp_ext)
                ws.cell(row=r, column=8, value=ensayado_por)

                ws.cell(row=r, column=9, value=p1)
                ws.cell(row=r, column=10, value=p2)
                ws.cell(row=r, column=11, value=p3)
                ws.cell(row=r, column=12, value=p4)
                ws.cell(row=r, column=13, value=p5)

                ws.cell(row=r, column=14, value=c1_v)
                ws.cell(row=r, column=15, value=c2_v)
                ws.cell(row=r, column=16, value=c3_v)
                ws.cell(row=r, column=17, value=c4_v)
                ws.cell(row=r, column=18, value=c5_v)

                ws.cell(row=r, column=19, value=ftb)
                ws.cell(row=r, column=20, value=ensayo_ok)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Puesta en servicio extrusora", fecha.strftime("%Y-%m-%d"), {"operador": operador})
                st.success(f"✅ Prueba inicial de extrusora registrada en la fila {r}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 5. SOLDADURA POR CUÑA -----------------
elif protocolo == "Soldadura por cuña":
    st.subheader("📋 Protocolo: Registro de Uniones Cuña y Pruebas de Aire")
    with st.form("form_cuna"):
        st.markdown("### 🛠️ Datos de Unión y Soldadura")
        c1, c2, c3 = st.columns(3)
        with c1:
            union_n = st.text_input("Unión Nº", "1")
            ubicacion = st.text_input("Ubicación / Paneles", "P1-P2")
            fecha_sold = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            hora_ini_sold = st.time_input("Hora Inicio Soldadura", datetime.strptime("12:56", "%H:%M").time())
            hora_ter_sold = st.time_input("Hora Término Soldadura", datetime.strptime("14:06", "%H:%M").time())
            maquina = st.text_input("Máquina Nº", "0872")
        with c3:
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            temp = st.text_input("Temp. Cº Cuña", "460°")
            velocidad = st.text_input("Velocidad Cuña", "3,0")
            qc_asistente = st.text_input("Probado Por (QC)", "Edgar Cruz")

        st.markdown("---")
        st.markdown("### 💨 Pruebas de Presión (Aire)")
        p1, p2, p3 = st.columns(3)
        with p1:
            fecha_prueba = st.date_input("Fecha Prueba Presión", datetime.now())
            hora_ini_psi = st.time_input("Hora Inicio PSI", datetime.strptime("14:30", "%H:%M").time())
        with p2:
            dummy_dt = datetime.combine(datetime.today(), hora_ini_psi) + timedelta(minutes=5)
            hora_ter_psi = dummy_dt.time()
            st.info(f"⏱️ Hora Término PSI (Auto +5 min): **{hora_ter_psi.strftime('%H:%M')}**")
            
            presion_ini = st.text_input("Presión Inicial (PSI)", "25")
        with p3:
            presion_ter = st.text_input("Presión Término (PSI)", "25")
            prueba_ok = st.selectbox("Prueba OK (Si / No)", ["SI", "NO"])

        st.markdown("---")
        observaciones = st.text_area("Observaciones", "Union Memb. Etapa 1")

        submitted = st.form_submit_button("💾 Guardar Soldadura por Cuña y Pruebas de Aire")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Soldadura por cuña"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1)

                ws.cell(row=r, column=2, value=union_n)
                ws.cell(row=r, column=3, value=ubicacion)
                ws.cell(row=r, column=4, value=fecha_sold.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=str(hora_ini_sold))
                ws.cell(row=r, column=6, value=str(hora_ter_sold))
                ws.cell(row=r, column=7, value=maquina)
                ws.cell(row=r, column=8, value=operador)
                ws.cell(row=r, column=9, value=temp)
                ws.cell(row=r, column=10, value=velocidad)
                ws.cell(row=r, column=11, value=qc_asistente)
                ws.cell(row=r, column=12, value=fecha_prueba.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=13, value=str(hora_ini_psi))
                ws.cell(row=r, column=14, value=str(hora_ter_psi))
                ws.cell(row=r, column=15, value=presion_ini)
                ws.cell(row=r, column=16, value=presion_ter)
                ws.cell(row=r, column=17, value=prueba_ok)
                ws.cell(row=r, column=18, value=observaciones)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Soldadura por cuña", fecha_sold.strftime("%Y-%m-%d"), {"union_n": union_n})
                st.success(f"✅ Costura {union_n} y prueba de aire registrada correctamente en la fila {r}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 6. SOLDADURA POR EXTRUSIÓN -----------------
elif protocolo == "Soldadura por extrusión":
    st.subheader("🔥 Protocolo: Soldadura por Extrusión y Pruebas de Vacío")
    with st.form("form_extrusion"):
        st.markdown("### 🛠️ Datos de Soldadura")
        c1, c2, c3 = st.columns(3)
        with c1:
            parche_n = st.text_input("Parche Nº", "1")
            ubicacion = st.text_input("Ubicación / Paneles", "P5-P6")
            fecha_sold = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            hora_sold = st.time_input("Hora Soldadura", datetime.strptime("12:54", "%H:%M").time())
            maquina = st.text_input("Máquina Nº (Leister)", "0532")
            operador = st.text_input("Nombre Operador", "Rene Espinoza")
        with c3:
            temp_ext = st.text_input("Temp. Extrusora °C", "230°")
            temp_pre = st.text_input("Temp. Precalentamiento", "180°")
            probado_por = st.text_input("Probado Por (QC)", "Diana Valdés")

        st.markdown("---")
        st.markdown("### 🔍 Prueba No Destructiva (Vacío / Spark Test)")
        p1, p2, p3 = st.columns(3)
        with p1:
            fecha_prueba = st.date_input("Fecha Prueba Vacío", datetime.now())
            hora_prueba = st.time_input("Hora Prueba Vacío", datetime.strptime("11:49", "%H:%M").time())
        with p2:
            resultado_ok = st.selectbox("Prueba No Destructiva OK", ["OK", "NO OK"])
            construccion = st.selectbox("Construcción / Reparación", ["Construcción", "Reparación"])
        with p3:
            obs_tipo = st.selectbox("Observaciones (PV o ST)", ["PV", "ST"])

        submitted = st.form_submit_button("💾 Guardar Soldadura Extrusión y Vacío")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Soldadura por extrusión"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1)

                ws.cell(row=r, column=2, value=parche_n)
                ws.cell(row=r, column=3, value=ubicacion)
                ws.cell(row=r, column=4, value=fecha_sold.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=str(hora_sold))
                ws.cell(row=r, column=6, value=maquina)
                ws.cell(row=r, column=7, value=operador)
                ws.cell(row=r, column=8, value=temp_ext)
                ws.cell(row=r, column=9, value=temp_pre)
                ws.cell(row=r, column=10, value=probado_por)
                ws.cell(row=r, column=11, value=fecha_prueba.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=12, value=str(hora_prueba))
                ws.cell(row=r, column=13, value=resultado_ok)
                ws.cell(row=r, column=14, value=construccion)
                ws.cell(row=r, column=15, value=obs_tipo)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Soldadura por extrusión", fecha_sold.strftime("%Y-%m-%d"), {"parche_n": parche_n})
                st.success(f"✅ Soldadura por extrusión {parche_n} guardada correctamente en la fila {r}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

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

        st.markdown("---")
        st.markdown("### 🧪 Ensayos de Desgarro (N) - Probetas 1 a 5 (Valores A y B)")
        
        col_ad1, col_bd1 = st.columns(2)
        with col_ad1: pd1_sup = st.number_input("Probeta 1 - [A] (Superior)", value=440.0, step=1.0)
        with col_bd1: pd1_inf = st.number_input("Probeta 1 - [B] (Inferior)", value=480.0, step=1.0)

        col_ad2, col_bd2 = st.columns(2)
        with col_ad2: pd2_sup = st.number_input("Probeta 2 - [A] (Superior)", value=480.0, step=1.0)
        with col_bd2: pd2_inf = st.number_input("Probeta 2 - [B] (Inferior)", value=480.0, step=1.0)

        col_ad3, col_bd3 = st.columns(2)
        with col_ad3: pd3_sup = st.number_input("Probeta 3 - [A] (Superior)", value=469.0, step=1.0)
        with col_bd3: pd3_inf = st.number_input("Probeta 3 - [B] (Inferior)", value=503.0, step=1.0)

        col_ad4, col_bd4 = st.columns(2)
        with col_ad4: pd4_sup = st.number_input("Probeta 4 - [A] (Superior)", value=469.0, step=1.0)
        with col_bd4: pd4_inf = st.number_input("Probeta 4 - [B] (Inferior)", value=434.0, step=1.0)

        col_ad5, col_bd5 = st.columns(2)
        with col_ad5: pd5_sup = st.number_input("Probeta 5 - [A] (Superior)", value=503.0, step=1.0)
        with col_bd5: pd5_inf = st.number_input("Probeta 5 - [B] (Inferior)", value=463.0, step=1.0)

        st.markdown("---")
        st.markdown("### 🧪 Ensayos de Cizalle (N) - Probetas 1 a 5")
        cc_d1, cc_d2, cc_d3, cc_d4, cc_d5 = st.columns(5)
        with cc_d1: cd1 = st.number_input("Cizalle Dest. 1", value=646.0)
        with cc_d2: cd2 = st.number_input("Cizalle Dest. 2", value=669.0)
        with cc_d3: cd3 = st.number_input("Cizalle Dest. 3", value=641.0)
        with cc_d4: cd4 = st.number_input("Cizalle Dest. 4", value=660.0)
        with cc_d5: cd5 = st.number_input("Cizalle Dest. 5", value=594.0)

        st.markdown("---")
        cf_d1, cf_d2 = st.columns(2)
        with cf_d1: ftb = st.selectbox("FTB (Si / No) - Destructivos", ["SI", "NO"])
        with cf_d2: ensayo_ok = st.selectbox("Ensayo OK (Si / No) - Destructivos", ["SI", "NO"])

        submitted = st.form_submit_button("💾 Guardar Ensayo Destructivo")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Ensayos destructivos"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=2)

                ws.cell(row=r, column=2, value=destructivo_n)
                ws.cell(row=r, column=3, value=union_paneles)
                ws.cell(row=r, column=4, value=fecha_sold.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=maquina_id)
                ws.cell(row=r, column=6, value=operador)
                ws.cell(row=r, column=7, value=temp)
                ws.cell(row=r, column=8, value=vel)
                ws.cell(row=r, column=9, value=ensayado_por)

                ws.cell(row=r, column=10, value=pd1_sup)
                ws.cell(row=r, column=11, value=pd2_sup)
                ws.cell(row=r, column=12, value=pd3_sup)
                ws.cell(row=r, column=13, value=pd4_sup)
                ws.cell(row=r, column=14, value=pd5_sup)

                ws.cell(row=r+1, column=10, value=pd1_inf)
                ws.cell(row=r+1, column=11, value=pd2_inf)
                ws.cell(row=r+1, column=12, value=pd3_inf)
                ws.cell(row=r+1, column=13, value=pd4_inf)
                ws.cell(row=r+1, column=14, value=pd5_inf)

                ws.cell(row=r, column=15, value=cd1)
                ws.cell(row=r, column=16, value=cd2)
                ws.cell(row=r, column=17, value=cd3)
                ws.cell(row=r, column=18, value=cd4)
                ws.cell(row=r, column=19, value=cd5)

                ws.cell(row=r, column=20, value=ftb)
                ws.cell(row=r, column=21, value=ensayo_ok)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Ensayos destructivos", fecha_sold.strftime("%Y-%m-%d"), {"destructivo_n": destructivo_n})
                st.success(f"✅ Ensayo destructivo {destructivo_n} registrado en las filas {r} y {r+1}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 8. ACEPTACIÓN SUPERFICIE -----------------
elif protocolo == "Aceptación Superficie":
    st.subheader("✅ Protocolo: Aceptación de Superficie de Subrasante")
    with st.form("form_superficie"):
        c1, c2 = st.columns(2)
        with c1:
            sector = st.text_input("Ítem / Área Evaluada", "Canaletas colectoras principal y secundarias")
            estado = st.selectbox("Condición", ["OK", "Rechazado"])
        with c2:
            desc = st.text_input("Observación / Descripción", "Presenta una excavación uniforme; fondo preparado para la recepción del revestimiento.")

        submitted = st.form_submit_button("💾 Guardar Aceptación de Superficie")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Aceptación Superficie"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=15, salto=1)

                ws.cell(row=r, column=2, value=sector)
                ws.cell(row=r, column=7, value=estado)
                ws.cell(row=r, column=9, value=desc)

                wb.save(EXCEL_FILE)
                guardar_en_nube("Aceptación Superficie", datetime.now().strftime("%Y-%m-%d"), {"sector": sector})
                st.success(f"✅ Aceptación de superficie guardada en la fila {r}.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")
