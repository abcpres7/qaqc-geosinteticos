
import streamlit as st
import openpyxl
from datetime import datetime
import os

EXCEL_FILE = "datos.xlsx"

st.set_page_config(page_title="QA/QC Geosintéticos HDPE", layout="wide")

st.title("🛡️ Sistema de Control de Calidad - Geosintéticos HDPE")
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
st.sidebar.subheader("📥 Exportar y Gestionar")
if os.path.exists(EXCEL_FILE):
    with open(EXCEL_FILE, "rb") as f:
        st.sidebar.download_button(
            label="📥 Descargar Excel Actualizado",
            data=f,
            file_name="datos_actualizados_QAQC.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.sidebar.warning("Archivo Excel no encontrado.")

# ----------------- BOTÓN DE REINICIO / LIMPIEZA DIARIA TOTAL -----------------
st.sidebar.markdown("---")
with st.sidebar.expander("⚙️ Opciones Avanzadas"):
    st.warning("⚠️ Este botón limpia por completo los registros de **todas** las planillas para comenzar un nuevo día (recuerda descargar tu Excel antes).")
    if st.button("🧹 Limpiar Todas las Planillas (Nuevo Día)"):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            rangos_limpieza = {
                "Despliegue geosintéticos": (18, 70),
                "Despliegue geotextil": (18, 100),
                "Puesta en servicio cuña": (17, 70),
                "Puesta en servicio extrusora": (17, 50),
                "Soldadura por cuña": (17, 85),
                "Soldadura por extrusión": (17, 80),
                "Ensayos destructivos": (17, 60),
                "Aceptación Superficie": (15, 25)
            }
            
            for sheet_name, (fila_ini, fila_fin) in rangos_limpieza.items():
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    celdas_combinadas = ws.merged_cells.ranges
                    for r in range(fila_ini, fila_fin + 1):
                        for c in range(1, 22):
                            es_combinada = False
                            for merged_range in celdas_combinadas:
                                if (r >= merged_range.min_row and r <= merged_range.max_row and 
                                    c >= merged_range.min_col and c <= merged_range.max_col):
                                    es_combinada = True
                                    break
                            if not es_combinada:
                                ws.cell(row=r, column=c).value = None
            
            wb.save(EXCEL_FILE)
            st.sidebar.success("✅ ¡Todas las planillas fueron reiniciadas y limpiadas al 100%!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al limpiar: {e}")

if not os.path.exists(EXCEL_FILE):
    st.error(f"No se encontró el archivo base '{EXCEL_FILE}'. Asegúrate de tenerlo en la misma carpeta.")
    st.stop()

def encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1):
    """Busca la primera fila vacía en la columna clave dentro de la tabla oficial."""
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
                st.success(f"✅ Panel N° {num_panel} registrado correctamente en la fila {r}.")
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
                st.success(f"✅ Geotextil N° {num_panel} registrado en la fila {r}.")
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
                st.success(f"✅ Prueba inicial de cuña registrada en las filas {r} y {r+1}.")
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
                # Extrusora utiliza 1 fila por registro
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1)

                ws.cell(row=r, column=2, value=operador)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=str(hora))
                ws.cell(row=r, column=5, value=maquina_id)
                ws.cell(row=r, column=6, value=temp_pre)
                ws.cell(row=r, column=7, value=temp_ext)
                ws.cell(row=r, column=8, value=ensayado_por)

                # Desgarro (Columnas 9 a 13 / I a M)
                ws.cell(row=r, column=9, value=p1)
                ws.cell(row=r, column=10, value=p2)
                ws.cell(row=r, column=11, value=p3)
                ws.cell(row=r, column=12, value=p4)
                ws.cell(row=r, column=13, value=p5)

                # Cizalle (Columnas 14 a 18 / N a R)
                ws.cell(row=r, column=14, value=c1_v)
                ws.cell(row=r, column=15, value=c2_v)
                ws.cell(row=r, column=16, value=c3_v)
                ws.cell(row=r, column=17, value=c4_v)
                ws.cell(row=r, column=18, value=c5_v)

                # FTB y Ensayo OK (Columnas 19 y 20)
                ws.cell(row=r, column=19, value=ftb)
                ws.cell(row=r, column=20, value=ensayo_ok)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Prueba inicial de extrusora registrada en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 5. SOLDADURA POR CUÑA -----------------
elif protocolo == "Soldadura por cuña":
    st.subheader("📋 Protocolo: Registro de Uniones Cuña y Pruebas de Aire")
    with st.form("form_cuna"):
        c1, c2, c3 = st.columns(3)
        with c1:
            union_n = st.text_input("Unión Nº", "101")
            paneles = st.text_input("Ubicación / Paneles", "P01-HD - P02-HD")
            fecha = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            hora_ini = st.time_input("Hora Inicio Soldadura", datetime.now().time())
            hora_ter = st.time_input("Hora Término Soldadura", datetime.now().time())
            maquina = st.text_input("Máquina Nº", "0872")
        with c3:
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            temp = st.text_input("Temp. Cº", "460°")
            velocidad = st.text_input("Velocidad Cuña", "2,7")
            qc_asistente = st.text_input("Control Calidad", "Johnny Arias")

        submitted = st.form_submit_button("💾 Guardar Soldadura por Cuña")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Soldadura por cuña"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1)

                ws.cell(row=r, column=2, value=union_n)
                ws.cell(row=r, column=3, value=paneles)
                ws.cell(row=r, column=4, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=str(hora_ini))
                ws.cell(row=r, column=6, value=str(hora_ter))
                ws.cell(row=r, column=7, value=maquina)
                ws.cell(row=r, column=8, value=operador)
                ws.cell(row=r, column=9, value=temp)
                ws.cell(row=r, column=10, value=velocidad)
                ws.cell(row=r, column=11, value=qc_asistente)
                ws.cell(row=r, column=12, value=fecha.strftime("%Y-%m-%d"))

                wb.save(EXCEL_FILE)
                st.success(f"✅ Costura {union_n} registrada correctamente en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 6. SOLDADURA POR EXTRUSIÓN -----------------
elif protocolo == "Soldadura por extrusión":
    st.subheader("🔥 Protocolo: Soldadura por Extrusión y Pruebas de Vacío")
    with st.form("form_extrusion"):
        c1, c2, c3 = st.columns(3)
        with c1:
            parche_n = st.text_input("Parche Nº", "01")
            ubicacion = st.text_input("Ubicación / Paneles", "P02-HD - P17,16-LL")
            fecha = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            hora_sold = st.time_input("Hora Soldadura", datetime.now().time())
            maquina = st.text_input("Máquina Nº (Leister)", "0532")
            operador = st.text_input("Nombre Operador", "Jose Albarran")
        with c3:
            temp_ext = st.text_input("Temp. Extrusora °C", "230°")
            temp_pre = st.text_input("Temp. Precalentamiento", "180°")
            probado_por = st.text_input("Probado Por (QC)", "Johnny Arias")
            fecha_prueba = st.date_input("Fecha Prueba Vacío", datetime.now())
            hora_prueba = st.time_input("Hora Prueba Vacío", datetime.now().time())

        submitted = st.form_submit_button("💾 Guardar Soldadura Extrusión")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Soldadura por extrusión"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17, salto=1)

                ws.cell(row=r, column=2, value=parche_n)
                ws.cell(row=r, column=3, value=ubicacion)
                ws.cell(row=r, column=4, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=str(hora_sold))
                ws.cell(row=r, column=6, value=maquina)
                ws.cell(row=r, column=7, value=operador)
                ws.cell(row=r, column=8, value=temp_ext)
                ws.cell(row=r, column=9, value=temp_pre)
                ws.cell(row=r, column=10, value=probado_por)
                ws.cell(row=r, column=11, value=fecha_prueba.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=12, value=str(hora_prueba))

                wb.save(EXCEL_FILE)
                st.success(f"✅ Soldadura por extrusión {parche_n} guardada en la fila {r}.")
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
            temp = st.text_input("Temp (ºC)", "460°")
        with c3:
            res_1 = st.number_input("Probeta 1 (N)", value=480.0)
            res_2 = st.number_input("Probeta 2 (N)", value=480.0)
            res_3 = st.number_input("Probeta 3 (N)", value=503.0)

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
                ws.cell(row=r, column=10, value=res_1)
                ws.cell(row=r, column=11, value=res_2)
                ws.cell(row=r, column=12, value=res_3)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Ensayo destructivo {destructivo_n} registrado en las filas {r} y {r+1}.")
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
                st.success(f"✅ Aceptación de superficie guardada en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")
