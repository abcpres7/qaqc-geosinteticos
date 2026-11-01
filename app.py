
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

# ----------------- BOTÓN DE REINICIO / LIMPIEZA DIARIA -----------------
st.sidebar.markdown("---")
with st.sidebar.expander("⚙️ Opciones Avanzadas"):
    st.warning("⚠️ El botón de abajo borra los registros de las tablas para comenzar un nuevo día (recuerda descargar tu Excel antes).")
    if st.button("🧹 Limpiar Planilla (Nuevo Día)"):
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            # Definimos rangos de filas de datos a limpiar por hoja (sin tocar membretes ni firmas)
            rangos_limpieza = {
                "Despliegue geosintéticos": (18, 56),
                "Despliegue geotextil": (18, 90),
                "Puesta en servicio cuña": (17, 58),
                "Puesta en servicio extrusora": (17, 36),
                "Soldadura por cuña": (17, 78),
                "Soldadura por extrusión": (17, 68),
                "Ensayos destructivos": (17, 42),
                "Aceptación Superficie": (15, 22)
            }
            
            for sheet_name, (fila_ini, fila_fin) in rangos_limpieza.items():
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    for r in range(fila_ini, fila_fin + 1):
                        for c in range(1, ws.max_column + 1):
                            ws.cell(row=r, column=c).value = None
            
            wb.save(EXCEL_FILE)
            st.sidebar.success("✅ ¡Planilla reiniciada con éxito para el nuevo día!")
            st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al limpiar: {e}")

if not os.path.exists(EXCEL_FILE):
    st.error(f"No se encontró el archivo base '{EXCEL_FILE}'. Asegúrate de tenerlo en la misma carpeta.")
    st.stop()

def encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17):
    """Busca la primera fila vacía en la columna clave dentro de la tabla oficial."""
    fila = fila_inicio
    while ws.cell(row=fila, column=columna_clave).value is not None:
        fila += 1
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
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=18)
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
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=18)
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
            temp = st.text_input("Temperatura (ºC)", "460°")
            vel = st.text_input("Velocidad Máquina", "2,7")
        with c3:
            ensayado_por = st.text_input("Ensayado Por (QC)", "Johnny Arias")
            p1 = st.number_input("Probeta 1 (N)", value=550.0)
            p2 = st.number_input("Probeta 2 (N)", value=530.0)
            p3 = st.number_input("Probeta 3 (N)", value=540.0)

        submitted = st.form_submit_button("💾 Guardar Prueba Inicial Cuña")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Puesta en servicio cuña"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

                ws.cell(row=r, column=2, value=operador)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=str(hora))
                ws.cell(row=r, column=5, value=maquina_id)
                ws.cell(row=r, column=6, value=temp)
                ws.cell(row=r, column=7, value=vel)
                ws.cell(row=r, column=8, value=ensayado_por)
                ws.cell(row=r, column=9, value=p1)
                ws.cell(row=r, column=10, value=p2)
                ws.cell(row=r, column=11, value=p3)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Prueba inicial de cuña registrada en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 4. PUESTA EN SERVICIO EXTRUSORA -----------------
elif protocolo == "Puesta en servicio extrusora":
    st.subheader("🔥 Protocolo: Pruebas Iniciales Extrusora")
    with st.form("form_puesta_ext"):
        c1, c2, c3 = st.columns(3)
        with c1:
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            fecha = st.date_input("Fecha Prueba", datetime.now())
            hora = st.time_input("Hora Prueba", datetime.now().time())
        with c2:
            maquina_id = st.text_input("Máquina N° (Leister)", "0532")
            temp_ext = st.text_input("Temperatura Extr. (ºC)", "230°")
            temp_pre = st.text_input("Temp. Precalentamiento", "180°")
        with c3:
            ensayado_por = st.text_input("Ensayado Por (QC)", "Diana Valdés")
            p1 = st.number_input("Probeta 1 (N)", value=538.0)
            p2 = st.number_input("Probeta 2 (N)", value=531.0)

        submitted = st.form_submit_button("💾 Guardar Prueba Inicial Extrusora")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Puesta en servicio extrusora"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

                ws.cell(row=r, column=2, value=operador)
                ws.cell(row=r, column=3, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=4, value=str(hora))
                ws.cell(row=r, column=5, value=maquina_id)
                ws.cell(row=r, column=6, value=temp_pre)
                ws.cell(row=r, column=7, value=temp_ext)
                ws.cell(row=r, column=8, value=ensayado_por)
                ws.cell(row=r, column=9, value=p1)
                ws.cell(row=r, column=10, value=p2)

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
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

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
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

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
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

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
                st.success(f"✅ Ensayo destructivo {destructivo_n} registrado en la fila {r}.")
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
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=15)

                ws.cell(row=r, column=2, value=sector)
                ws.cell(row=r, column=7, value=estado)
                ws.cell(row=r, column=9, value=desc)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Aceptación de superficie guardada en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")
