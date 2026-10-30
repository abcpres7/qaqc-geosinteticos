import streamlit as st
import openpyxl
from datetime import datetime
import os

# Nombre de tu archivo de Excel oficial en la carpeta
EXCEL_FILE = "datos.xlsx"

st.set_page_config(page_title="QA/QC Geosintéticos HDPE", layout="wide")

st.title("🛡️ Sistema de Registro QA/QC - Geosintéticos HDPE")
st.markdown("Carga automática y estructurada de datos hacia los protocolos oficiales de Calidad.")

# Selección del protocolo a llenar
protocolo = st.sidebar.selectbox(
    "Selecciona el Protocolo / Módulo",
    [
        "Despliegue geosintéticos",
        "Puesta en servicio cuña",
        "Soldadura por cuña",
        "Soldadura por extrusión",
        "Ensayos destructivos",
        "Aceptación Superficie"
    ]
)

# Cargar libro de Excel
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
    st.subheader("📋 Registro de Despliegue de Geosintéticos (Paneles / Rollos)")
    with st.form("form_despliegue"):
        c1, c2, c3 = st.columns(3)
        with c1:
            num_panel = st.text_input("N° de Panel (ej. 01, 02)", "01")
            fecha = st.date_input("Fecha Instalación", datetime.now())
            rollo_n = st.text_input("N° de Rollo", "R-001")
        with c2:
            hora_inst = st.time_input("Hora de Instalado", datetime.now().time())
            largo = st.number_input("Largo (m)", min_value=0.0, value=25.0, step=0.1)
            ancho = st.number_input("Ancho (m)", min_value=0.0, value=7.01, step=0.01)
        with c3:
            espesor = st.text_input("Espesor (mm)", "1.5")
            textura = st.selectbox("Textura", ["Doble Texturado", "Liso / Liso", "Texturado / Liso"])
            observaciones = st.text_input("Observaciones", "HDPE (Colector principal, Talud)")
            inspector = st.text_input("Técnico Q.C.", "J.A.C.")

        submitted = st.form_submit_button("💾 Guardar en Protocolo Despliegue")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Despliegue geosintéticos"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)
                
                area_panel = round(largo * ancho, 3)

                # Asignación exacta según columnas del protocolo oficial
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

# ----------------- 2. PUESTA EN SERVICIO CUÑA -----------------
elif protocolo == "Puesta en servicio cuña":
    st.subheader("⚙️ Puesta en Servicio - Máquina de Cuña")
    with st.form("form_puesta_cuna"):
        c1, c2, c3 = st.columns(3)
        with c1:
            operador = st.text_input("Nombre Operador", "Jose Albarran")
            fecha = st.date_input("Fecha Prueba", datetime.now())
        with c2:
            hora = st.time_input("Hora Prueba", datetime.now().time())
            maquina_id = st.text_input("Máquina N° / ID", "GeoStar G7")
            temp = st.text_input("Temperatura (ºC)", "450°")
        with c3:
            vel = st.text_input("Velocidad Máquina", "2.5")
            ensayado_por = st.text_input("Ensayado Por (QC)", "Johnny Arias")
            res_desgarro = st.number_input("Resistencia Desgarro (N)", value=550.0)

        submitted = st.form_submit_button("💾 Guardar Puesta en Servicio")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Puesta en servicio cuña"]
                r = encontrar_siguiente_fila(ws, columna_clave=1, fila_inicio=17)

                ws.cell(row=r, column=1, value=operador)
                ws.cell(row=r, column=2, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=3, value=str(hora))
                ws.cell(row=r, column=4, value=maquina_id)
                ws.cell(row=r, column=5, value=temp)
                ws.cell(row=r, column=6, value=vel)
                ws.cell(row=r, column=7, value=ensayado_por)
                ws.cell(row=r, column=8, value=res_desgarro)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Puesta en servicio registrada con éxito en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 3. SOLDADURA POR CUÑA -----------------
elif protocolo == "Soldadura por cuña":
    st.subheader("📋 Registro de Soldadura por Cuña")
    with st.form("form_cuna"):
        c1, c2, c3 = st.columns(3)
        with c1:
            union_n = st.text_input("Unión Nº", "101")
            paneles = st.text_input("Ubicación / Paneles", "P01-HD - P02-HD")
            fecha = st.date_input("Fecha", datetime.now())
        with c2:
            hora = st.time_input("Hora", datetime.now().time())
            maquina = st.text_input("Máquina Nº", "0872")
            operador = st.text_input("Nombre Operador", "Jose Albarran")
        with c3:
            temp = st.text_input("Temp. Cº", "460°")
            velocidad = st.text_input("Velocidad Cuña", "2,7")
            presion_aire = st.text_input("Prueba Aire / Observación", "Conforme (C)")

        submitted = st.form_submit_button("💾 Guardar Soldadura por Cuña")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Soldadura por cuña"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

                ws.cell(row=r, column=2, value=union_n)
                ws.cell(row=r, column=3, value=paneles)
                ws.cell(row=r, column=4, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=str(hora))
                ws.cell(row=r, column=6, value=maquina)
                ws.cell(row=r, column=7, value=operador)
                ws.cell(row=r, column=8, value=temp)
                ws.cell(row=r, column=9, value=velocidad)
                ws.cell(row=r, column=10, value=presion_aire)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Costura {union_n} registrada en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 4. SOLDADURA POR EXTRUSIÓN -----------------
elif protocolo == "Soldadura por extrusión":
    st.subheader("🔥 Registro de Soldadura por Extrusión")
    with st.form("form_extrusion"):
        c1, c2, c3 = st.columns(3)
        with c1:
            parche_n = st.text_input("Parche Nº", "01")
            ubicacion = st.text_input("Ubicación / Paneles", "P02-HD - P17,16-LL")
            fecha = st.date_input("Fecha", datetime.now())
        with c2:
            hora = st.time_input("Hora", datetime.now().time())
            maquina = st.text_input("Maquina Nº", "0532")
            operador = st.text_input("Nombre Operador", "Jose Albarran")
        with c3:
            temp_ext = st.text_input("Temperatura Extr. °C", "230°")
            temp_precal = st.text_input("Temp. Precalentamiento", "180°")
            probado_por = st.text_input("Probado Por (QC)", "Johnny Arias")

        submitted = st.form_submit_button("💾 Guardar Soldadura Extrusión")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Soldadura por extrusión"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

                ws.cell(row=r, column=2, value=parche_n)
                ws.cell(row=r, column=3, value=ubicacion)
                ws.cell(row=r, column=4, value=fecha.strftime("%Y-%m-%d"))
                ws.cell(row=r, column=5, value=str(hora))
                ws.cell(row=r, column=6, value=maquina)
                ws.cell(row=r, column=7, value=operador)
                ws.cell(row=r, column=8, value=temp_ext)
                ws.cell(row=r, column=9, value=temp_precal)
                ws.cell(row=r, column=10, value=probado_por)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Soldadura por extrusión {parche_n} guardada en fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 5. ENSAYOS DESTRUCTIVOS -----------------
elif protocolo == "Ensayos destructivos":
    st.subheader("🧪 Registro de Ensayos Destructivos (ASTM D6392)")
    with st.form("form_destructivos"):
        c1, c2, c3 = st.columns(3)
        with c1:
            destructivo_n = st.text_input("Destructivo N°", "ED-01")
            union_paneles = st.text_input("Unión de paneles", "P01 - P02")
            fecha_sold = st.date_input("Fecha Soldadura", datetime.now())
        with c2:
            maquina_id = st.text_input("Maquina N°", "0872")
            operador = st.text_input("Nombre operador", "Jose Albarran")
            temp = st.text_input("Temp (ºC)", "450°")
        with c3:
            vel = st.text_input("Velocidad Máquina", "2.5")
            ensayado_por = st.text_input("Ensayado Por", "Johnny Arias")
            res_cizalle = st.number_input("Resistencia Cizalle", value=125.0)

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
                ws.cell(row=r, column=8, value=vel)
                ws.cell(row=r, column=9, value=ensayado_por)
                ws.cell(row=r, column=10, value=res_cizalle)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Ensayo destructivo {destructivo_n} registrado en fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

# ----------------- 6. ACEPTACIÓN SUPERFICIE -----------------
elif protocolo == "Aceptación Superficie":
    st.subheader("✅ Protocolo de Aceptación de Superficie")
    with st.form("form_superficie"):
        c1, c2 = st.columns(2)
        with c1:
            sector = st.text_input("Sector / Área", "Canaletas colectoras principal y secundarias")
            estado = st.selectbox("Condición de Superficie", ["OK - Conforme", "Rechazado"])
        with c2:
            desc = st.text_input("Descripción / Observación", "Presenta una excavación uniforme; fondo preparado para la recepción del revestimiento.")

        submitted = st.form_submit_button("💾 Guardar Aceptación de Superficie")
        if submitted:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb["Aceptación Superficie"]
                r = encontrar_siguiente_fila(ws, columna_clave=2, fila_inicio=17)

                ws.cell(row=r, column=2, value=sector)
                ws.cell(row=r, column=7, value=estado)
                ws.cell(row=r, column=9, value=desc)

                wb.save(EXCEL_FILE)
                st.success(f"✅ Aceptación de superficie guardada en la fila {r}.")
            except Exception as e:
                st.error(f"Error al guardar: {e}")