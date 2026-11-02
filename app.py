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
st.sidebar.subheader("📥 Exportar y Gestionar (Cloud Sync)")

def generar_excel_desde_nube():
    """Consulta la nube de Supabase y reconstruye el Excel con todos los datos sincronizados"""
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
                # Buscar la siguiente fila vacía
                fila = 17
                while ws.cell(row=fila, column=2).value is not None:
                    fila += 1
                
                # Escribir los datos dinámicamente en las columnas
                col = 2
                for k, v in datos.items():
                    ws.cell(row=fila, column=col, value=v)
                    col += 1

        wb.save(EXCEL_FILE)
        with open(EXCEL_FILE, "rb") as f:
            return f.read()
    except Exception as e:
        st.sidebar.error(f"Error al sincronizar con la nube para descarga: {e}")
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                return f.read()
        return None

excel_bytes = generar_excel_desde_nube()
if excel_bytes:
    st.sidebar.download_button(
        label="📥 Descargar Excel Sincronizado (Cloud)",
        data=excel_bytes,
        file_name="datos_actualizados_QAQC_sincronizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_cloud_sync"
    )
else:
    st.sidebar.warning("No se pudo generar el archivo desde la nube.")
