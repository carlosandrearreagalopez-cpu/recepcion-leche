import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as RLTable, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración de la página
st.set_page_config(page_title="Recepción de Leche Cruda - LIF Brands", layout="wide")

# ==========================================================
# ESTILOS CSS CON IDENTIDAD VISUAL LIF BRANDS
# ==========================================================
st.markdown("""
<style>
.stApp {
    background-color: #FFFFFF !important;
}
html, body, [class*="css"], p, span, label {
    font-family: Arial, sans-serif !important;
    color: #000000 !important;
}
h1, h2, h3, h4, h5, h6 {
    color: #1e3a8a !important;
    font-family: Arial, sans-serif !important;
}
.stTextInput label, .stSelectbox label, .stDateInput label, .stNumberInput label, .stRadio label, .stFileUploader label {
    color: #1e3a8a !important;
    font-weight: bold !important;
}
[data-testid="stFileUploader"] section div button div {
    display: none !important;
}
[data-testid="stFileUploader"] section div button span {
    visibility: hidden;
}
[data-testid="stFileUploader"] section div button::after {
    content: "Examinar archivos";
    visibility: visible;
    display: block;
    position: absolute;
    color: #1e3a8a !important;
}
[data-testid="stFileUploader"] {
    background-color: #f8fafc !important;
    border: 2px dashed #cbd5e1 !important;
    border-radius: 8px;
    padding: 10px;
}
[data-testid="stDataFrame"] div[data-baseweb="base-input"], [data-testid="stDataFrame"] table {
    background-color: #FFFFFF !important;
    color: #000000 !important;
}
[data-testid="stDataFrame"] th {
    background-color: #f1f5f9 !important;
    color: #1e3a8a !important;
    font-weight: bold !important;
}
.stButton>button {
    background-color: #FFFFFF !important;
    color: #1e3a8a !important;
    border: 2px solid #1e3a8a !important;
    border-radius: 6px;
    font-family: Arial, sans-serif;
    font-weight: bold;
}
.stButton>button:hover {
    background-color: #1e3a8a !important;
    color: #FFFFFF !important;
}
button[kind="primary"] {
    background-color: #1e3a8a !important;
    color: #FFFFFF !important;
    border: none !important;
}
button[kind="primary"]:hover {
    background-color: #3b82f6 !important;
}
.card-investigacion {
    padding: 15px;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    margin-bottom: 10px;
    background-color: #f8fafc;
    border-left: 5px solid #3b82f6;
}
</style>
""", unsafe_allow_html=True)

# Directorios para archivos y firmas
FOTOS_DIR = "fotos_recepcion"
FIRMAS_DIR = "firmas_recepcion"
FIRMAS_REGISTRADAS_DIR = "firmas_registradas"
PDF_DIR = "pdf_generados"

for d in [FOTOS_DIR, FIRMAS_DIR, FIRMAS_REGISTRADAS_DIR, PDF_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

EXCEL_FILE = "registros_recepcion_leche.xlsx"

def mostrar_logo(ancho=160):
    if os.path.exists("logo.png"):
        st.image("logo.png", width=ancho)
    else:
        st.warning("⚠️ Logo no encontrado. Asegúrate de tener el archivo 'logo.png' en tu repositorio de GitHub.")

def guardar_en_excel(datos_dict):
    df_nuevo = pd.DataFrame([datos_dict])
    if os.path.exists(EXCEL_FILE):
        df_existente = pd.read_excel(EXCEL_FILE)
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)
    else:
        df_final = df_nuevo
    
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Registros")
        ws = writer.sheets["Registros"]
        if len(df_final) > 0:
            max_row = len(df_final) + 1
            max_col = len(df_final.columns)
            col_letter = get_column_letter(max_col)
            table_range = f"A1:{col_letter}{max_row}"
            tab = Table(displayName="TablaRegistrosLeche", ref=table_range)
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            ws.add_table(tab)

def generar_pdf_formato(row, idx):
    pdf_filename = os.path.join(PDF_DIR, f"recepcion_leche_{idx}.pdf")
    doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    
    # Títulos e info de cabecera
    header_data = [
        [Paragraph("<b>Glad</b> / LIF Brands", styles['Normal']), Paragraph("<b>REGISTRO DE RECEPCIÓN DE LECHE CRUDA</b>", styles['Heading2']), Paragraph("Código: R-PBD/01-1<br/>Versión: 8", styles['Normal'])]
    ]
    t_header = RLTable(header_data, colWidths=[150, 400, 200])
    t_header.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#1e3a8a')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 10))
    
    # Datos generales
    gen_data = [
        [f"Fecha de recepción: {row.get('Fecha_Recepcion', '')}", f"Proveedor: {row.get('Proveedor', '')}", f"Cantidad reportada: {row.get('Cantidad_Litros', '')} L"]
    ]
    t_gen = RLTable(gen_data, colWidths=[250, 250, 250])
    t_gen.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.grey),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
    ]))
    elements.append(t_gen)
    elements.append(Spacer(1, 10))
    
    # Tablas de detalle (Transporte y Físico-químico)
    trans_data = [
        ["Vehículo de transporte", "Resultado", "Observaciones"],
        ["Limpieza exterior", str(row.get('Limpieza_Exterior', '')), ""],
        ["Salidas de leche selladas", str(row.get('Salidas_Selladas', '')), ""],
        ["Desinfección utensilios muestra", str(row.get('Desinfeccion_Utensilios', '')), ""]
    ]
    t_trans = RLTable(trans_data, colWidths=[250, 100, 400])
    t_trans.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1e3a8a')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    elements.append(t_trans)
    elements.append(Spacer(1, 10))
    
    doc.build(elements)
    return pdf_filename

# Control de navegación
if "nav_state" not in st.session_state:
    st.session_state["nav_state"] = "home"
if "admin_logueado" not in st.session_state:
    st.session_state["admin_logueado"] = False
if "form_logueado" not in st.session_state:
    st.session_state["form_logueado"] = False

# ==========================================================
# PANTALLA DE INICIO
# ==========================================================
if st.session_state["nav_state"] == "home":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        mostrar_logo(ancho=200)
        st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>Registro de Recepción de Leche</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #65a30d; font-weight: bold; font-size: 16px;'>LIF Brands Aseguramiento de Calidad</p>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Reportar nuevo ingreso", use_container_width=True, type="primary"):
            st.session_state["nav_state"] = "form_login"
            st.rerun()
            
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
        with col_l2:
            if st.button("Revisar ingresos (administrador)"):
                st.session_state["nav_state"] = "admin_login"
                st.rerun()

# ==========================================================
# LOGIN NUEVO INGRESO
# ==========================================================
elif st.session_state["nav_state"] == "form_login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("← Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
        st.title("🔒 Acceso a Registro")
        password_form = st.text_input("Contraseña de ingreso", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_form == "1234":
                st.session_state["form_logueado"] = True
                st.session_state["nav_state"] = "form"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================================
# FORMULARIO DE NUEVO INGRESO
# ==========================================================
elif st.session_state["nav_state"] == "form":
    if not st.session_state["form_logueado"]:
        st.session_state["nav_state"] = "form_login"
        st.rerun()
        
    if st.button("← Volver al inicio"):
        st.session_state["form_logueado"] = False
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    mostrar_logo(ancho=140)
    st.title("🥛 Registro de Recepción de Leche Cruda")

    if "enviado_exitoso" not in st.session_state:
        st.session_state["enviado_exitoso"] = False

    if st.session_state["enviado_exitoso"]:
        st.success("¡Muchas gracias por tu registro!")
        if st.button("🏠 Volver al inicio", use_container_width=True):
            st.session_state["enviado_exitoso"] = False
            st.session_state["form_logueado"] = False
            st.session_state["nav_state"] = "home"
            st.rerun()
    else:
        with st.form("form_recepcion_leche"):
            st.header("1. Datos de recepción")
            col1, col2 = st.columns(2)
            with col1:
                nombre_responsable = st.selectbox("Responsable de recepción", ["Sandra Garcia", "Daniel Castro", "Luis Perez", "Carlos López", "Marlon Escobar"])
                fecha_recepcion = st.date_input("Fecha de recepción")
            with col2:
                proveedor_opcion = st.selectbox("Proveedor", ["Pasajinak", "Otro"])
                proveedor_final = st.text_input("Especifique el nombre del nuevo proveedor") if proveedor_opcion == "Otro" else proveedor_opcion
                cantidad_leche = st.number_input("Cantidad de leche recibida (litros)", min_value=0.0, format="%.2f")

            st.header("2. Vehículo de Transporte")
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1: limpieza_exterior = st.radio("Limpieza Exterior", ["Bueno", "Malo"], horizontal=True)
            with col_v2: salidas_selladas = st.radio("Salidas de Leche Selladas", ["Bueno", "Malo"], horizontal=True)
            with col_v3: desinfeccion_utensilios = st.radio("Desinfección utensilios", ["Bueno", "Malo"], horizontal=True)

            st.header("3. Análisis Físico-químico")
            temp = st.number_input("Temperatura (°C)", format="%.2f")
            c1, c2, c3 = st.columns(3)
            with c1: color = st.radio("Color", ["Característico", "No característico"], horizontal=True)
            with c2: olor = st.radio("Olor", ["Característico", "No característico"], horizontal=True)
            with c3: sabor = st.radio("Sabor", ["Característico", "No característico"], horizontal=True)
            apariencia = st.radio("Apariencia", ["Sin coágulos", "Con coágulos"], horizontal=True)

            fq1, fq2, fq3 = st.columns(3)
            with fq1:
                ph = st.number_input("pH", format="%.2f")
                grasa = st.number_input("% Grasa", format="%.2f")
                densidad = st.number_input("Densidad", format="%.4f")
                lactosa = st.number_input("Lactosa", format="%.2f")
                antibioticos_res = st.radio("Antibióticos", ["Negativo", "Positivo"], horizontal=True)
            with fq2:
                acido_lactico = st.number_input("% Ácido láctico", format="%.4f")
                sng = st.number_input("% SNG", format="%.2f")
                congelacion = st.number_input("Congelación", format="%.4f")
                conductividad = st.number_input("Conductividad", format="%.2f")
                peroxido = st.radio("Peróxido", ["Negativo", "Positivo"], horizontal=True)
            with fq3:
                st_val = st.number_input("% Sólido Total", format="%.2f")
                proteina = st.number_input("% Proteína", format="%.2f")
                agua_anadida = st.number_input("% Agua Añadida", format="%.2f")

            foto_antibioticos = st.file_uploader("Evidencia fotográfica (Antibióticos)", type=["jpg", "png"])
            
            st.header("4. Resolución y Firma")
            rc1, rc2, rc3 = st.columns(3)
            with rc1: adecuado_proceso = st.radio("Proceso adecuado", ["Si", "No"], horizontal=True)
            with rc2: afecto_ambiente = st.radio("Afectó ambiente", ["Si", "No"], horizontal=True)
            with rc3: resolucion_recepcion = st.radio("Resolución", ["Si", "No"], horizontal=True)

            canvas_result = st_canvas(fill_color="rgba(101, 163, 13, 0.3)", stroke_width=2, stroke_color="#1e3a8a", background_color="#FFFFFF", height=150, width=500, key="canvas_firma")

            submitted = st.form_submit_button("Guardar Registro", type="primary")

        if submitted:
            timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            nombre_firma_guardada = "Sin firma"
            if canvas_result.image_data is not None:
                import numpy as np
                from PIL import Image
                img = Image.fromarray(canvas_result.image_data.astype('uint8'), mode="RGBA")
                nombre_firma_guardada = f"firma_{timestamp_str}.png"
                img.save(os.path.join(FIRMAS_DIR, nombre_firma_guardada))

            nombre_foto_guardada = "Sin imagen"
            if foto_antibioticos is not None:
                nombre_foto_guardada = f"antibioticos_{timestamp_str}.png"
                with open(os.path.join(FOTOS_DIR, nombre_foto_guardada), "wb") as f:
                    f.write(foto_antibioticos.getbuffer())

            registro_datos = {
                "Estado_Aprobacion": "Pendiente",
                "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Responsable": nombre_responsable,
                "Fecha_Recepcion": str(fecha_recepcion),
                "Proveedor": proveedor_final,
                "Cantidad_Litros": cantidad_leche,
                "Limpieza_Exterior": limpieza_exterior,
                "Salidas_Selladas": salidas_selladas,
                "Desinfeccion_Utensilios": desinfeccion_utensilios,
                "Temperatura_C": temp,
                "Color": color, "Olor": olor, "Sabor": sabor, "Apariencia": apariencia,
                "pH": ph, "Acido_Lactico": acido_lactico, "Grasa": grasa, "Solido_No_Graso": sng,
                "Solido_Total": st_val, "Densidad": densidad, "Punto_Congelacion": congelacion,
                "Proteina": proteina, "Lactosa": lactosa, "Conductividad": conductividad,
                "Agua_Anadida": agua_anadida, "Antibioticos_Resultado": antibioticos_res,
                "Peroxido": peroxido, "Evidencia_Foto": nombre_foto_guardada,
                "Carga_Adecuada": adecuado_proceso, "Afecto_Ambiente": afecto_ambiente,
                "Resolucion": resolucion_recepcion, "Firma_Archivo": nombre_firma_guardada
            }
            guardar_en_excel(registro_datos)
            st.session_state["enviado_exitoso"] = True
            st.rerun()

# ==========================================================
# LOGIN ADMINISTRADOR
# ==========================================================
elif st.session_state["nav_state"] == "admin_login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("← Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
        st.title("🔐 Panel de Administrador")
        password_input = st.text_input("Contraseña de administrador", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_input == "glad726lif":
                st.session_state["admin_logueado"] = True
                st.session_state["nav_state"] = "admin_dashboard"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================================
# PANEL DE ADMINISTRADOR
# ==========================================================
elif st.session_state["nav_state"] == "admin_dashboard":
    if not st.session_state["admin_logueado"]:
        st.session_state["nav_state"] = "admin_login"
        st.rerun()
        
    if st.button("← Volver al inicio"):
        st.session_state["nav_state"] = "home"
        st.rerun()
        
    mostrar_logo(ancho=140)
    st.header("📊 Panel de Administrador")

    if os.path.exists(EXCEL_FILE):
        df_registros = pd.read_excel(EXCEL_FILE)
        if "Estado_Aprobacion" not in df_registros.columns:
            df_registros["Estado_Aprobacion"] = "Pendiente"

        tab_tabla, tab_investigacion = st.tabs(["📋 Todos los Registros y Aprobación", "🔍 Investigación y Formato PDF"])

        with tab_tabla:
            st.subheader("✏️ Gestión y Edición General (Sincronizada)")
            df_editado = st.data_editor(df_registros, num_rows="dynamic", key="editor_excel_principal")

            if st.button("💾 Guardar cambios y aprobaciones", type="primary"):
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                    df_editado.to_excel(writer, index=False, sheet_name="Registros")
                    ws = writer.sheets["Registros"]
                    if len(df_editado) > 0:
                        max_row = len(df_editado) + 1
                        max_col = len(df_editado.columns)
                        tab = Table(displayName="TablaRegistrosLeche", ref=f"A1:{get_column_letter(max_col)}{max_row}")
                        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                        ws.add_table(tab)
                st.success("¡Registros actualizados correctamente!")
                st.rerun()

        with tab_investigacion:
            st.subheader("🔍 Investigación y Formato Oficial Aprobado")
            
            # Filtrar solo aprobados o todos según se requiera
            df_aprobados = df_registros[df_registros["Estado_Aprobacion"].str.lower() == "aprobado"]
            
            if len(df_aprobados) == 0:
                st.info("No hay registros aprobados actualmente por el administrador para mostrar el formato oficial.")
            else:
                indices_aprobados = list(df_aprobados.index)
                sel_idx = st.selectbox("Seleccione el registro aprobado a visualizar", indices_aprobados, format_func=lambda x: f"Ingreso #{x} - Proveedor: {df_registros.loc[x, 'Proveedor']} ({df_registros.loc[x, 'Fecha_Recepcion']})")
                
                row = df_registros.loc[sel_idx]
                
                st.markdown("---")
                st.markdown(f"### 📄 Formato Oficial de Recepción - Ingreso #{sel_idx}")
                
                # Visualización tipo tabla formato físico
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    st.markdown(f"**Fecha de recepción:** {row.get('Fecha_Recepcion', '')}")
                    st.markdown(f"**Proveedor:** {row.get('Proveedor', '')}")
                    st.markdown(f"**Cantidad reportada:** {row.get('Cantidad_Litros', '')} L")
                with col_f2:
                    st.markdown(f"**Responsable:** {row.get('Responsable', '')}")
                    st.markdown(f"**Estado:** 🟢 {row.get('Estado_Aprobacion', '')}")
                
                st.markdown("#### Parámetros Evaluados")
                data_tabla_resumen = {
                    "Parámetro / Requisito": ["Limpieza Exterior", "Salidas Selladas", "Temperatura", "pH", "Grasa", "Densidad", "Antibióticos"],
                    "Resultado": [
                        row.get('Limpieza_Exterior', ''), row.get('Salidas_Selladas', ''), 
                        f"{row.get('Temperatura_C', '')} °C", row.get('pH', ''), 
                        f"{row.get('Grasa', '')}%", row.get('Densidad', ''), 
                        row.get('Antibioticos_Resultado', '')
                    ]
                }
                st.table(pd.DataFrame(data_tabla_resumen))
                
                # Botón de Descarga PDF
                if st.button("📥 Descargar Reporte en PDF (Formato Oficial Horizontal)", type="primary"):
                    pdf_path = generar_pdf_formato(row, sel_idx)
                    with open(pdf_path, "rb") as pdf_file:
                        st.download_button(
                            label="Descargar archivo PDF ahora",
                            data=pdf_file,
                            file_name=f"Recepcion_Leche_Ingreso_{sel_idx}.pdf",
                            mime="application/pdf"
                        )
    else:
        st.info("Aún no hay registros guardados en el sistema.")
