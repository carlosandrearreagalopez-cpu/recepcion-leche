import streamlit as st
import pandas as pd
import os
from datetime import datetime
from streamlit_drawable_canvas import st_canvas

# Configuración de la página
st.set_page_config(page_title="Recepción de Leche Cruda - LIF Brands", layout="wide")

# ==========================================
# ESTILOS CSS CON IDENTIDAD VISUAL LIF BRANDS
# ==========================================
st.markdown("""
    <style>
        /* Fondo general de la aplicación totalmente blanco */
        .stApp {
            background-color: #FFFFFF !important;
        }
        
        /* Tipografía Arial para todos los textos */
        html, body, [class*="css"], p, span, label {
            font-family: Arial, sans-serif !important;
            color: #000000 !important;
        }

        /* Títulos principales con el Azul Oscuro de LIF Brands */
        h1, h2, h3, h4, h5, h6 {
            color: #1e3a8a !important;
            font-family: Arial, sans-serif !important;
        }

        /* Etiquetas de los formularios visibles y formales */
        .stTextInput label, .stSelectbox label, .stDateInput label, .stNumberInput label, .stRadio label, .stFileUploader label {
            color: #1e3a8a !important;
            font-weight: bold !important;
        }

        /* Limpiar y ocultar el texto duplicado del botón de carga manteniendo el diseño limpio */
        [data-testid="stFileUploader"] section div button {
            color: transparent !important;
        }
        [data-testid="stFileUploader"] section div button::after {
            content: "Examinar";
            color: #ffffff !important;
            position: absolute;
            left: 0; right: 0; top: 0; bottom: 0;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        /* Caja de carga de archivos (file_uploader) */
        [data-testid="stFileUploader"] {
            background-color: #f8fafc !important;
            border: 2px dashed #cbd5e1 !important;
            border-radius: 8px;
            padding: 10px;
        }
        [data-testid="stFileUploader"] section {
            background-color: #f8fafc !important;
        }
        [data-testid="stFileUploader"] button {
            background-color: #1e3a8a !important;
            border: none !important;
            border-radius: 4px !important;
        }
        [data-testid="stFileUploader"] button:hover {
            background-color: #3b82f6 !important;
        }
        [data-testid="stFileUploader"] small, [data-testid="stFileUploader"] span, [data-testid="stFileUploader"] div {
            color: #334155 !important;
        }

        /* Botones generales (Fondo blanco, borde y texto en Azul Corporativo) */
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

        /* Botón de envío principal con el Azul Corporativo lleno */
        button[kind="primary"] {
            background-color: #1e3a8a !important;
            color: #FFFFFF !important;
            border: none !important;
        }
        button[kind="primary"]:hover {
            background-color: #3b82f6 !important;
        }

        /* Campos de entrada (inputs) claros y legibles */
        input, select {
            background-color: #f8fafc !important;
            color: #000000 !important;
            border: 1px solid #cbd5e1 !important;
            border-radius: 4px;
        }

        /* Tarjetas de investigación con acento en Azul Claro */
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
if not os.path.exists(FOTOS_DIR):
    os.makedirs(FOTOS_DIR)
if not os.path.exists(FIRMAS_DIR):
    os.makedirs(FIRMAS_DIR)

EXCEL_FILE = "registros_recepcion_leche.xlsx"

# Función segura para mostrar el logo local subido a GitHub
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
    df_final.to_excel(EXCEL_FILE, index=False)

# Control de navegación principal
if "nav_state" not in st.session_state:
    st.session_state["nav_state"] = "home"

if "admin_logueado" not in st.session_state:
    st.session_state["admin_logueado"] = False

if "form_logueado" not in st.session_state:
    st.session_state["form_logueado"] = False

# ==========================================
# PANTALLA DE INICIO (ESTILO PORTADA)
# ==========================================
if st.session_state["nav_state"] == "home":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        mostrar_logo(ancho=200)
        
        st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>Ingresos a bodega</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #65a30d; font-weight: bold; font-size: 16px;'>LIF Brands — Aseguramiento de Calidad</p>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #475569;'>Documentación de recepción de materiales</p>", unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("Reportar nuevo ingreso", use_container_width=True, type="primary"):
            st.session_state["nav_state"] = "form_login"
            st.rerun()
            
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        
        col_link1, col_link2, col_link3 = st.columns([1, 2, 1])
        with col_link2:
            if st.button("Revisar ingresos (administrador)"):
                st.session_state["nav_state"] = "admin_login"
                st.rerun()

# ==========================================
# LOGIN PARA NUEVO INGRESO
# ==========================================
elif st.session_state["nav_state"] == "form_login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("← Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
            
        st.title("🔒 Acceso a Registro")
        st.markdown("Ingrese la contraseña autorizada para reportar un nuevo ingreso:")
        
        password_form = st.text_input("Contraseña de ingreso", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_form == "1234":
                st.session_state["form_logueado"] = True
                st.session_state["nav_state"] = "form"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================
# FORMULARIO DE NUEVO INGRESO
# ==========================================
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
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("🔄 Ingresar una nueva respuesta", use_container_width=True):
                st.session_state["enviado_exitoso"] = False
                st.rerun()
        with col_b2:
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
                nombre_responsable = st.selectbox(
                    "Responsable de recepción",
                    ["Sandra Garcia", "Daniel Castro", "Luis Perez", "Carlos López", "Marlon Escobar"]
                )
                fecha_recepcion = st.date_input("Fecha de recepción")
            
            with col2:
                proveedor_opcion = st.selectbox("Proveedor", ["Pasajinak", "Otro"])
                if proveedor_opcion == "Otro":
                    proveedor_final = st.text_input("Especifique el nombre del nuevo proveedor")
                else:
                    proveedor_final = proveedor_opcion
                    
                cantidad_leche = st.number_input("Cantidad de leche recibida (litros)", min_value=0.0, format="%.2f")

            st.header("2. Vehículo de Transporte")
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                limpieza_exterior = st.radio("Limpieza Exterior", ["Bueno", "Malo"], horizontal=True)
            with col_v2:
                salidas_selladas = st.radio("Salidas de Leche Selladas", ["Bueno", "Malo"], horizontal=True)
            with col_v3:
                desinfeccion_utensilios = st.radio("Desinfección utensilios para tomar muestra", ["Bueno", "Malo"], horizontal=True)

            st.header("3. Análisis Físico-químico")
            st.subheader("Temperatura (°C)")
            st.caption("Parámetro esperado: Menor o igual 7 °C.")
            temp = st.number_input("Resultado Temperatura (°C)", format="%.2f", key="temp")
            
            c_col1, c_col2, c_col3 = st.columns(3)
            with c_col1:
                st.markdown("**Color**")
                color = st.radio("Color res", ["Característico", "No característico"], horizontal=True, label_visibility="collapsed", key="color")
            with c_col2:
                st.markdown("**Olor**")
                olor = st.radio("Olor res", ["Característico", "No característico"], horizontal=True, label_visibility="collapsed", key="olor")
            with c_col3:
                st.markdown("**Sabor**")
                sabor = st.radio("Sabor res", ["Característico", "No característico"], horizontal=True, label_visibility="collapsed", key="sabor")
            
            st.subheader("Apariencia")
            apariencia = st.radio("Resultado Apariencia", ["Sin presencia de sedimentación y coágulos", "Con presencia de sedimentación y coágulos"], horizontal=True, key="apariencia")
            
            fq_c1, fq_c2, fq_c3 = st.columns(3)
            with fq_c1:
                st.markdown("**pH** (Esperado: 6.5 - 6.8)")
                ph = st.number_input("pH res", format="%.2f", label_visibility="collapsed", key="ph")
                st.markdown("**% Grasa** (Esperado: Mínimo 3%)")
                grasa = st.number_input("Grasa res", format="%.2f", label_visibility="collapsed", key="grasa")
                st.markdown("**Densidad** (Esperado: 1.030 - 1.033 g/ml)")
                densidad = st.number_input("Densidad res", format="%.4f", label_visibility="collapsed", key="densidad")
                st.markdown("**Lactosa** (Esperado: Mínimo 4.2%)")
                lactosa = st.number_input("Lactosa res", format="%.2f", label_visibility="collapsed", key="lactosa")
                st.markdown("**Antibióticos (Resultado)**")
                antibioticos_res = st.radio("Antibióticos res", ["Negativo", "Positivo"], horizontal=True, label_visibility="collapsed", key="antibioticos_res")

            with fq_c2:
                st.markdown("**% Ácido láctico** (Esperado: 0.13 - 0.17)")
                acido_lactico = st.number_input("Ácido res", format="%.4f", label_visibility="collapsed", key="acido")
                st.markdown("**% Sólido No Graso** (Esperado: Mínimo 8.3%)")
                sng = st.number_input("SNG res", format="%.2f", label_visibility="collapsed", key="sng")
                st.markdown("**Punto de Congelación** (-0.51 a -0.55 °C)")
                congelacion = st.number_input("Congelación res", format="%.4f", label_visibility="collapsed", key="congelacion")
                st.markdown("**Conductividad** (Máximo 6 mS/cm)")
                conductividad = st.number_input("Conductividad res", format="%.2f", label_visibility="collapsed", key="conductividad")
                st.markdown("**Peróxido**")
                peroxido = st.radio("Peróxido res", ["Positivo", "Negativo"], horizontal=True, label_visibility="collapsed", key="peroxido")

            with fq_c3:
                st.markdown("**% Sólido Total** (Esperado: Mínimo 11.3%)")
                st_val = st.number_input("ST res", format="%.2f", label_visibility="collapsed", key="st")
                st.markdown("**% Proteína** (Esperado: Mínimo 3%)")
                proteina = st.number_input("Proteína res", format="%.2f", label_visibility="collapsed", key="proteina")
                st.markdown("**Agua Añadida (%)** (Esperado: 0%)")
                agua_anadida = st.number_input("Agua res", format="%.2f", label_visibility="collapsed", key="agua")

            st.subheader("Evidencia fotográfica (Antibióticos)")
            st.caption("Cargue una imagen o tome una foto de la prueba.")
            foto_antibioticos = st.file_uploader("Subir imagen / Tomar foto", type=["jpg", "jpeg", "png"], key="foto_antibioticos")

            st.header("4. Carga/Descarga y Resolución")
            st.markdown("**Evaluación final del proceso y decisión de calidad.**")
            
            rc1, rc2, rc3 = st.columns(3)
            with rc1:
                adecuado_proceso = st.radio("¿El proceso de carga/descarga fue el adecuado?", ["Si", "No"], horizontal=True)
            with rc2:
                afecto_ambiente = st.radio("¿Afectó de forma potencial al medio ambiente?", ["Si", "No"], horizontal=True)
            with rc3:
                resolucion_recepcion = st.radio("Resolución de Recepción", ["Si", "No"], horizontal=True)
            
            submitted = st.form_submit_button("Guardar Registro de Recepción", type="primary")

        st.subheader("Firma del Responsable (Dibuje su firma en el recuadro)")
        canvas_result = st_canvas(
            fill_color="rgba(101, 163, 13, 0.3)",
            stroke_width=2,
            stroke_color="#1e3a8a",
            background_color="#FFFFFF",
            height=150,
            width=500,
            drawing_mode="freedraw",
            key="canvas_firma",
        )

        if submitted:
            if proveedor_opcion == "Otro" and not proveedor_final.strip():
                st.error("Por favor, ingrese el nombre del nuevo proveedor.")
            else:
                timestamp_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                
                nombre_firma_guardada = "Sin firma"
                if canvas_result.image_data is not None:
                    import numpy as np
                    from PIL import Image
                    img_data = canvas_result.image_data
                    img = Image.fromarray(img_data.astype('uint8'), mode="RGBA")
                    nombre_firma_guardada = f"firma_{timestamp_str}.png"
                    ruta_firma = os.path.join(FIRMAS_DIR, nombre_firma_guardada)
                    img.save(ruta_firma)

                nombre_foto_guardada = "Sin imagen"
                if foto_antibioticos is not None:
                    nombre_foto_guardada = f"antibioticos_{timestamp_str}.png"
                    ruta_foto = os.path.join(FOTOS_DIR, nombre_foto_guardada)
                    with open(ruta_foto, "wb") as f:
                        f.write(foto_antibioticos.getbuffer())

                registro_datos = {
                    "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Responsable": nombre_responsable,
                    "Fecha_Recepcion": str(fecha_recepcion),
                    "Proveedor": proveedor_final,
                    "Cantidad_Litros": cantidad_leche,
                    "Limpieza_Exterior": limpieza_exterior,
                    "Salidas_Selladas": salidas_selladas,
                    "Desinfeccion_Utensilios": desinfeccion_utensilios,
                    "Temperatura_C": temp,
                    "Color": color,
                    "Olor": olor,
                    "Sabor": sabor,
                    "Apariencia": apariencia,
                    "pH": ph,
                    "Acido_Lactico": acido_lactico,
                    "Grasa": grasa,
                    "Solido_No_Graso": sng,
                    "Solido_Total": st_val,
                    "Densidad": densidad,
                    "Punto_Congelacion": congelacion,
                    "Proteina": proteina,
                    "Lactosa": lactosa,
                    "Conductividad": conductividad,
                    "Agua_Anadida": agua_anadida,
                    "Antibioticos_Resultado": antibioticos_res,
                    "Peroxido": peroxido,
                    "Evidencia_Foto": nombre_foto_guardada,
                    "Carga_Adecuada": adecuado_proceso,
                    "Afecto_Ambiente": afecto_ambiente,
                    "Resolucion": resolucion_recepcion,
                    "Firma_Archivo": nombre_firma_guardada
                }
                
                guardar_en_excel(registro_datos)
                st.session_state["enviado_exitoso"] = True
                st.rerun()

# ==========================================
# LOGIN DE ADMINISTRADOR
# ==========================================
elif st.session_state["nav_state"] == "admin_login":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("← Volver al inicio"):
            st.session_state["nav_state"] = "home"
            st.rerun()
            
        st.title("🔒 Panel de Administrador")
        st.markdown("Ingrese la contraseña de administrador para acceder a los registros y reportes.")
        
        password_input = st.text_input("Contraseña de administrador", type="password")
        if st.button("Verificar Acceso", use_container_width=True, type="primary"):
            if password_input == "glad726lif":
                st.session_state["admin_logueado"] = True
                st.session_state["nav_state"] = "admin_dashboard"
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")

# ==========================================
# VISTA DE ADMINISTRADOR (DASHBOARD)
# ==========================================
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
        
        tab_tabla, tab_investigacion = st.tabs(["📋 Todos los Registros (Excel y Edición)", "🔍 Investigación"])
        
        # --- PESTAÑA 1: TABLA GENERAL Y EDICIÓN ---
        with tab_tabla:
            df_registros["Fecha_Recepcion_dt"] = pd.to_datetime(df_registros["Fecha_Recepcion"], errors="coerce")
            df_registros["Semana"] = df_registros["Fecha_Recepcion_dt"].dt.to_period("W").astype(str)
            
            st.subheader("📈 Resumen de Ingresos Semanales (Litros)")
            if "Cantidad_Litros" in df_registros.columns and "Semana" in df_registros.columns:
                df_semanal = df_registros.groupby("Semana")["Cantidad_Litros"].sum().reset_index()
                df_semanal.columns = ["Semana", "Total Litros Recibidos"]
                st.dataframe(df_semanal, use_container_width=True)
                st.bar_chart(df_semanal.set_index("Semana"))
            
            st.write("---")
            st.subheader("✏️ Corregir o Editar Registros")
            df_editado = st.data_editor(df_registros, num_rows="dynamic", key="editor_excel")
            
            if st.button("💾 Guardar correcciones en el Excel", type="primary"):
                if "Fecha_Recepcion_dt" in df_editado.columns:
                    df_editado = df_editado.drop(columns=["Fecha_Recepcion_dt"])
                if "Semana" in df_editado.columns:
                    df_editado = df_editado.drop(columns=["Semana"])
                df_editado.to_excel(EXCEL_FILE, index=False)
                st.success("¡Correcciones guardadas exitosamente!")
                st.rerun()

            st.write("---")
            st.subheader("🗑️ Eliminar Registro de Prueba")
            indices_disponibles = list(df_registros.index)
            if indices_disponibles:
                fila_a_eliminar = st.selectbox("Seleccione el número de fila del registro a eliminar", indices_disponibles)
                if st.button("🗑️ Eliminar este registro definitivamente"):
                    df_registros = df_registros.drop(fila_a_eliminar).reset_index(drop=True)
                    if "Fecha_Recepcion_dt" in df_registros.columns:
                        df_registros = df_registros.drop(columns=["Fecha_Recepcion_dt"])
                    if "Semana" in df_registros.columns:
                        df_registros = df_registros.drop(columns=["Semana"])
                    df_registros.to_excel(EXCEL_FILE, index=False)
                    st.success(f"¡El registro de la fila {fila_a_eliminar} fue eliminado correctamente!")
                    st.rerun()
            
            st.write("---")
            with open(EXCEL_FILE, "rb") as f:
                st.download_button(
                    label="📥 Descargar Excel completo",
                    data=f,
                    file_name="registros_recepcion_leche.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # --- PESTAÑA 2: INVESTIGACIÓN DETALLADA ---
        with tab_investigacion:
            st.subheader("🔍 Consulta e Investigación por Registro")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                filtro_fecha = st.date_input("Filtrar por fecha de recepción (Opcional)", value=None)
            with col_f2:
                proveedores_disponibles = ["Todos"] + list(df_registros["Proveedor"].dropna().unique())
                filtro_proveedor = st.selectbox("Filtrar por Proveedor", proveedores_disponibles)
                
            df_filtrado = df_registros.copy()
            if filtro_fecha:
                df_filtrado = df_filtrado[df_filtrado["Fecha_Recepcion"].astype(str) == str(filtro_fecha)]
            if filtro_proveedor != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Proveedor"] == filtro_proveedor]
                
            st.markdown(f"**Mostrando {len(df_filtrado)} registros encontrados:**")
            st.write("---")
            
            for idx, row in df_filtrado.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="card-investigacion">
                        <b style="color: #1e3a8a;">⏳ Ingreso #{idx} — Proveedor: {row.get('Proveedor', 'N/A')}</b> <br>
                        <span style="color: #475569;">Fecha: {row.get('Fecha_Recepcion', 'N/A')} | Responsable: {row.get('Responsable', 'N/A')} | Litros: {row.get('Cantidad_Litros', 0)} L</span>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    if st.button(f"Ver detalle completo #{idx}", key=f"btn_ver_{idx}"):
                        st.markdown(f"### 📋 Detalle del Registro #{idx}")
                        
                        col_d1, col_d2 = st.columns(2)
                        with col_d1:
                            st.markdown(f"**Fecha y Hora del Registro:** {row.get('Fecha_Hora', '')}")
                            st.markdown(f"**Responsable:** {row.get('Responsable', '')}")
                            st.markdown(f"**Fecha de Recepción:** {row.get('Fecha_Recepcion', '')}")
                            st.markdown(f"**Proveedor:** {row.get('Proveedor', '')}")
                            st.markdown(f"**Cantidad Recibida:** {row.get('Cantidad_Litros', '')} Litros")
                            st.markdown(f"**Limpieza Exterior:** {row.get('Limpieza_Exterior', '')}")
                            st.markdown(f"**Salidas Selladas:** {row.get('Salidas_Selladas', '')}")
                            st.markdown(f"**Desinfección Utensilios:** {row.get('Desinfeccion_Utensilios', '')}")
                            st.markdown(f"**Temperatura:** {row.get('Temperatura_C', '')} °C")
                            st.markdown(f"**Color:** {row.get('Color', '')}")
                            st.markdown(f"**Olor:** {row.get('Olor', '')}")
                            st.markdown(f"**Sabor:** {row.get('Sabor', '')}")
                            st.markdown(f"**Apariencia:** {row.get('Apariencia', '')}")
                        
                        with col_d2:
                            st.markdown(f"**pH:** {row.get('pH', '')}")
                            st.markdown(f"**Ácido Láctico:** {row.get('Acido_Lactico', '')}")
                            st.markdown(f"**Grasa:** {row.get('Grasa', '')}%")
                            st.markdown(f"**Sólido No Graso:** {row.get('Solido_No_Graso', '')}%")
                            st.markdown(f"**Sólido Total:** {row.get('Solido_Total', '')}%")
                            st.markdown(f"**Densidad:** {row.get('Densidad', '')}")
                            st.markdown(f"**Punto de Congelación:** {row.get('Punto_Congelacion', '')}")
                            st.markdown(f"**Proteína:** {row.get('Proteina', '')}%")
                            st.markdown(f"**Lactosa:** {row.get('Lactosa', '')}%")
                            st.markdown(f"**Conductividad:** {row.get('Conductividad', '')}")
                            st.markdown(f"**Agua Añadida:** {row.get('Agua_Anadida', '')}%")
                            st.markdown(f"**Antibióticos:** {row.get('Antibióticos_Resultado', '')}")
                            st.markdown(f"**Peróxido:** {row.get('Peróxido', '')}")
                        
                        st.markdown("---")
                        st.markdown("#### 🖼️ Evidencias y Firmas Registradas")
                        img_col1, img_col2 = st.columns(2)
                        
                        with img_col1:
                            st.markdown("**Evidencia Fotográfica (Antibióticos):**")
                            nombre_foto = row.get('Evidencia_Foto', 'Sin imagen')
                            ruta_foto_servidor = os.path.join(FOTOS_DIR, str(nombre_foto))
                            if nombre_foto != "Sin imagen" and os.path.exists(ruta_foto_servidor):
                                st.image(ruta_foto_servidor, caption="Prueba de Antibióticos", width=300)
                            else:
                                st.info("No hay imagen adjunta o no se encuentra en el servidor.")
                                
                        with img_col2:
                            st.markdown("**Firma del Responsable:**")
                            nombre_firma = row.get('Firma_Archivo', 'Sin firma')
                            ruta_firma_servidor = os.path.join(FIRMAS_DIR, str(nombre_firma))
                            if nombre_firma != "Sin firma" and os.path.exists(ruta_firma_servidor):
                                st.image(ruta_firma_servidor, caption="Firma digital", width=300)
                            else:
                                st.info("No hay firma registrada.")
                        st.write("---")

    else:
        st.info("Aún no hay registros guardados en el sistema.")
