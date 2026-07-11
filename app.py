import streamlit as st
from datetime import datetime
import os
import json
import time
from docx import Document
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ==========================================
# CONFIGURACIÓN CRÍTICA
# ==========================================
ID_CARPETA_SANTIAGO = "TU_ID_DE_CARPETA_DE_GOOGLE_DRIVE_AQUÍ"

st.set_page_config(
    page_title="Portal de Convenios | Acuario Michin CDMX",
    page_icon="🌊",
    layout="centered"
)

# Estilos CSS
st.markdown("""
    <style>
    [data-testid="stAppApp"] {
        background-color: #001222 !important;
        background-image: radial-gradient(circle at 10% 20%, #062039 0%, transparent 40%), radial-gradient(circle at 80% 50%, #002e5a 0%, transparent 30%) !important;
    }
    .main-title { color: #00b9cc !important; font-size: 36px; font-weight: 800; text-align: center; margin-top: 10px; margin-bottom: 0px; }
    .subtitle { color: #ffffff; text-align: center; font-size: 24px; margin-top: 5px; opacity: 0.9; }
    [data-testid="stForm"] { background-color: #062039 !important; border: 1px solid #104169; border-radius: 12px; padding: 30px; }
    h4 { color: #00b9cc !important; font-family: 'Merriweather', serif; font-weight: 700; margin-top: 25px; margin-bottom: 15px; border-bottom: 2px solid #005B94; padding-bottom: 5px; }
    label { color: #ffffff !important; font-weight: 600 !important; }
    .stTextInput input, .stTextArea textarea { background-color: #1a324d !important; color: #ffffff !important; border: 1px solid #20507a !important; }
    .stButton>button { background-color: #00b9cc !important; color: #0B2545 !important; border-radius: 6px !important; width: 100%; font-weight: 700 !important; text-transform: uppercase; }
    </style>
""", unsafe_allow_html=True)

# --- ENCABEZADO: LOGOS Y TÍTULOS ---
col_izq, col_centro, col_der = st.columns([1, 3, 1])
with col_izq:
    st.image("logo1.png", use_container_width=True)
with col_der:
    st.image("logo2.png", use_container_width=True)

st.markdown("<div class='main-title'>ACUARIO MICHIN CDMX</div>", unsafe_allow_html=True)
st.markdown("""
    <div style='text-align: center; margin-top: 5px;'>
        <span style='background-color: #ffffff; color: #000000; padding: 5px 20px; border-radius: 5px; font-weight: 800; font-size: 24px; text-transform: uppercase;'>
            CONVENIOS
        </span>
    </div>
""", unsafe_allow_html=True)

# 4. Funciones Backend
def guardar_en_sheets(datos):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if "gcp" in st.secrets:
            credenciales_dict = json.loads(st.secrets["gcp"]["service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(credenciales_dict, scope)
        else:
            creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
            
        cliente = gspread.authorize(creds)
        hoja = cliente.open("CRM Convenios Michin").sheet1
        fila = [
            datos["fecha_registro"], datos["tipo_entidad"], datos["razon_social"], 
            datos["nombre_comercial"], datos["rfc"], datos["giro"], 
            datos["colaboradores"], datos["nombre_firmante"], datos["cargo_firmante"], 
            datos["nombre_contacto"], datos["correo_contacto"], datos["tel_contacto"], 
            datos["domicilio_fiscal"], datos["domicilio_empresa"]
        ]
        hoja.append_row(fila)
        return True
    except Exception as e:
        st.error(f"⚠️ Error al conectar con Google Sheets: {e}")
        return False

# Funciones de Word y Drive (omitidas para brevedad, mantén tus versiones actuales)
# ... (asegúrate de que las funciones de generar y subir no llamen a la variable 'ejecutivo')

# 5. Formulario (Sin ejecutivos)
with st.form("form_convenio", clear_on_submit=False):
    st.markdown("<h4>1. Clasificación</h4>", unsafe_allow_html=True)
    tipo_entidad = st.selectbox("Clasificación Jurídica:", ["Empresa Privada", "Sindicato o Secretaría Pública"])
        
    st.markdown("<h4>2. Datos de la Empresa</h4>", unsafe_allow_html=True)
    col3, col4 = st.columns(2)
    with col3:
        razon_social = st.text_input("Razón Social Oficial:")
        rfc = st.text_input("RFC:", max_chars=13)
        giro = st.text_input("Giro:")
    with col4:
        nombre_comercial = st.text_input("Nombre Comercial:")
        colaboradores = st.text_input("Cantidad de Colaboradores:", value="50")
        
    st.markdown("<h4>3. Domicilios</h4>", unsafe_allow_html=True)
    domicilio_fiscal = st.text_area("Domicilio Fiscal:")
    domicilio_empresa = st.text_area("Domicilio Operativo:")
    
    st.markdown("<h4>4. Contacto</h4>", unsafe_allow_html=True)
    col5, col6 = st.columns(2)
    with col5:
        nombre_firmante = st.text_input("Responsable de Firma Ej. RH, Beneficios, PR")
        nombre_contacto = st.text_input("Contacto RH:")
        st.markdown("<h4>5. Oportunidades Comerciales</h4>", unsafe_allow_html=True)
eventos_anuales = st.text_area("Eventos corporativos que realizan al año (Ej. Día de la Familia, Posadas, Aniversarios):", placeholder="Cuéntanos qué celebran y en qué meses...")
with col6:
        cargo_firmante = st.text_input("Cargo del Firmante:")
        tel_contacto = st.text_input("Teléfono:", max_chars=10)
        correo_contacto = st.text_input("Correo Institucional:")
        
        submit_button = st.form_submit_button(label="Iniciar Circuito de Convenio")

if submit_button:
    datos_empresa = {
        "fecha_registro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "tipo_entidad": tipo_entidad,
        "razon_social": razon_social,
        "nombre_comercial": nombre_comercial,
        "rfc": rfc,
        "giro": giro,
        "colaboradores": colaboradores,
        "nombre_firmante": nombre_firmante,
        "eventos_anuales": eventos_anuales,
        "cargo_firmante": cargo_firmante,
        "nombre_contacto": nombre_contacto,
        "correo_contacto": correo_contacto,
        "tel_contacto": tel_contacto,
        "domicilio_fiscal": domicilio_fiscal,
        "domicilio_empresa": domicilio_empresa if domicilio_empresa else domicilio_fiscal
    }
    
    with st.spinner("Registrando..."):
        if guardar_en_sheets(datos_empresa):
            st.success("¡Registro completado!")