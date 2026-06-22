import streamlit as st
from datetime import datetime
import os
from docx import Document
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# Importamos las herramientas oficiales para conectar con Google Drive
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import time

# ==========================================
# CONFIGURACIÓN CRÍTICA (ALONSO: PEGA AQUÍ TU ID DE CARPETA)
# ==========================================
# Reemplaza este texto por la ID de la carpeta de Google Drive que creaste para Santiago:
ID_CARPETA_SANTIAGO = "TU_ID_DE_CARPETA_DE_GOOGLE_DRIVE_AQUÍ"


# 1. Configuración de la página
st.set_page_config(
    page_title="Portal de Gestión  | Acuario Michin CDMX",
    page_icon="🌊",
    layout="centered"
)

# 2. Estilos CSS (Océano Profundo - Espectacular e Intacto)
st.markdown("""
    <style>
    [data-testid="stAppApp"] {
        background-color: #001222 !important;
        background-image: radial-gradient(circle at 10% 20%, #062039 0%, transparent 40%), radial-gradient(circle at 80% 50%, #002e5a 0%, transparent 30%) !important;
    }
    [data-testid="stHeader"] { background-color: transparent !important; }
    .stMarkdown, p, li { color: #ffffff !important; font-family: 'Montserrat', 'Helvetica Neue', sans-serif !important; }
    .main-title { color: #00b9cc !important; font-size: 32px; font-weight: 800; text-align: center; margin-bottom: 5px; text-shadow: 0 0 15px rgba(0, 185, 204, 0.4); }
    .subtitle { color: #ffffff; text-align: center; opacity: 0.8; font-size: 16px; margin-bottom: 40px; font-weight: 300; }
    [data-testid="stForm"] { background-color: #062039 !important; border: 1px solid #104169; border-radius: 12px; padding: 30px; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.4); }
    h4 { color: #00b9cc !important; font-family: 'Merriweather', serif; font-weight: 700; margin-top: 25px; margin-bottom: 15px; border-bottom: 2px solid #005B94; padding-bottom: 5px; }
    label, [data-testid="stWidgetLabel"] p { color: #ffffff !important; font-weight: 600 !important; font-size: 15px !important; }
    .stTextInput input, .stTextArea textarea { color: #ffffff !important; background-color: #1a324d !important; border: 1px solid #20507a !important; border-radius: 6px; }
    .stTextInput input:focus { border-color: #00b9cc !important; box-shadow: 0 0 8px rgba(0, 185, 204, 0.3); }
    div[data-baseweb="select"] { color: #ffffff !important; background-color: #1a324d !important; border: 1px solid #20507a !important; border-radius: 6px; }
    div[data-baseweb="popover"] ul { background-color: #1a324d !important; }
    div[data-baseweb="popover"] li { color: #ffffff !important; }
    div[data-baseweb="popover"] li:hover { background-color: #005B94 !important; }
    .stButton>button { background-color: #00b9cc !important; color: #0B2545 !important; border-radius: 6px !important; padding: 12px 24px !important; font-weight: 700 !important; font-size: 18px; border: none !important; width: 100%; transition: all 0.3s ease; text-transform: uppercase; letter-spacing: 1px; }
    .stButton>button:hover { background-color: #ffffff !important; box-shadow: 0 4px 15px rgba(0, 185, 204, 0.5); }
    </style>
""", unsafe_allow_html=True)

# 3. Encabezado Institucional (Con los nombres exactos de tus archivos)
col_espacio1, col_logos, col_espacio2 = st.columns([1, 1.5, 1])
with col_logos:
    try:
        st.image("logo1.png", use_container_width=True)
        st.image("logo2.png", use_container_width=True)
    except Exception:
        st.warning("⚠️ Asegúrate de que las imágenes conserven sus nombres originales en tu VS Code.")

st.markdown("<div class='main-title'>Acuario Michin CDMX</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Portal de Gestión B2B &bull; Ecosistema Corporativo</div>", unsafe_allow_html=True)

# 4. Funciones Backend Centralizadas
def guardar_en_sheets(datos):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
        cliente = gspread.authorize(creds)
        hoja = cliente.open("CRM Convenios Michin").sheet1
        fila = [
            datos["fecha_registro"], datos["ejecutivo_comercial"], datos["tipo_entidad"], 
            datos["razon_social"], datos["nombre_comercial"], datos["rfc"], 
            datos["giro"], datos["colaboradores"], datos["nombre_firmante"], 
            datos["cargo_firmante"], datos["nombre_contacto"], datos["correo_contacto"], 
            datos["tel_contacto"], datos["domicilio_fiscal"], datos["domicilio_empresa"]
        ]
        hoja.append_row(fila)
        return True
    except Exception as e:
        st.error(f"⚠️ Error al conectar con Google Sheets: {e}")
        return False

def generar_documento_convenio(datos):
    try:
        doc = Document("Machote_convenio.docx")
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        fecha_actual = datetime.now()
        fecha_formateada = f"{fecha_actual.day} de {meses[fecha_actual.month - 1]} de {fecha_actual.year}"
        
        reemplazos = {
            "{{NOMBRE_EMPRESA}}": datos["razon_social"], 
            "{{NOMBRE_COMERCIAL}}": datos["nombre_comercial"], 
            "{{FIRMANTE}}": datos["nombre_firmante"], 
            "{{CARGO_FIRMANTE}}": datos["cargo_firmante"], 
            "{{DOMICILIO_FISCAL}}": datos["domicilio_fiscal"], 
            "{{RFC}}": datos["rfc"], 
            "{{FECHA_CONVENIO}}": fecha_formateada
        }
        
        for p in doc.paragraphs:
            for clave, valor in reemplazos.items():
                if clave in p.text: p.text = p.text.replace(clave, valor)
                
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    for p in celda.paragraphs:
                        for clave, valor in reemplazos.items():
                            if clave in p.text: p.text = p.text.replace(clave, valor)
                            
        ruta_salida = f"Convenio_{datos['nombre_comercial'].replace(' ', '_')}_BORRADOR.docx"
        doc.save(ruta_salida)
        return ruta_salida
    except Exception as e:
        st.error(f"⚠️ Error al procesar el Word base: {e}")
        return None

def subir_a_google_drive(ruta_archivo, nombre_comercial):
    """Automatización exclusiva para enviar el archivo directamente al control de Santiago Segoviano"""
    scope = ["https://www.googleapis.com/auth/drive"]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
        servicio = build('drive', 'v3', credentials=creds)
        
        metadatos_archivo = {
            'name': f"Convenio_{nombre_comercial.replace(' ', '_')}_SUPERVISIÓN_SANTIAGO.docx",
            'parents': [ID_CARPETA_SANTIAGO]
        }
        media = MediaFileUpload(ruta_archivo, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        
        archivo_drive = servicio.files().create(body=metadatos_archivo, media_body=media, fields='id').execute()
        return archivo_drive.get('id')
    except Exception as e:
        st.error(f"⚠️ Alerta: El convenio se guardó en el CRM pero no pudo subirse a la nube de Santiago: {e}")
        return None

# 5. Formulario de Captura Completo (15 datos protegidos para Clientes)
with st.form("form_convenio", clear_on_submit=False):
    st.markdown("<h4>1. Información Comercial de Enlace</h4>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        # Se adaptó para que el cliente seleccione quién lo está atendiendo
        ejecutivo = st.selectbox("¿Quién es tu Ejecutivo Comercial asignado en Acuario Michin?", ["Alonso Garcia", "Regina Cedeño", "Hugo Sandoval", "Raul Sanchez", "Oscar Flores", "Ana Lau"])
    with col2:
        tipo_entidad = st.selectbox("Clasificación Jurídica de su Entidad:", ["Empresa Privada", "Sindicato o Secretaría Pública"])
        
    st.markdown("<h4>2. Datos Fiscales, Legales y Comerciales de su Empresa</h4>", unsafe_allow_html=True)
    col3, col4 = st.columns(2)
    with col3:
        razon_social = st.text_input("Razón Social Oficial (Como aparece en el SAT):", placeholder="Ej. Comercializadora S.A. de C.V.")
        rfc = st.text_input("RFC de la Empresa:", max_chars=13, placeholder="XAXX010101000")
        giro = st.text_input("Giro o Actividad Principal de la Empresa:")
    with col4:
        nombre_comercial = st.text_input("Nombre Comercial (Cómo se conocerá la alianza):", placeholder="Ej. Tiendas Aqua")
        colaboradores = st.text_input("Cantidad Total de Colaboradores potenciales:", value="50")
        web_social = st.text_input("Página Web o Red Social Corporativa:", placeholder="www.empresa.com")
        
    st.markdown("<h4>3. Ubicaciones Oficiales</h4>", unsafe_allow_html=True)
    domicilio_fiscal = st.text_area("Domicilio Fiscal Completo:", height=60, placeholder="Calle, Número, Colonia, CP, Estado.")
    domicilio_empresa = st.text_area("Domicilio Operativo Principal (Si es diferente al fiscal):", height=60, placeholder="Dejar en blanco si es el mismo.")
    
    st.markdown("<h4>4. Datos de Firma y Contacto Operativo Diario</h4>", unsafe_allow_html=True)
    col5, col6 = st.columns(2)
    with col5:
        nombre_firmante = st.text_input("Nombre Completo del Representante Legal que firmará el convenio:")
        nombre_contacto = st.text_input("Nombre Completo del Contacto de Recursos Humanos / Convenios:")
        correo_contacto = st.text_input("Correo Electrónico del Contacto de Recursos Humanos:")
    with col6:
        cargo_firmante = st.text_input("Cargo Legal de quien firma el convenio:", placeholder="Ej. Apoderado Legal / Director General")
        tel_contacto = st.text_input("Número Telefónico de Oficina o Móvil (10 dígitos):", max_chars=10)
        correo_registro = st.text_input("Correo Electrónico Institucional Principal:")
        
    st.markdown("<br>", unsafe_allow_html=True)
    submit_button = st.form_submit_button(label="Finalizar Registro e Iniciar Circuito de Convenio")

# 6. Procesamiento Inteligente de Datos
if submit_button:
    if not razon_social or not nombre_comercial or not rfc:
        st.error("⚠️ Error Mandatorio: Por favor llene los campos obligatorios para poder generar el marco legal (Razón Social, Nombre Comercial y RFC).")
    else:
        datos_empresa = {
            "fecha_registro": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "ejecutivo_comercial": ejecutivo,
            "tipo_entidad": tipo_entidad,
            "razon_social": razon_social,
            "nombre_comercial": nombre_comercial,
            "rfc": rfc,
            "giro": giro,
            "colaboradores": colaboradores,
            "nombre_firmante": nombre_firmante,
            "cargo_firmante": cargo_firmante,
            "nombre_contacto": nombre_contacto,
            "correo_contacto": correo_contacto,
            "tel_contacto": tel_contacto,
            "domicilio_fiscal": domicilio_fiscal,
            "domicilio_empresa": domicilio_empresa if domicilio_empresa else domicilio_fiscal
        }
        
        # Paso A: Generamos el borrador físico en la máquina
        archivo_generado = generar_documento_convenio(datos_empresa)
        
        if archivo_generado:
            # Paso B: Guardamos los datos estructurales en Google Sheets
            with st.spinner("Registrando su empresa en el Ecosistema B2B Michin..."):
                registro_guardado = guardar_en_sheets(datos_empresa)
                
            # Paso C: Subimos de forma transparente el Word al Drive de Santiago
            with st.spinner("Enviando borrador legal al equipo de validación comercial..."):
                id_drive = subir_a_google_drive(archivo_generado, nombre_comercial)
                
            if registro_guardado:
             st.session_state["archivo_listo"] = archivo_generado
             st.markdown("---")
             st.success(f"🎉 ¡Registro Exitoso, bienvenido a la red comercial Michin! Los datos de **{nombre_comercial}** han sido vinculados al CRM corporativo. El equipo legal, encabezado por Santiago Segoviano, ha recibido el documento para auditoría y seguimiento.")
             st.balloons()
             time.sleep(1.5) # Le damos 1.5 segundos de respiro a la pantalla
             st.rerun() # Refresca la interfaz de forma limpia
             
# 7. Respaldo de descarga para el cliente (Opcional por cortesía)
if "archivo_listo" in st.session_state:
    st.markdown("<p style='font-size: 13px; opacity: 0.7;'>Como respaldo, usted puede descargar una copia local de los datos procesados en su machote:</p>", unsafe_allow_html=True)
    with open(st.session_state["archivo_listo"], "rb") as file:
        st.download_button(
            label="Descargar Copia de Respaldo (.docx)", 
            data=file, 
            file_name=st.session_state["archivo_listo"], 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )