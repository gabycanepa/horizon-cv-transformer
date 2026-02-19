import streamlit as st
import os
import json
import re
import fitz  # PyMuPDF
from google import genai
from pptx import Presentation
from pptx.util import Pt, Inches
import tempfile
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Transformador Horizon CV",
    page_icon="🎨",
    layout="centered"
)

# ==========================================
# ESTILOS CSS
# ==========================================
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 2rem;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background: #d4edda;
        border: 2px solid #28a745;
        padding: 1.5rem;
        border-radius: 10px;
        margin-top: 1rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-size: 18px;
        font-weight: bold;
        padding: 0.75rem;
        border-radius: 10px;
        border: none;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #764ba2 0%, #667eea 100%);
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# LÓGICA DE RESET (session_state)
# ==========================================
if "reset_key" not in st.session_state:
    st.session_state.reset_key = 0

def reset_app():
    st.session_state.reset_key += 1

# ==========================================
# CONFIGURACIÓN DE GEMINI
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)
MODELO = "gemini-2.0-flash"

# ==========================================
# FUNCIONES DE PROCESAMIENTO
# ==========================================
def extraer_foto_y_texto(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto_completo = ""
    foto_bytes = None
    for page in doc:
        texto_completo += page.get_text()
        if not foto_bytes:
            images = page.get_images(full=True)
            if images:
                xref = images[0][0]
                base_image = doc.extract_image(xref)
                foto_bytes = base_image["image"]
    doc.close()
    return texto_completo, foto_bytes

def redactar_con_gemini(texto_cv):
    prompt = f"""
    Eres un transcriptor de datos de alta fidelidad para Horizon Consulting. 
    Tu única misión es NO PERDER NINGUNA EXPERIENCIA LABORAL.
    ESTRUCTURA JSON:
    {{ 
      "nombre": "", "rol": "", 
      "contacto": {{ "telefono": "", "email": "", "ubicacion": "", "linkedin": "" }},
      "perfil": "", 
      "experiencias": [{{ "empresa": "", "puesto": "", "periodo": "", "descripcion": "" }}], 
      "educacion": "", "habilidades_tecnicas": "", "herramientas": "",
      "expertise": "", "certificaciones": "", "idiomas": "" 
    }}
    CV A PROCESAR:
    {texto_cv}
    """
    response = client.models.generate_content(model=MODELO, contents=prompt, config={'response_mime_type': 'application/json'})
    return json.loads(response.text)

def ajustar_fuente(shape, size=9):
    if hasattr(shape, "text_frame"):
        for p in shape.text_frame.paragraphs:
            for run in p.runs: run.font.size = Pt(size)

def llenar_shape_con_titulos(slide, placeholder, texto, title_size=Pt(10), body_size=Pt(8)):
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame"): continue
        try:
            if placeholder not in shape.text: continue
        except: continue
        tf = shape.text_frame
        tf.text = ""
        lines = [ln.rstrip() for ln in texto.splitlines() if ln.strip() != ""]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            is_title = False
            stripped = line.strip()
            if stripped.endswith(':') or (stripped.isupper() and len(stripped) > 2) or stripped.startswith('•') or stripped.startswith('-'):
                is_title = True
            if not p.runs: p.add_run()
            for run in p.runs:
                run.font.bold = is_title
                run.font.size = title_size if is_title else body_size
        return True
    return False

def eliminar_cuadro_foto(slide):
    shapes_to_delete = [s for s in slide.shapes if hasattr(s, "text") and "FOTO" in s.text]
    for shape in shapes_to_delete:
        sp = shape.element
        sp.getparent().remove(sp)

def agregar_foto(slide, foto_bytes):
    if foto_bytes:
        eliminar_cuadro_foto(slide)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            tmp.write(foto_bytes)
            tmp_path = tmp.name
        slide.shapes.add_picture(tmp_path, Inches(8.5), Inches(0.5), width=Inches(1.2))
        os.unlink(tmp_path)

def actualizar_encabezado(slide, nombre, rol):
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            if any(x in shape.text for x in ["Colaborador propuesto", "Edwin", "Nombre"]):
                shape.text = f"Colaborador propuesto – {nombre}\n{rol}"
                ajustar_fuente(shape, 11)
                return

def generar_pptx(datos, template_bytes, foto_bytes):
    prs = Presentation(BytesIO(template_bytes))
    nombre, rol = datos.get('nombre', 'N/A'), datos.get('rol', '')
    s1 = prs.slides[0]
    actualizar_encabezado(s1, nombre, rol)
    c = datos.get('contacto', {})
    perfil_txt = f"PERFIL PROFESIONAL:\n{datos.get('perfil', '')}\n\nCONTACTO:\n"
    if c.get('email'): perfil_txt += f"• Email: {c['email']}\n"
    if c.get('telefono'): perfil_txt += f"• Teléfono: {c['telefono']}\n"
    perfil_txt += f"\nEDUCACIÓN:\n{datos.get('educacion', '')}\n\nHABILIDADES TÉCNICAS:\n{datos.get('habilidades_tecnicas', '')}\n\nIDIOMAS:\n{datos.get('idiomas', '')}"
    llenar_shape_con_titulos(s1, "{{PERFIL_COMPLETO}}", perfil_txt, Pt(10), Pt(8))
    agregar_foto(s1, foto_bytes)
    exps = datos.get('experiencias', [])
    if len(prs.slides) > 1:
        s2 = prs.slides[1]
        actualizar_encabezado(s2, nombre, rol)
        txt2 = "EXPERIENCIA LABORAL:\n\n"
        for exp in exps[:2]: txt2 += f"• {exp['empresa']} | {exp['puesto']} ({exp['periodo']})\n  {exp['descripcion']}\n\n"
        llenar_shape_con_titulos(s2, "{{EXPERIENCIA_1_2}}", txt2, Pt(10), Pt(9))
        agregar_foto(s2, foto_bytes)
    if len(exps) > 2 and len(prs.slides) > 2:
        s3 = prs.slides[2]
        actualizar_encabezado(s3, nombre, rol)
        txt3 = "EXPERIENCIA LABORAL (Cont.):\n\n"
        for exp in exps[2:]: txt3 += f"• {exp['empresa']} | {exp['puesto']} ({exp['periodo']})\n  {exp['descripcion']}\n\n"
        llenar_shape_con_titulos(s3, "{{EXPERIENCIA_3_PLUS}}", txt3, Pt(10), Pt(8))
        agregar_foto(s3, foto_bytes)
    out = BytesIO(); prs.save(out); out.seek(0)
    return out, nombre, len(exps)

# ==========================================
# INTERFAZ DE STREAMLIT
# ==========================================
st.markdown('<div class="main-header"><h1>🎨 TRANSFORMADOR HORIZON CV</h1><p>Convierte CVs al formato Horizon automáticamente</p></div>', unsafe_allow_html=True)

with st.expander("ℹ️ Instrucciones de uso"):
    st.markdown("1. Sube el CV del candidato (.pdf)\n2. Haz clic en 'Transformar'\n3. Descarga el resultado.")

st.markdown("---")

# PASO 1: Template (Precargado)
col_main, col_btn = st.columns([8,1])
with col_main: st.markdown("### 📁 Paso 1: Template Horizon")
with col_btn: st.button("🧹", on_click=reset_app, key="btn_limpiar")

DEFAULT_TEMPLATE = "Modelo de CV Horizon.pptx"
template_bytes = None
template_file = st.file_uploader("Sube un nuevo .pptx si quieres cambiar el modelo", type=['pptx'], key=f"t_{st.session_state.reset_key}")

if template_file:
    template_bytes = template_file.read()
    st.success(f"✓ Usando template subido: {template_file.name}")
elif os.path.exists(DEFAULT_TEMPLATE):
    with open(DEFAULT_TEMPLATE, "rb") as f: template_bytes = f.read()
    st.info(f"ℹ️ Usando modelo estándar: {DEFAULT_TEMPLATE}")
else:
    st.warning("⚠️ Sube un template .pptx para comenzar.")

st.markdown("---")

# PASO 2: CV
st.markdown("### 📄 Paso 2: Subir CV del Candidato")
cv_file = st.file_uploader("Selecciona el archivo .pdf del CV", type=['pdf'], key=f"cv_{st.session_state.reset_key}")

if cv_file:
    st.success(f"✓ CV cargado: {cv_file.name}")

st.markdown("---")

# BOTÓN DE PROCESAMIENTO (Versión Completa)
if template_bytes and cv_file:
    if st.button("🚀 TRANSFORMAR A FORMATO HORIZON"):
        with st.spinner("⏳ Procesando..."):
            try:
                cv_bytes = cv_file.read()
                st.info("✓ Extrayendo información del PDF...")
                texto, foto_bytes = extraer_foto_y_texto(cv_bytes)
                
                st.info("✓ Analizando con Gemini...")
                datos_json = redactar_con_gemini(texto)
                
                st.info("✓ Generando presentación Horizon...")
                output_pptx, nombre, num_exp = generar_pptx(datos_json, template_bytes, foto_bytes)
                
                st.markdown(f"""
                <div class="success-box">
                    <h3 style='color: #155724; margin: 0;'>✅ ¡TRANSFORMACIÓN COMPLETADA!</h3>
                    <p style='margin: 10px 0;'><strong>Candidato:</strong> {nombre}</p>
                    <p style='margin: 10px 0;'><strong>Experiencias detectadas:</strong> {num_exp}</p>
                </div>
                """, unsafe_allow_html=True)
                
                st.download_button(
                    label="📥 DESCARGAR CV HORIZON",
                    data=output_pptx,
                    file_name=f"CV_Horizon_{nombre.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
else:
    st.info("👆 Por favor sube el CV para continuar")

st.markdown("---")
st.markdown('<div style="text-align: center; color: #666; padding: 1rem;"><p>Desarrollado por Horizon Consulting | Powered by Gemini 2.0 Flash</p></div>', unsafe_allow_html=True)
