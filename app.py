import streamlit as st
import os
import json
import fitz  # PyMuPDF
from google import genai
from pptx import Presentation
from pptx.util import Pt, Inches
import tempfile
from io import BytesIO

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(page_title="Horizon CV Transformer v2", page_icon="🎨", layout="centered")

st.markdown("""
<style>
    .main-header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 2rem; border-radius: 10px; text-align: center; margin-bottom: 2rem; }
    .success-box { background: #d4edda; border: 2px solid #28a745; padding: 1.5rem; border-radius: 10px; margin-top: 1rem; }
    .stButton>button { width: 100%; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-size: 18px; font-weight: bold; padding: 0.75rem; border-radius: 10px; border: none; }
</style>
""", unsafe_allow_html=True)

if "reset_key" not in st.session_state: st.session_state.reset_key = 0
def reset_app(): st.session_state.reset_key += 1

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
    Eres un experto en reclutamiento para Horizon Consulting. 
    Extrae la información del CV y devuélvela en JSON. 
    IMPORTANTE: No resumas las experiencias, mantén el detalle técnico.
    
    ESTRUCTURA JSON:
    {{ 
      "nombre": "", "puesto_propuesto": "", 
      "perfil": "", 
      "experiencias": [{{ "empresa": "", "puesto": "", "periodo": "", "descripcion": "" }}], 
      "educacion_certificaciones": "", 
      "habilidades_idiomas": "" 
    }}
    CV: {texto_cv}
    """
    response = client.models.generate_content(model=MODELO, contents=prompt, config={'response_mime_type': 'application/json'})
    return json.loads(response.text)

def llenar_shape_con_estilo(slide, placeholder, texto, title_size=Pt(10), body_size=Pt(8)):
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
            is_title = line.strip().endswith(':') or (line.strip().isupper() and len(line.strip()) > 2) or line.strip().startswith('•')
            if not p.runs: p.add_run()
            for run in p.runs:
                run.font.bold = is_title
                run.font.size = title_size if is_title else body_size
        return True
    return False

def manejar_foto(slide, foto_bytes):
    for shape in [s for s in slide.shapes if hasattr(s, "text") and "FOTO" in s.text]:
        sp = shape.element
        sp.getparent().remove(sp)
    if foto_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            tmp.write(foto_bytes)
            tmp_path = tmp.name
        # Posición ajustada para el nuevo modelo
        slide.shapes.add_picture(tmp_path, Inches(8.2), Inches(0.4), width=Inches(1.3))
        os.unlink(tmp_path)

def generar_pptx(datos, template_bytes, foto_bytes):
    prs = Presentation(BytesIO(template_bytes))
    nombre = datos.get('nombre', 'N/A')
    puesto = datos.get('puesto_propuesto', 'Consultor')
    exps = datos.get('experiencias', [])

    # --- PROCESAR TODAS LAS LÁMINAS ---
    for i, slide in enumerate(prs.slides):
        llenar_shape_con_estilo(slide, "{{NOMBRE}}", nombre, Pt(14), Pt(14))
        llenar_shape_con_estilo(slide, "{{PUESTO}}", puesto, Pt(11), Pt(11))
        manejar_foto(slide, foto_bytes)

        if i == 0: # Lámina 1
            llenar_shape_con_estilo(slide, "{{PERFIL}}", f"PERFIL:\n{datos.get('perfil','')}", Pt(10), Pt(8))
            llenar_shape_con_estilo(slide, "{{EDUCACION}}", f"EDUCACIÓN Y CERTIFICACIONES:\n{datos.get('educacion_certificaciones','')}", Pt(10), Pt(8))
            llenar_shape_con_estilo(slide, "{{HABILIDADES}}", f"HABILIDADES E IDIOMAS:\n{datos.get('habilidades_idiomas','')}", Pt(10), Pt(8))
            
            txt_exp12 = "EXPERIENCIA LABORAL:\n\n"
            for e in exps[:2]:
                txt_exp12 += f"• {e['empresa']} | {e['puesto']} ({e['periodo']})\n  {e['descripcion']}\n\n"
            llenar_shape_con_estilo(slide, "{{EXPERIENCIA_1_2}}", txt_exp12, Pt(10), Pt(8))

        elif i == 1: # Lámina 2
            txt_resto = "EXPERIENCIA LABORAL (Cont.):\n\n"
            for e in exps[2:6]: # De la 3 a la 6
                txt_resto += f"• {e['empresa']} | {e['puesto']} ({e['periodo']})\n  {e['descripcion']}\n\n"
            llenar_shape_con_estilo(slide, "{{EXPERIENCIA_RESTO}}", txt_resto, Pt(10), Pt(8))

        elif i == 2: # Lámina 3
            txt_final = "EXPERIENCIA LABORAL (Final):\n\n"
            for e in exps[6:]: # De la 7 en adelante
                txt_final += f"• {e['empresa']} | {e['puesto']} ({e['periodo']})\n  {e['descripcion']}\n\n"
            llenar_shape_con_estilo(slide, "{{EXPERIENCIA_FINAL}}", txt_final, Pt(10), Pt(8))

    out = BytesIO(); prs.save(out); out.seek(0)
    return out, nombre

# ==========================================
# INTERFAZ
# ==========================================
st.markdown('<div class="main-header"><h1>🎨 HORIZON TRANSFORMER V2</h1></div>', unsafe_allow_html=True)

DEFAULT_TEMPLATE = "CV HORIZON-MODELO 2.pptx"
template_bytes = None

t_file = st.file_uploader("Cambiar Modelo (.pptx)", type=['pptx'], key=f"t_{st.session_state.reset_key}")
if t_file:
    template_bytes = t_file.read()
elif os.path.exists(DEFAULT_TEMPLATE):
    with open(DEFAULT_TEMPLATE, "rb") as f: template_bytes = f.read()
    st.info(f"ℹ️ Usando: {DEFAULT_TEMPLATE}")

cv_file = st.file_uploader("Subir CV Candidato (.pdf)", type=['pdf'], key=f"cv_{st.session_state.reset_key}")

if template_bytes and cv_file:
    if st.button("🚀 TRANSFORMAR"):
        with st.spinner("Procesando..."):
            try:
                texto, foto = extraer_foto_y_texto(cv_file.read())
                datos = redactar_con_gemini(texto)
                out, nom = generar_pptx(datos, template_bytes, foto)
                st.success(f"✅ {nom} procesado con éxito")
                st.download_button("📥 DESCARGAR", out, f"CV_Horizon_{nom}.pptx")
            except Exception as e:
                st.error(f"Error: {e}")
