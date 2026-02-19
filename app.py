import streamlit as st
import os
import json
import time
import fitz  # PyMuPDF
from google import genai
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
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
# LÓGICA DE RESET
# ==========================================
if "reset_key" not in st.session_state:
    st.session_state.reset_key = 0

def reset_app():
    st.session_state.reset_key += 1

# ==========================================
# CONFIGURACIÓN DE GEMINI (CON ROBUSTEZ)
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)
MODELO = "gemini-1.5-flash" # Más estable para límites de cuota

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
    
    for intento in range(3):
        try:
            response = client.models.generate_content(
                model=MODELO,
                contents=prompt,
                config={'response_mime_type': 'application/json'}
            )
            resultado = json.loads(response.text)
            
            # CORRECCIÓN: Si devuelve una lista, tomar el primer objeto
            if isinstance(resultado, list):
                resultado = resultado[0] if resultado else {}
                
            return resultado
        except Exception as e:
            if "429" in str(e) and intento < 2:
                time.sleep(5) # Espera 5 segundos si la cuota se agotó
                continue
            raise e

# ==========================================
# FUNCIONES DE PROCESAMIENTO PPTX
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

def llenar_shape_con_titulos(slide, placeholder, texto, title_size=Pt(11), body_size=Pt(9.5), font_color=(0, 0, 0)):
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or placeholder not in (shape.text or ""):
            continue

        tf = shape.text_frame
        tf.text = ""
        lines = [ln.rstrip() for ln in texto.splitlines() if ln.strip() != ""]

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            stripped = line.strip()
            is_title = stripped.endswith(':') or (stripped.isupper() and len(stripped) > 2) or stripped.startswith('•')

            if not p.runs: p.add_run()
            for run in p.runs:
                run.font.bold = is_title
                run.font.size = title_size if is_title else body_size
                run.font.color.rgb = RGBColor(*font_color)
        return True
    return False

def llenar_experiencias(slide, placeholder, experiencias, font_color=(0, 0, 0)):
    for shape in slide.shapes:
        if not hasattr(shape, "text_frame") or placeholder not in (shape.text or ""):
            continue

        tf = shape.text_frame
        tf.text = ""
        
        p_titulo = tf.paragraphs[0]
        p_titulo.text = "EXPERIENCIA LABORAL:"
        for run in p_titulo.runs:
            run.font.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(*font_color)

        for exp in experiencias:
            tf.add_paragraph().text = "" # Espacio separador
            
            p_h = tf.add_paragraph()
            p_h.text = f"• {exp.get('empresa','?')} | {exp.get('puesto','?')} ({exp.get('periodo','?')})"
            for run in p_h.runs:
                run.font.bold = True
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(*font_color)

            desc = exp.get('descripcion', '')
            for d_line in desc.splitlines():
                if d_line.strip():
                    p_d = tf.add_paragraph()
                    p_d.text = f"  {d_line.strip()}"
                    for run in p_d.runs:
                        run.font.size = Pt(9.5)
                        run.font.color.rgb = RGBColor(*font_color)
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
        slide.shapes.add_picture(tmp_path, Inches(0.2), Inches(0.3), width=Inches(1.5))
        os.unlink(tmp_path)

def actualizar_encabezado(slide, nombre, rol):
    for shape in slide.shapes:
        if hasattr(shape, "text") and any(x in shape.text for x in ["{{NOMBRE}}", "COLABORADOR PROPUESTO"]):
            tf = shape.text_frame
            tf.text = ""
            p1 = tf.paragraphs[0]
            p1.text = f"COLABORADOR PROPUESTO – {nombre}"
            for run in p1.runs:
                run.font.bold = True
                run.font.size = Pt(13)
                run.font.color.rgb = RGBColor(88, 24, 139)
            p2 = tf.add_paragraph()
            p2.text = rol
            for run in p2.runs:
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(88, 24, 139)
            return

def generar_pptx(datos, template_bytes, foto_bytes):
    # Validación extra de formato de datos
    if isinstance(datos, list):
        datos = datos[0] if datos else {}

    prs = Presentation(BytesIO(template_bytes))
    nombre = datos.get('nombre', 'N/A')
    rol = datos.get('rol', '')
    exps = datos.get('experiencias', [])

    # Slide 1
    s1 = prs.slides[0]
    actualizar_encabezado(s1, nombre, rol)
    llenar_shape_con_titulos(s1, "{{PERFIL}}", f"PERFIL:\n{datos.get('perfil','')}", Pt(11), Pt(9.5), (255,255,255))
    hab = f"HABILIDADES:\n{datos.get('habilidades_tecnicas','')}\n\nHERRAMIENTAS:\n{datos.get('herramientas','')}"
    llenar_shape_con_titulos(s1, "{{HABILIDADES}}", hab, Pt(11), Pt(9.5), (255,255,255))
    edu = f"EDUCACIÓN:\n{datos.get('educacion','')}\n\nIDIOMAS:\n{datos.get('idiomas','')}"
    llenar_shape_con_titulos(s1, "{{EDUCACION}}", edu, Pt(11), Pt(9.5), (255,255,255))
    llenar_experiencias(s1, "{{EXPERIENCIA_1_2}}", exps[:2], (0,0,0))
    agregar_foto(s1, foto_bytes)

    # Slides siguientes
    rem = exps[2:]
    for i in range(1, len(prs.slides)):
        slide = prs.slides[i]
        actualizar_encabezado(slide, nombre, rol)
        if rem:
            llenar_experiencias(slide, "{{EXPERIENCIA_3_PLUS}}", rem[:4], (0,0,0))
            rem = rem[4:]
        agregar_foto(slide, foto_bytes)

    out = BytesIO(); prs.save(out); out.seek(0)
    return out, nombre, len(exps)

# ==========================================
# INTERFAZ DE STREAMLIT
# ==========================================
st.markdown('<div class="main-header"><h1>🎨 TRANSFORMADOR HORIZON CV</h1></div>', unsafe_allow_html=True)

DEFAULT_TEMPLATE = "CV HORIZON-MODELO 2.pptx"
template_bytes = None
t_file = st.file_uploader("Template (.pptx)", type=['pptx'], key=f"t_{st.session_state.reset_key}")

if t_file:
    template_bytes = t_file.read()
elif os.path.exists(DEFAULT_TEMPLATE):
    with open(DEFAULT_TEMPLATE, "rb") as f: template_bytes = f.read()

cv_file = st.file_uploader("CV (.pdf)", type=['pdf'], key=f"cv_{st.session_state.reset_key}")

if template_bytes and cv_file:
    if st.button("🚀 TRANSFORMAR"):
        with st.spinner("⏳ Procesando..."):
            try:
                texto, foto = extraer_foto_y_texto(cv_file.read())
                datos = redactar_con_gemini(texto)
                out, nom, n_exp = generar_pptx(datos, template_bytes, foto)
                st.success(f"✅ ¡Completado! Candidato: {nom}")
                st.download_button("📥 DESCARGAR", out, f"CV_Horizon_{nom}.pptx")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
