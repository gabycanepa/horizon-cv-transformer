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
# CONFIGURACIÓN DE GEMINI
# ==========================================
# Puedes cambiar esto por st.secrets["GEMINI_API_KEY"] para mayor seguridad
API_KEY = "st.secrets["GEMINI_API_KEY"]"
client = genai.Client(api_key=API_KEY)
MODELO = "gemini-2.0-flash"

# ==========================================
# FUNCIONES DE PROCESAMIENTO
# ==========================================

def extraer_foto_y_texto(pdf_bytes):
    """Extrae texto y foto del PDF"""
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
    """Procesa el CV con Gemini y devuelve JSON estructurado"""
    prompt = f"""
    Eres un transcriptor de datos de alta fidelidad para Horizon Consulting. 
    Tu única misión es NO PERDER NINGUNA EXPERIENCIA LABORAL.

    PASO 1: Identifica TODAS las empresas y periodos mencionados en el CV.
    PASO 2: Para CADA UNA de esas empresas, extrae el puesto y la descripción completa.
    
    REGLAS DE ORO:
    - Si el CV tiene 6 experiencias, el JSON DEBE tener 6 experiencias.
    - NO resumas. NO omitas las antiguas. NO combines puestos.
    - Traduce al español manteniendo el rigor técnico.

    ESTRUCTURA JSON:
    {{ 
      "nombre": "", 
      "rol": "", 
      "contacto": {{ "telefono": "", "email": "", "ubicacion": "", "linkedin": "" }},
      "perfil": "", 
      "experiencias": [
        {{ "empresa": "", "puesto": "", "periodo": "", "descripcion": "" }}
      ], 
      "educacion": "", 
      "habilidades_tecnicas": "", 
      "herramientas": "",
      "expertise": "", 
      "certificaciones": "", 
      "idiomas": "" 
    }}

    CV A PROCESAR:
    {texto_cv}
    """
    
    response = client.models.generate_content(
        model=MODELO, 
        contents=prompt, 
        config={'response_mime_type': 'application/json'}
    )
    return json.loads(response.text)

def ajustar_fuente(shape, size=9):
    """Ajusta el tamaño de fuente de un shape"""
    if hasattr(shape, "text_frame"):
        for p in shape.text_frame.paragraphs:
            for run in p.runs: 
                run.font.size = Pt(size)

def reemplazar_placeholder(slide, placeholder, nuevo_texto, font_size=9):
    """Reemplaza un placeholder en el slide"""
    for shape in slide.shapes:
        if hasattr(shape, "text_frame"):
            if placeholder in shape.text:
                shape.text = nuevo_texto
                ajustar_fuente(shape, font_size)
                return True
    return False

def eliminar_cuadro_foto(slide):
    """Elimina el cuadro placeholder de FOTO"""
    shapes_to_delete = []
    for shape in slide.shapes:
        if hasattr(shape, "text") and "FOTO" in shape.text:
            shapes_to_delete.append(shape)
    for shape in shapes_to_delete:
        sp = shape.element
        sp.getparent().remove(sp)

def agregar_foto(slide, foto_bytes):
    """Agrega la foto del candidato al slide"""
    if foto_bytes:
        eliminar_cuadro_foto(slide)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            tmp.write(foto_bytes)
            tmp_path = tmp.name
        slide.shapes.add_picture(tmp_path, Inches(8.5), Inches(0.5), width=Inches(1.2))
        os.unlink(tmp_path)

def actualizar_encabezado(slide, nombre, rol):
    """Actualiza el encabezado del slide"""
    for shape in slide.shapes:
        if hasattr(shape, "text"):
            if "Colaborador propuesto" in shape.text or "Edwin" in shape.text or "Nombre" in shape.text:
                shape.text = f"Colaborador propuesto – {nombre}\n{rol}"
                ajustar_fuente(shape, 11)
                return

def generar_pptx(datos, template_bytes, foto_bytes):
    """Genera el PPTX final"""
    # Cargar template desde bytes
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        tmp.write(template_bytes)
        tmp_template = tmp.name
    
    prs = Presentation(tmp_template)
    nombre = datos.get('nombre', 'N/A')
    rol = datos.get('rol', '')
    
    # SLIDE 1: PERFIL COMPLETO
    s1 = prs.slides[0]
    actualizar_encabezado(s1, nombre, rol)
    
    c = datos.get('contacto', {})
    perfil_completo = f"PERFIL PROFESIONAL:\n{datos.get('perfil', '')}\n\n"
    perfil_completo += f"CONTACTO:\n"
    if c.get('email'): perfil_completo += f"• Email: {c['email']}\n"
    if c.get('telefono'): perfil_completo += f"• Teléfono: {c['telefono']}\n"
    if c.get('ubicacion'): perfil_completo += f"• Ubicación: {c['ubicacion']}\n"
    if c.get('linkedin'): perfil_completo += f"• LinkedIn: {c['linkedin']}\n"
    perfil_completo += f"\nEDUCACIÓN:\n{datos.get('educacion', '')}\n\n"
    if datos.get('certificaciones'):
        perfil_completo += f"CERTIFICACIONES:\n{datos['certificaciones']}\n\n"
    perfil_completo += f"HABILIDADES TÉCNICAS:\n{datos.get('habilidades_tecnicas', '')}\n\n"
    perfil_completo += f"HERRAMIENTAS:\n{datos.get('herramientas', '')}\n\n"
    perfil_completo += f"EXPERTISE:\n{datos.get('expertise', '')}\n\n"
    perfil_completo += f"IDIOMAS:\n{datos.get('idiomas', '')}"
    
    reemplazar_placeholder(s1, "{{PERFIL_COMPLETO}}", perfil_completo, 8)
    agregar_foto(s1, foto_bytes)

    # EXPERIENCIAS
    experiencias = datos.get('experiencias', [])
    exp_slide2 = experiencias[:2]
    exp_slide3 = experiencias[2:]

    # SLIDE 2: EXPERIENCIA 1-2
    if len(prs.slides) > 1:
        s2 = prs.slides[1]
        actualizar_encabezado(s2, nombre, rol)
        
        txt_exp_1_2 = "EXPERIENCIA LABORAL:\n\n"
        for exp in exp_slide2:
            txt_exp_1_2 += f"• {exp['empresa']} | {exp['puesto']} ({exp['periodo']})\n"
            txt_exp_1_2 += f"  {exp['descripcion']}\n\n"
        
        reemplazar_placeholder(s2, "{{EXPERIENCIA_1_2}}", txt_exp_1_2, 9)
        agregar_foto(s2, foto_bytes)

    # SLIDE 3: EXPERIENCIA 3+
    if len(exp_slide3) > 0 and len(prs.slides) > 2:
        s3 = prs.slides[2]
        actualizar_encabezado(s3, nombre, rol)
        
        txt_exp_3_plus = "EXPERIENCIA LABORAL (Continuación):\n\n"
        for exp in exp_slide3:
            txt_exp_3_plus += f"• {exp['empresa']} | {exp['puesto']} ({exp['periodo']})\n"
            txt_exp_3_plus += f"  {exp['descripcion']}\n\n"
        
        reemplazar_placeholder(s3, "{{EXPERIENCIA_3_PLUS}}", txt_exp_3_plus, 7)
        agregar_foto(s3, foto_bytes)

    # Guardar en BytesIO
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    
    # Limpiar archivo temporal
    os.unlink(tmp_template)
    
    return output, nombre, len(experiencias)

# ==========================================
# INTERFAZ DE STREAMLIT
# ==========================================

# Header
st.markdown("""
<div class="main-header">
    <h1>🎨 TRANSFORMADOR HORIZON CV</h1>
    <p>Convierte CVs al formato Horizon automáticamente</p>
</div>
""", unsafe_allow_html=True)

# Instrucciones
with st.expander("ℹ️ Instrucciones de uso"):
    st.markdown("""
    1. **Sube el Template Horizon** (archivo .pptx con los placeholders)
    2. **Sube el CV del candidato** (archivo .pdf)
    3. Haz clic en **"Transformar a Formato Horizon"**
    4. Descarga el resultado automáticamente
    
    **Nota:** El proceso toma entre 10-30 segundos dependiendo del tamaño del CV.
    """)

st.markdown("---")

# Upload Template
st.markdown("### 📁 Paso 1: Subir Template Horizon")
template_file = st.file_uploader(
    "Selecciona el archivo .pptx del template",
    type=['pptx'],
    help="Debe contener los placeholders: {{PERFIL_COMPLETO}}, {{EXPERIENCIA_1_2}}, {{EXPERIENCIA_3_PLUS}}"
)

if template_file:
    st.success(f"✓ Template cargado: {template_file.name}")

st.markdown("---")

# Upload CV
st.markdown("### 📄 Paso 2: Subir CV del Candidato")
cv_file = st.file_uploader(
    "Selecciona el archivo .pdf del CV",
    type=['pdf'],
    help="El CV debe estar en formato PDF"
)

if cv_file:
    st.success(f"✓ CV cargado: {cv_file.name}")

st.markdown("---")

# Botón de procesamiento
if template_file and cv_file:
    if st.button("🚀 TRANSFORMAR A FORMATO HORIZON"):
        with st.spinner("⏳ Procesando..."):
            try:
                # Leer archivos
                template_bytes = template_file.read()
                cv_bytes = cv_file.read()
                
                # Paso 1: Extraer
                st.info("✓ Extrayendo información del PDF...")
                texto, foto_bytes = extraer_foto_y_texto(cv_bytes)
                
                # Paso 2: Procesar con Gemini
                st.info("✓ Analizando con Gemini...")
                datos_json = redactar_con_gemini(texto)
                
                # Paso 3: Generar PPTX
                st.info("✓ Generando presentación Horizon...")
                output_pptx, nombre, num_exp = generar_pptx(datos_json, template_bytes, foto_bytes)
                
                # Éxito
                st.markdown(f"""
                <div class="success-box">
                    <h3 style='color: #155724; margin: 0;'>✅ ¡TRANSFORMACIÓN COMPLETADA!</h3>
                    <p style='margin: 10px 0;'><strong>Candidato:</strong> {nombre}</p>
                    <p style='margin: 10px 0;'><strong>Experiencias detectadas:</strong> {num_exp}</p>
                </div>
                """, unsafe_allow_html=True)
                
                # Botón de descarga
                st.download_button(
                    label="📥 DESCARGAR CV HORIZON",
                    data=output_pptx,
                    file_name=f"CV_Horizon_{nombre.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {str(e)}")
                st.exception(e)
else:
    st.info("👆 Por favor sube ambos archivos para continuar")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 1rem;'>
    <p>Desarrollado para Horizon Consulting | Powered by Gemini 2.0 Flash</p>
</div>
""", unsafe_allow_html=True)
