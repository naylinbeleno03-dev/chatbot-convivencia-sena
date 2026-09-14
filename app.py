import io
import os
import re
import time
import streamlit as st
from google import genai
from google.genai import types

# Intentar importar python-docx para crear archivos .docx reales
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==============================================================================
# ENLACES NORMATIVOS
# ==============================================================================
URL_CONSTITUCION_POLITICA = "https://www.registraduria.gov.co/IMG/pdf/constitucio-politica-colombia-1991.pdf"
URL_LEY_115 = (
    "https://www.mineducacion.gov.co/1621/articles-85906_archivo_pdf.pdf"
)
URL_LEY_1098 = (
    "https://www.icbf.gov.co/sites/default/files/codigoinfancialey1098.pdf"
)
URL_LEY_1620 = "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=52287"
URL_DECRETO_1965 = "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=54537"
URL_MANUAL_CONVIVENCIA = "https://drive.google.com/file/d/10WqGY5EvXzCMPROBZB6Ga6J4zjrEzmOG/view?usp=sharing"
# ==============================================================================

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema Integral de Convivencia Escolar - Institución Educativa Técnica Sagrado Corazón",
    layout="centered",
)

# Estilos CSS personalizados
st.markdown(
    """
    <style>
    /* 1. Fondo principal en Gris Clarito */
    .stApp {
        background-color: #F3F4F6;
        color: #1F2937;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* 2. Barra lateral en Gris un poco más oscuro */
    section[data-testid="stSidebar"] {
        background-color: #475569;
        border-right: 1px solid #334155;
    }

    /* Textos de la barra lateral en blanco */
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] li,
    section[data-testid="stSidebar"] a {
        color: #FFFFFF !important;
    }

    /* Visibilidad del desplegable (st.expander) en la barra lateral */
    section[data-testid="stSidebar"] details {
        background-color: #334155 !important;
        border: 1px solid #64748B !important;
        border-radius: 8px !important;
        padding: 6px 10px !important;
    }

    section[data-testid="stSidebar"] summary {
        color: #FFFFFF !important;
        background-color: #334155 !important;
        font-weight: 600 !important;
    }

    section[data-testid="stSidebar"] details[open] {
        background-color: #334155 !important;
    }

    section[data-testid="stSidebar"] details p,
    section[data-testid="stSidebar"] details li {
        color: #FFFFFF !important;
    }

    /* Resaltado para los enlaces normativos */
    section[data-testid="stSidebar"] details a {
        color: #FDE047 !important;
        text-decoration: underline !important;
        font-weight: 600 !important;
    }

    /* 3. Encabezado superior */
    .header-box {
        background-color: #0F172A;
        padding: 22px;
        border-radius: 12px;
        border-left: 6px solid #D4AF37;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2);
        margin-bottom: 24px;
    }

    .header-box h1 {
        color: #FFFFFF !important;
        margin: 0 !important;
        font-size: 1.65rem !important;
        font-weight: 600 !important;
    }

    .header-box p {
        color: #E2E8F0 !important;
        margin-top: 6px !important;
        margin-bottom: 0 !important;
        font-size: 0.95rem !important;
    }

    /* Botón de reiniciar consulta en el menú lateral */
    section[data-testid="stSidebar"] .stButton>button,
    section[data-testid="stSidebar"] .stButton>button p,
    section[data-testid="stSidebar"] .stButton>button div,
    section[data-testid="stSidebar"] .stButton>button span {
        background-color: #FFFFFF !important;
        color: #1F2937 !important;
        border: 1px solid #FFFFFF !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        width: 100%;
    }

    section[data-testid="stSidebar"] .stButton>button:hover,
    section[data-testid="stSidebar"] .stButton>button:hover p,
    section[data-testid="stSidebar"] .stButton>button:hover div,
    section[data-testid="stSidebar"] .stButton>button:hover span {
        background-color: #000000 !important;
        color: #FFFFFF !important;
        border-color: #000000 !important;
    }

    /* Botones de opciones predeterminadas en el panel principal */
    .stMainBlockContainer div.stButton > button {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        padding: 10px 14px !important;
        font-weight: 500 !important;
        width: 100%;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }

    .stMainBlockContainer div.stButton > button:hover {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        border-color: #0F172A !important;
    }

    /* Campo de clave API */
    .stTextInput>div>div>input {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
    }

    /* Tarjetas de mensajes en el chat */
    [data-testid="stChatMessage"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E5E7EB !important;
        border-left: 5px solid #0F172A !important;
        border-radius: 10px !important;
        padding: 16px !important;
        margin-bottom: 14px !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04) !important;
    }

    [data-testid="stChatMessage"] p, 
    [data-testid="stChatMessage"] div,
    [data-testid="stChatMessage"] span {
        color: #0F172A !important;
        font-size: 0.98rem;
        line-height: 1.6;
    }

    /* Campo de entrada de hechos */
    [data-testid="stChatInput"] {
        background-color: #FFFFFF !important;
        border: 2px solid #475569 !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05) !important;
    }

    [data-testid="stChatInput"] textarea {
        color: #000000 !important;
        background-color: #FFFFFF !important;
    }

    /* Botón de Enviar */
    [data-testid="stChatInput"] button {
        background-color: #000000 !important;
        border-color: #000000 !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
    }

    [data-testid="stChatInput"] button svg {
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }

    hr {
        border-color: #64748B;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Encabezado superior
st.markdown(
    """
    <div class="header-box">
        <h1>Sistema Integral de Convivencia Escolar</h1>
        <p>Institución Educativa Técnica Sagrado Corazón</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Obtener clave API automáticamente si está guardada en Secrets o pedirla
api_key = st.secrets.get("GEMINI_API_KEY", "")


def limpiar_texto_para_word(texto: str) -> str:
    """Limpia guiones, viñetas de markdown y líneas divisorias."""
    if not texto:
        return ""
    # Eliminar líneas divisorias (---, ***)
    texto = re.sub(r"^[-*]{3,}\s*$", "", texto, flags=re.MULTILINE)
    # Eliminar viñetas de markdown al inicio de línea
    texto = re.sub(r"^\s*[-\*]\s+", "", texto, flags=re.MULTILINE)
    # Eliminar guiones dobles
    texto = re.sub(r"-{2,}", "", texto)
    return texto.strip()


def extraer_solo_documento(texto_contenido: str) -> str:
    """Extrae exclusivamente el modelo o plantilla de documento para el archivo Word digital."""
    if not texto_contenido:
        return ""

    patrones = [
        r"(ACTA DE COMPROMISO.*)",
        r"(MODELO DE CARTA.*)",
        r"(MODELO DE REGISTRO.*)",
        r"(REGISTRO EN EL OBSERVADOR.*)",
        r"(CARTA DE DESCARGOS.*)",
        r"(CITACIÓN A ACUDIENTE.*)",
        r"(5\.\s*MODELO.*)",
        r"(5\.\s*DOCUMENTO.*)",
        r"(5\.\s*ACTA.*)",
        r"(5\.\s*REGISTRO.*)",
    ]

    for patron in patrones:
        match = re.search(patron, texto_contenido, re.IGNORECASE | re.DOTALL)
        if match:
            texto_extraido = match.group(1)
            lineas = texto_extraido.split("\n")
            if re.match(
                r"^5\.\s*(MODELO|DOCUMENTO|PLANTILLA|ACTA|REGISTRO)",
                lineas[0].strip(),
                re.IGNORECASE,
            ):
                texto_extraido = "\n".join(lineas[1:])
            return texto_extraido.strip()

    return texto_contenido.strip()


# Función para generar archivo de Microsoft Word (.docx) formal
def generar_documento_word(texto_contenido):
    doc = Document()

    # Extraer estrictamente solo la plantilla formal para la biblioteca digital
    texto_documento = extraer_solo_documento(texto_contenido)

    # Configuración de página y secciones
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

        # 1. ENCABEZADO INSTITUCIONAL (Esquina superior derecha, discreto)
        header = section.header
        p_head = header.paragraphs[0]
        p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_head.paragraph_format.space_after = Pt(0)

        posibles_nombres_imagen = [
            "escudo.png",
            "escudo.jpg",
            "escudo_intesac.png",
            "escudo_intesac.jpg",
            "logo.png",
            "logo.jpg",
        ]
        imagen_encontrada = None
        for nombre in posibles_nombres_imagen:
            if os.path.exists(nombre):
                imagen_encontrada = nombre
                break

        if imagen_encontrada:
            try:
                r_img = p_head.add_run()
                r_img.add_picture(imagen_encontrada, width=Inches(0.55))
                p_head.add_run("\n")
            except Exception:
                pass

        r_head = p_head.add_run(
            "INSTITUCIÓN EDUCATIVA TÉCNICA SAGRADO CORAZÓN\nSistema Digital de"
            " Seguimiento y Convivencia"
        )
        r_head.font.size = Pt(8)
        r_head.font.name = "Arial"
        r_head.font.bold = True
        r_head.font.color.rgb = RGBColor(100, 116, 139)

        # 2. PIE DE PÁGINA
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_foot = p_foot.add_run(
            "Institución Educativa Técnica Sagrado Corazón — Archivo Digital"
            " de Convivencia Escolar"
        )
        r_foot.font.size = Pt(8.5)
        r_foot.font.name = "Arial"
        r_foot.font.italic = True
        r_foot.font.color.rgb = RGBColor(100, 116, 139)

    # 3. CUERPO DEL DOCUMENTO
    lineas = texto_documento.split("\n")
    for linea in lineas:
        linea_limpia = limpiar_texto_para_word(linea)
        if not linea_limpia:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15

        # Detectar si la línea es el título principal del documento
        if (
            linea.strip().startswith("#")
            or (
                linea.strip().startswith("**")
                and linea.strip().endswith("**")
            )
            or "ACTA DE COMPROMISO" in linea.upper()
            or "MODELO DE CARTA" in linea.upper()
            or "CARTA DE DESCARGOS" in linea.upper()
            or "MODELO DE REGISTRO" in linea.upper()
            or "REGISTRO EN EL OBSERVADOR" in linea.upper()
        ):

            texto_titulo = linea_limpia.replace("#", "").replace("**", "")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(texto_titulo)
            run.bold = True
            run.font.size = Pt(12)
            run.font.name = "Arial"
            run.font.color.rgb = RGBColor(15, 23, 42)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(12)

        # Detectar líneas de firma para darles espaciado amplio superior
        elif "________________" in linea or "FIRMA" in linea.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(linea_limpia.replace("**", ""))
            run.font.size = Pt(10)
            run.font.name = "Arial"
            run.font.bold = "FIRMA" in linea.upper()
            run.font.color.rgb = RGBColor(31, 41, 55)
            # Dar espacio vertical suficiente (36pt) para la firma
            if "________________" in linea:
                p.paragraph_format.space_before = Pt(36)
                p.paragraph_format.space_after = Pt(2)
            else:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(4)

        else:
            # Alineación justificada para el texto normal
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Formato moderado evitando negritas innecesarias
            partes = re.split(r"(\*\*.*?\*\*)", linea_limpia)
            for parte in partes:
                if parte.startswith("**") and parte.endswith("**"):
                    run = p.add_run(parte[2:-2])
                    run.bold = True
                else:
                    run = p.add_run(parte)

                run.font.size = Pt(10)
                run.font.name = "Arial"
                run.font.color.rgb = RGBColor(31, 41, 55)

    # Guardar archivo en memoria y retornar bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# Barra lateral izquierda
    with st.sidebar:
    st.header("Configuración del Sistema")

    if not api_key:
        api_key = st.text_input("Clave API de Gemini", type="password")

# Selector de Rol del Usuario
        user_role = st.selectbox(
             "Perfil del Consultante:",
             ["Estudiante", "Acudiente / Padre de Familia", "Docente / Directivo"],
        )

        st.sidebar.markdown("---")
        st.sidebar.markdown("### Biblioteca Virtual")
        st.sidebar.markdown(
            "[Abrir Repositorio Digital](https://repositorioconvivencia-k3hz5bcgykhxqwn6gn7cjh.streamlit.app/)"
        )
    
    # Consulta de Marco Legal e Institucional con ENLACES DIRECTOS
    with st.expander("Ver Marco Legal e Institucional"):
        st.markdown(
            f"* **[Constitución Política]({URL_CONSTITUCION_POLITICA})**: Art."
            " 29 (Debido Proceso y Derechos Fundamentales)."
        )
        st.markdown(
            f"* **[Ley 115 de 1994]({URL_LEY_115})**: Ley General de"
            " Educación."
        )
        st.markdown(
            f"* **[Ley 1098 de 2006]({URL_LEY_1098})**: Código de la Infancia"
            " y la Adolescencia."
        )
        st.markdown(
            f"* **[Ley 1620 de 2013]({URL_LEY_1620})**: Sistema Nacional de"
            " Convivencia Escolar."
        )
        st.markdown(
            f"* **[Decreto 1965 de 2013]({URL_DECRETO_1965})**: Reglamentación"
            " de la Ley 1620."
        )
        st.markdown(
            f"* **[Manual de Convivencia]({URL_MANUAL_CONVIVENCIA})**: Manual"
            " Institucional de Convivencia."
        )

    st.markdown("---")
    st.markdown("**Guía de consulta:**")
    st.markdown("1. Ingrese los detalles de la situación acontecida.")
    st.markdown(
        "2. El sistema categorizará el hecho de acuerdo con el marco legal"
        " colombiano y el Manual de Convivencia."
    )
    st.markdown(
        "3. Se estructurará el procedimiento a seguir y la plantilla"
        " digital para la biblioteca o repositorio institucional."
    )

    if st.button("Reiniciar consulta"):
        st.session_state.messages = []
        st.rerun()

# Prompt de sistema institucional
SYSTEM_PROMPT = f"""
Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón de Soledad.

Estás orientando a un usuario con el perfil de: {user_role}.

Tu propósito es asesorar formal y pedagógicamente a la comunidad educativa ante situaciones disciplinarias, asegurando el cumplimiento de la Constitución Política de Colombia (Art. 29 - Debido Proceso), la Ley 115 de 1994, la Ley 1098 de 2006 (Código de Infancia y Adolescencia), la Ley 1620 de 2013, el Decreto 1965 de 2013 y el Manual de Convivencia de la Institución Educativa Técnica Sagrado Corazón.

Reglas interactivas OBLIGATORIAS de recolección de datos:
Para poder generar el reporte y el documento digital oficial completo, se requieren OBLIGATORIAMENTE los siguientes 7 datos:
1. Nombre completo del estudiante
2. Número de documento de identidad del estudiante
3. Grado y curso del estudiante
4. Asignatura / Clase en la que ocurrió el incidente(o si fue en la entrada a clases, recreo o salida)
5. Nombre completo del docente a cargo / reportante
6. Horario o Rango de horas en que sucedió
7. Fecha exacta de los hechos


EVALUACIÓN DEL HISTORIAL EN CADA TURNO:
- Antes de responder, analiza detalladamente TODO el historial de la conversación.
- SI FALTA UNO O MÁS DE LOS 7 DATOS MENCIONADOS: ESTÁ ESTRICTAMENTE PROHIBIDO generar los 5 puntos de la asesoría, el acta, la plantilla o cualquier documento formal. En su lugar, responde de forma amable, clara y formal indicando exactamente cuáles datos faltan y pidiéndoselos al usuario.
- SI EL USUARIO NO RESPONDE O DEJA CAMPOS INCOMPLETOS EN SU SIGUIENTE MENSAJE: Vuelve a preguntarle insistente pero respetuosamente por los datos faltantes. NO avances ni generes nada hasta tener los 7 datos completos.
- SOLO CUANDO TENGAS LOS 7 DATOS COMPLETOS EN EL HISTORIAL: Procede inmediatamente a generar la asesoría completa estructurada en 5 puntos, incluyendo al finalizar la consulta el modelo de documento digital con las líneas de firma (usando &#95;)
  
Instrucciones de formato para el documento digital (Una vez recolectados todos los datos):
- Este sistema es 100% digital para el archivo y repositorio institucional por año escolar. Queda ESTRICTAMENTE PROHIBIDO mencionar que el documento debe ser impreso, firmado en papel o presentado en físico.
- Rellena e integra en la redacción del documento TODOS los datos recolectados (Nombre, Documento, Grado, Asignatura, Docente, Horario, Fecha).
- ÚNICAMENTE deben quedar con líneas de subrayado (____________________) los espacios dedicados a las FIRMAS. El resto del texto debe quedar totalmente redactado con la información suministrada.
- NO incluyas meta-etiquetas ni subtítulos innecesarios dentro de la plantilla.
- El texto debe fluir de forma continua, limpia y profesional.
- No recargues el texto con negritas ni mayúsculas sostenidas. Mantén los párrafos en texto normal.
- Los compromisos o acuerdos se representan únicamente como una lista numerada secuencial (1., 2., 3.).

Estructura de la respuesta cuando los datos están completos (5 PUNTOS):
1. Resumen de la situación: Síntesis objetiva integrando los datos recolectados.
2. Clasificación de la falta (Según el Manual de Convivencia y Ley 1620 de 2013):
   - Situación Tipo I (Leve): Conflictos manejados inadecuadamente o faltas menores a los deberes.
   - Situación Tipo II (Grave): Acoso escolar (bullying), ciberacoso o agresiones físicas/verbales sin incapacidad médica.
   - Situación Tipo III (Gravísima): Presuntos delitos penales o agresiones físicas con incapacidad médica.
3. Procedimiento institucional: Protocolo a aplicar según el nivel de falta.
4. Garantías y Debido Proceso: Derechos aplicables protegidos por el Artículo 29 de la Constitución Política.
5. Modelo de Documento Digital Sugerido:
   - Si el perfil es Estudiante o Acudiente: Inicia con "ACTA DE COMPROMISO Y DESCARGOS ESTUDIANTILES" y redacta el modelo integrando todos los datos en la narración.
   - Si el perfil es Docente / Directivo: Inicia con "REGISTRO EN EL OBSERVADOR DE CONVIVENCIA ESCOLAR" incorporando al inicio los datos de la novedad y la descripción de los hechos con compromisos numerados.

Bloque final de firmas para el documento digital (ÚNICO LUGAR CON SUBRAYADOS):
  Lugar y fecha de diligenciamiento: Soledad, Atlántico, [Fecha suministrada]


  ____________________________________
  Firma del Estudiante
  Documento de Identidad N.° [Documento suministrado]


  ____________________________________
  Firma del Acudiente / Representante Legal
  Documento de Identidad N.° ____________________


  ____________________________________
  Firma del Docente Reportante / Coordinación
"""

WELCOME_MESSAGE = f"""
Saludos. Bienvenido(a) al Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón.

Sesión iniciada como: **{user_role}**.

Este portal brinda orientación sobre el protocolo disciplinario institucional, el marco legal colombiano y el debido proceso.

Por favor, seleccione una de las situaciones predeterminadas a continuación o redacte detalladamente lo sucedido en la casilla de texto inferior.
"""

# Inicialización del historial de chat
if "messages" not in st.session_state or len(st.session_state.messages) == 0:
    st.session_state.messages = [
        {"role": "assistant", "content": WELCOME_MESSAGE}
    ]

# Renderizar historial de mensajes y botón de descarga para cada respuesta del asistente
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

        # Generación del botón de descarga directa en Word únicamente cuando la respuesta contenga el documento oficial al finalizar
        if (
            msg["role"] == "assistant"
            and idx > 0
            and (
                "ACTA DE COMPROMISO" in msg["content"]
                or "REGISTRO EN EL OBSERVADOR" in msg["content"]
                or "MODELO DE DOCUMENTO" in msg["content"]
            )
        ):
            if HAS_DOCX:
                docx_bytes = generar_documento_word(msg["content"])
                st.download_button(
                    label=(
                        "📄 Descargar Documento Oficial en Microsoft Word (.docx)"
                    ),
                    data=docx_bytes,
                    file_name=(
                        f"Documento_Convivencia_SagradoCorazon_{idx}.docx"
                    ),
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                    key=f"dl_word_{idx}",
                )
            else:
                st.download_button(
                    label="📄 Guardar texto (.txt)",
                    data=msg["content"],
                    file_name=(
                        f"Documento_Convivencia_SagradoCorazon_{idx}.txt"
                    ),
                    mime="text/plain",
                    key=f"dl_txt_{idx}",
                )

# Opciones predefinidas rápidas al inicio de la conversación
selected_option = None
if len(st.session_state.messages) <= 1:
    st.markdown("**Seleccione el tipo de situación o escriba su caso abajo:**")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Exceso de maquillaje"):
            selected_option = "Se reporta un llamado de atención por exceso de maquillaje o incumplimiento del código de presentación personal."
        if st.button("Conflictos / Agresión verbal o física"):
            selected_option = "Ocurrió una situación de conflicto o agresión entre estudiantes dentro de la institución educativa."
        if st.button("Presunto Acoso Escolar (Bullying)"):
            selected_option = "Se presenta una situación reiterada de presunto acoso escolar (bullying) o ciberacoso."
        if st.button("Fraude académico / Plagio"):
            selected_option = "Se reporta una falta relacionada con fraude académico o plagio en evaluación durante la clase de la asignatura correspondiente."
        if st.button("Desacato o falta de respeto a docente"):
            selected_option = "Se presentó un acto de desobediencia o falta de respeto verbal hacia un docente durante el desarrollo de la clase."

    with col2:
        if st.button("Corte de cabello / Uniforme"):
            selected_option = "Se presenta un llamado de atención por corte de cabello inadecuado o porte incorrecto del uniforme institucional."
        if st.button("Incumplimiento de deberes / Asistencia"):
            selected_option = "Se presentó un incumplimiento en los deberes académicos, faltas de asistencia o impuntualidad."
        if st.button("Uso no autorizado de celular/equipos"):
            selected_option = "Se reporta el uso no autorizado de teléfono celular o dispositivos electrónicos durante la jornada escolar."
        if st.button("Evasión de clase / Ausencia en aula"):
            selected_option = "El estudiante ingresó a la institución pero evadió la clase o se ausentó del aula entre determinadas horas de la jornada escolar."
        if st.button("Daño a propiedad institucional"):
            selected_option = "Se reportan daños materiales a los pupitres, paredes u otros bienes de la institución."

# Captura de mensaje del usuario
user_input = st.chat_input("Escriba aquí los hechos de la situación a evaluar...")
prompt = selected_option or user_input

if prompt:
    if not api_key:
        st.warning("Se requiere ingresar la Clave API para continuar.")
        st.stop()

    st.session_state.messages.append({"role": "user", "content": prompt})

    try:
        client = genai.Client(api_key=api_key)

        contents = []
        for m in st.session_state.messages:
            role = "user" if m["role"] == "user" else "model"
            contents.append(
                types.Content(
                    role=role,
                    parts=[types.Part.from_text(text=m["content"])],
                )
            )

        with st.spinner("Procesando información institucional..."):
            response = client.models.generate_content(
                model="gemini-3.6-flash",
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT, temperature=0.2
                ),
            )

            if response and response.text:
                st.session_state.messages.append(
                    {"role": "assistant", "content": response.text}
                )
                st.rerun()

    except Exception as e:
        error_msg = str(e)
        if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
            st.warning(
                "⚠️ El servicio ha alcanzado el límite de consultas por minuto"
                " de la capa gratuita. Por favor, espere 30 a 40 segundos e"
                " intente de nuevo."
            )
        else:
            st.error(f"Error de comunicación con el servicio: {e}")
