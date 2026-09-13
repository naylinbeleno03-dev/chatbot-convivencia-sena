import hashlib
import io
import os
import re
import sqlite3
import time
from datetime import datetime
import pandas as pd
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

# CLAVE MAESTRA Y SEGURIDAD PARA DOCENTES / DIRECTIVOS
CLAVE_DIRECTIVA_CORRECTA = st.secrets.get("CLAVE_DOCENTE", "INTESAC2026")
CORREO_INSTITUCIONAL = "convivencia@intesac.edu.co"

# ==============================================================================
# BASE DE DATOS SQLITE - REPOSITORIO VIRTUAL SEGURO E INSTITUCIONAL
# ==============================================================================
DB_NAME = "repositorio_convivencia.db"


def init_db():
    """Inicializa la base de datos SQLite con soporte completo y validación de esquemas."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS registros_convivencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            estudiante_nombre TEXT NOT NULL,
            estudiante_documento TEXT NOT NULL,
            grado TEXT NOT NULL,
            jornada TEXT NOT NULL DEFAULT 'Mañana',
            asignatura TEXT,
            docente TEXT,
            horario TEXT,
            fecha_hechos TEXT NOT NULL,
            tipo_falta TEXT,
            an_lectivo INTEGER NOT NULL,
            contenido_texto TEXT NOT NULL,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


def guardar_registro_db(
    nombre,
    documento,
    grado,
    jornada,
    asignatura,
    docente,
    horario,
    fecha,
    tipo_falta,
    contenido,
):
    """Guarda un expediente oficial validado estrictamente por un directivo o sistema."""
    try:
        an_actual = datetime.now().year
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO registros_convivencia 
            (estudiante_nombre, estudiante_documento, grado, jornada, asignatura, docente, horario, fecha_hechos, tipo_falta, an_lectivo, contenido_texto)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                nombre,
                str(documento).strip(),
                grado,
                jornada,
                asignatura,
                docente,
                horario,
                fecha,
                tipo_falta,
                an_actual,
                contenido,
            ),
        )
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error al guardar en la base de datos: {e}")
        return False


def consultar_anios_db():
    """Obtiene los años lectivos disponibles en la base de datos."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT DISTINCT an_lectivo FROM registros_convivencia WHERE an_lectivo IS NOT NULL"
        )
        rows = cursor.fetchall()
        conn.close()
        return [str(r[0]) for r in rows if r[0]]
    except Exception:
        return [str(datetime.now().year)]


def consultar_registros_db(
    busqueda="", an=None, grado=None, jornada=None, doc_exacto=None
):
    """Consulta registros asegurando filtros controlados y adaptados a DataFrame."""
    conn = sqlite3.connect(DB_NAME)
    query = "SELECT id, estudiante_nombre, estudiante_documento, grado, jornada, asignatura, docente, fecha_hechos, tipo_falta, an_lectivo, contenido_texto, fecha_registro FROM registros_convivencia WHERE 1=1"
    params = []

    if doc_exacto:
        query += " AND estudiante_documento = ?"
        params.append(str(doc_exacto).strip())
    else:
        if busqueda:
            query += " AND (estudiante_nombre LIKE ? OR estudiante_documento LIKE ? OR docente LIKE ? OR tipo_falta LIKE ?)"
            like_term = f"%{busqueda}%"
            params.extend([like_term, like_term, like_term, like_term])

        if an and an != "Todos":
            query += " AND an_lectivo = ?"
            params.append(int(an))

        if grado and grado != "Todos":
            query += " AND grado = ?"
            params.append(grado)

        if jornada and jornada != "Todas":
            query += " AND jornada = ?"
            params.append(jornada)

    query += " ORDER BY fecha_registro DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


init_db()

# ==============================================================================
# ENLACES NORMATIVOS INSTITUCIONALES
# ==============================================================================
URL_CONSTITUCION_POLITICA = "https://www.registraduria.gov.co/IMG/pdf/constitucio-politica-colombia-1991.pdf"
URL_LEY_115 = (
    "https://www.mineducacion.gov.co/1621/articles-85906_archivo_pdf.pdf"
)
URL_LEY_1098 = (
    "https://www.icbf.gov.co/sites/default/files/codigoinfancialey1098.pdf"
)
URL_LEY_1620 = (
    "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=52287"
)
URL_DECRETO_1965 = (
    "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=54537"
)
URL_MANUAL_CONVIVENCIA = "https://drive.google.com/file/d/10WqGY5EvXzCMPROBZB6Ga6J4zjrEzmOG/view?usp=sharing"

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema Integral de Convivencia Escolar - INTESAC",
    layout="wide",
)

# Estilos CSS personalizados e institucionales avanzados
st.markdown(
    """
    <style>
    .stApp {
        background-color: #F3F4F6;
        color: #1F2937;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #1E293B;
        color: #F8FAFC;
        padding-top: 1rem;
        border-right: 1px solid #334155;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] a {
        color: #F8FAFC !important;
    }
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
    section[data-testid="stSidebar"] details a {
        color: #FDE047 !important;
        text-decoration: underline !important;
        font-weight: 600 !important;
    }
    .header-box {
        background-color: #0F172A;
        padding: 22px;
        border-radius: 12px;
        border-left: 6px solid #D4AF37;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2);
        margin-bottom: 20px;
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
    section[data-testid="stSidebar"] .stButton>button {
        background-color: #FFFFFF !important;
        color: #1F2937 !important;
        border: 1px solid #FFFFFF !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        width: 100%;
    }
    section[data-testid="stSidebar"] .stButton>button:hover {
        background-color: #E2E8F0 !important;
    }
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
    [data-testid="stChatMessage"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E5E7EB !important;
        border-left: 5px solid #0F172A !important;
        border-radius: 10px !important;
        padding: 16px !important;
        margin-bottom: 14px !important;
    }
    [data-testid="stChatMessage"] p, 
    [data-testid="stChatMessage"] div {
        color: #0F172A !important;
        font-size: 0.98rem;
        line-height: 1.6;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Encabezado superior institucional
st.markdown(
    """
    <div class="header-box">
        <h1>Sistema Integral de Convivencia Escolar</h1>
        <p>Institución Educativa Técnica Sagrado Corazón de Soledad (6° a 11°)</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Obtener clave API automáticamente si está guardada en Secrets o pedirla
api_key = st.secrets.get("GEMINI_API_KEY", "")


def limpiar_texto_para_word(texto: str) -> str:
    if not texto:
        return ""
    texto = re.sub(r"^[-\*_]{3,}\s*$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"^\s*[-\*]\s+", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"-{2,}", "", texto)
    return texto.strip()


def extraer_solo_documento(texto_contenido: str) -> str:
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


def generar_documento_word(texto_contenido):
    doc = Document()
    texto_documento = extraer_solo_documento(texto_contenido)

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

        header = section.header
        p_head = header.paragraphs[0]
        p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_head.paragraph_format.space_after = Pt(0)

        posibles_nombres_imagen = [
            "escudo.png",
            "escudo.jpg",
            "escudo_intesac.png",
            "logo.png",
        ]
        imagen_encontrada = next(
            (
                nombre
                for nombre in posibles_nombres_imagen
                if os.path.exists(nombre)
            ),
            None,
        )

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

        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_foot = p_foot.add_run(
            "Institución Educativa Técnica Sagrado Corazón — Archivo Digital de"
            " Convivencia Escolar"
        )
        r_foot.font.size = Pt(8.5)
        r_foot.font.name = "Arial"
        r_foot.font.italic = True
        r_foot.font.color.rgb = RGBColor(100, 116, 139)

    lineas = texto_documento.split("\n")
    for linea in lineas:
        linea_limpia = limpiar_texto_para_word(linea)
        if not linea_limpia:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15

        if (
            linea.strip().startswith("#")
            or (linea.strip().startswith("**") and linea.strip().endswith("**"))
            or "ACTA DE COMPROMISO" in linea.upper()
            or "MODELO DE CARTA" in linea.upper()
            or "CARTA DE DESCARGOS" in linea.upper()
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
        elif "________________" in linea or "FIRMA" in linea.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(linea_limpia.replace("**", ""))
            run.font.size = Pt(10)
            run.font.name = "Arial"
            run.font.bold = (
                "FIRMA" in linea.upper()
                or "ESTUDIANTE" in linea.upper()
                or "ACUDIENTE" in linea.upper()
                or "DOCENTE" in linea.upper()
            )
            run.font.color.rgb = RGBColor(31, 41, 55)
            if "________________" in linea:
                p.paragraph_format.space_before = Pt(36)
                p.paragraph_format.space_after = Pt(2)
            else:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(4)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
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

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# Barra lateral con control de acceso seguro y roles
with st.sidebar:
    st.header("Configuración del Sistema")

    if not api_key:
        api_key = st.text_input("Clave API de Gemini", type="password")

    user_role = st.selectbox(
        "Perfil del Consultante:",
        ["Estudiante", "Acudiente / Padre de Familia", "Docente / Directivo"],
    )

    autenticado_directivo = False
    if user_role == "Docente / Directivo":
        clave_ingresada = st.text_input(
            "Clave de Acceso Directivo:", type="password"
        )
        if clave_ingresada == CLAVE_DIRECTIVA_CORRECTA:
            autenticado_directivo = True
            st.success("Acceso Directivo Autorizado")
        elif clave_ingresada != "":
            st.error("Clave incorrecta")

    st.markdown("---")

    # Marco Legal e institucional
    with st.expander("Ver Marco Legal e Institucional"):
        st.markdown(
            f"* **[Constitución Política]({URL_CONSTITUCION_POLITICA})**: Art. 29"
            " (Debido Proceso)."
        )
        st.markdown(
            f"* **[Ley 115 de 1994]({URL_LEY_115})**: Ley General de Educación."
        )
        st.markdown(
            f"* **[Ley 1098 de 2006]({URL_LEY_1098})**: Código de Infancia y"
            " Adolescencia."
        )
        st.markdown(
            f"* **[Ley 1620 de 2013]({URL_LEY_1620})**: Convivencia Escolar."
        )
        st.markdown(
            f"* **[Decreto 1965 de 2013]({URL_DECRETO_1965})**: Reglamentación"
            " Ley 1620."
        )
        st.markdown(
            f"* **[Manual de Convivencia]({URL_MANUAL_CONVIVENCIA})**: Manual"
            " Institucional."
        )

    st.markdown("---")
    st.markdown("**Guía de consulta:**")
    st.markdown("1. Ingrese los detalles de la situación.")
    st.markdown("2. El sistema categorizará el hecho y el protocolo.")
    st.markdown("3. Se generará la plantilla oficial descargable en Word.")

    if st.button("Reiniciar consulta de chat"):
        st.session_state.messages = []
        st.rerun()

# Pestañas principales de navegación unificadas
tab_chat, tab_repositorio = st.tabs(
    ["Asistente y Generador de Casos", "Biblioteca y Repositorio Virtual"]
)

with tab_chat:
    SYSTEM_PROMPT = f"""
    Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón de Soledad.
    Estás orientando a un usuario con el perfil de: {user_role}.
    Tu propósito es asesorar formal y pedagógicamente ante situaciones disciplinarias, asegurando el cumplimiento estricto de la Constitución Política (Art. 29), Ley 115, Ley 1098, Ley 1620, Decreto 1965 y el Manual de Convivencia.

    REGLAS DE CONOCIMIENTO Y LENGUAJE:
    - Tú conoces a fondo el Manual de Convivencia institucional. NUNCA utilices expresiones condicionales, dudosas o evasivas como "si aplica", "si el Manual contempla", "dependiendo de lo que diga el manual". Afirma de forma directa, certera y categórica las normas, protocolos y medidas disciplinarias establecidas.
    - NO repitas innecesariamente el nombre de la institución ni utilices exceso de negritas en el texto o en el documento final. Mantén una redacción limpia, fluida y profesional.

    Reglas interactivas OBLIGATORIAS de recolección de datos (Se requieren los 7 datos completos):
    1. Nombre completo del estudiante
    2. Número de documento de identidad del estudiante
    3. Grado y curso del estudiante
    4. Asignatura / Clase en la que ocurrió el incidente
    5. Nombre completo del docente a cargo / reportante
    6. Horario o Rango de horas en que sucedió
    7. Fecha exacta de los hechos

    EVALUACIÓN DEL HISTORIAL EN CADA TURNO:
    - Si falta uno o más datos: NO generes los 5 puntos ni la plantilla. Pídeselos amablemente al usuario.
    - Si ya tienes los 7 datos completos: Procede inmediatamente a generar la asesoría completa en 5 puntos y la plantilla digital.

    Estructura de la respuesta (5 PUNTOS):
    1. Resumen de la situación integrando los datos recolectados.
    2. Clasificación de la falta (Tipo I Leve, Tipo II Grave, Tipo III Gravísima según Ley 1620).
    3. Procedimiento institucional.
    4. Garantías y Debido Proceso (Art. 29).
    5. Modelo de Documento Digital Sugerido (Iniciando con "ACTA DE COMPROMISO DEL ESTUDIANTE" o "REGISTRO EN EL OBSERVADOR DE CONVIVENCIA ESCOLAR"). 
        El bloque final de firmas debe ser limpio y estructurado exactamente así:

        Lugar y fecha de diligenciamiento: Soledad, Atlántico, [Fecha suministrada]


        ________________________________________          ________________________________________
        Firma del Estudiante                              Firma del Acudiente / Representante Legal
        Documento N.° [Documento suministrado]              Documento N.° ____________________


        ________________________________________
        Firma del Docente Reportante / Coordinación
    """

    WELCOME_MESSAGE = f"""
    Saludos. Bienvenido(a) al Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón.
    Sesión iniciada como: **{user_role}**.
    Por favor, seleccione una de las situaciones rápidas a continuación o escriba los hechos en la barra inferior.
    """

    if "messages" not in st.session_state or len(st.session_state.messages) == 0:
        st.session_state.messages = [
            {"role": "assistant", "content": WELCOME_MESSAGE}
        ]

    # Renderizar historial de chat
    for idx, msg in enumerate(st.session_state.messages):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if msg["role"] == "assistant" and idx > 0:
                texto_msg = msg["content"].upper()
                if (
                    "5." in texto_msg
                    or "ACTA" in texto_msg
                    or "OBSERVADOR" in texto_msg
                ):
                    content_hash = hashlib.md5(
                        msg["content"].encode("utf-8")
                    ).hexdigest()[:8]
                    if HAS_DOCX:
                        docx_bytes = generar_documento_word(msg["content"])
                        st.download_button(
                            label=(
                                "Descargar Documento Oficial en Microsoft Word"
                                " (.docx)"
                            ),
                            data=docx_bytes,
                            file_name=(
                                f"Documento_Convivencia_SagradoCorazon_{idx}.docx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                            key=f"dl_word_{idx}_{content_hash}",
                        )
                    else:
                        st.download_button(
                            label="Guardar texto (.txt)",
                            data=msg["content"],
                            file_name=(
                                f"Documento_Convivencia_SagradoCorazon_{idx}.txt"
                            ),
                            mime="text/plain",
                            key=f"dl_txt_{idx}_{content_hash}",
                        )

    # Botones rápidos de selección frecuente
    selected_option = None
    if len(st.session_state.messages) <= 1:
        st.markdown("**Seleccione el tipo de situación frecuente:**")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Exceso de maquillaje"):
                selected_option = (
                    "Se reporta un llamado de atención por exceso de maquillaje"
                    " o incumplimiento del código de presentación personal."
                )
            if st.button("Conflictos / Agresión verbal o física"):
                selected_option = (
                    "Ocurrió una situación de conflicto o agresión entre"
                    " estudiantes dentro de la institución educativa."
                )
            if st.button("Presunto Acoso Escolar (Bullying)"):
                selected_option = (
                    "Se presenta una situación reiterada de presunto acoso"
                    " escolar (bullying) o ciberacoso."
                )
            if st.button("Fraude académico / Plagio"):
                selected_option = (
                    "Se reporta una falta relacionada con fraude académico o"
                    " plagio en evaluación durante la clase."
                )
            if st.button("Desacato o falta de respeto a docente"):
                selected_option = (
                    "Se presentó un acto de desobediencia o falta de respeto"
                    " verbal hacia un docente durante el desarrollo de la"
                    " clase."
                )
        with col2:
            if st.button("Corte de cabello / Uniforme"):
                selected_option = (
                    "Se presenta un llamado de atención por corte de cabello"
                    " inadecuado o porte incorrecto del uniforme"
                    " institucional."
                )
            if st.button("Incumplimiento de deberes / Asistencia"):
                selected_option = (
                    "Se presentó un incumplimiento en los deberes académicos,"
                    " faltas de asistencia o impuntualidad."
                )
            if st.button("Uso no autorizado de celular/equipos"):
                selected_option = (
                    "Se reporta el uso no autorizado de teléfono celular o"
                    " dispositivos electrónicos durante la jornada escolar."
                )
            if st.button("Evasión de clase / Ausencia en aula"):
                selected_option = (
                    "El estudiante ingresó a la institución pero evadió la"
                    " clase o se ausentó del aula entre determinadas horas."
                )
            if st.button("Daño a propiedad institucional"):
                selected_option = (
                    "Se reportan daños materiales a los pupitres, paredes u"
                    " otros bienes de la institución."
                )

    user_input = st.chat_input(
        "Escriba aquí los hechos de la situación a evaluar..."
    )
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
                    model="gemini-2.5-flash",
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT, temperature=0.2
                    ),
                )

                if response and response.text:
                    respuesta_texto = response.text
                    st.session_state.messages.append(
                        {"role": "assistant", "content": respuesta_texto}
                    )

                    # Guardado automático al repositorio si el modelo genera un acta o registro válido
                    if (
                        "ACTA" in respuesta_texto.upper()
                        or "OBSERVADOR" in respuesta_texto.upper()
                    ):
                        guardar_registro_db(
                            nombre="Estudiante Registrado",
                            documento="N/A",
                            grado="General",
                            jornada="Mañana",
                            asignatura="General",
                            docente="Docente Institucional",
                            horario="Jornada Escolar",
                            fecha=datetime.now().strftime("%Y-%m-%d"),
                            tipo_falta="Evaluado por IA",
                            contenido=respuesta_texto,
                        )

                    st.rerun()

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                st.warning(
                    "El servicio ha alcanzado el límite de consultas por"
                    " minuto. Espere unos segundos e intente de nuevo."
                )
            else:
                st.error(f"Error de comunicación con el servicio: {e}")

with tab_repositorio:
    st.subheader("Biblioteca y Repositorio Virtual de Convivencia Escolar")
    st.markdown(
        "Consulte el archivo histórico de actas y registros institucionales"
        " organizados por **Año**, **Grado**, **Jornada** y términos de"
        " búsqueda."
    )

    if user_role == "Docente / Directivo":
        if autenticado_directivo:
            st.info("Módulo de Publicación y Gestión Directiva Autorizado.")
            with st.expander(
                "Publicar Documento u Acta Oficial en la Biblioteca"
            ):
                with st.form("form_registro_directivo"):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        f_nombre = st.text_input(
                            "Nombre Completo del Estudiante"
                        )
                        f_doc = st.text_input(
                            "Tarjeta de Identidad / Documento"
                        )
                        f_grado = st.selectbox(
                            "Grado", ["6°", "7°", "8°", "9°", "10°", "11°"]
                        )
                        f_jornada = st.selectbox("Jornada", ["Mañana", "Tarde"])
                    with col_b:
                        f_asignatura = st.text_input(
                            "Asignatura / Clase", "Convivencia"
                        )
                        f_docente = st.text_input("Docente / Evaluador")
                        f_fecha = st.date_input("Fecha de los Hechos")
                        f_falta = st.selectbox(
                            "Clasificación de Falta",
                            [
                                "Situación Tipo I (Leve)",
                                "Situación Tipo II (Grave)",
                                "Situación Tipo III (Gravísima)",
                                "Carta de Descargos / Trámite",
                            ],
                        )

                    f_texto = st.text_area("Contenido definitivo del Documento")

                    if st.form_submit_button(
                        "Publicar Oficialmente en Biblioteca"
                    ):
                        if f_nombre and f_doc and f_texto:
                            guardar_registro_db(
                                nombre=f_nombre,
                                documento=f_doc,
                                grado=f_grado,
                                jornada=f_jornada,
                                asignatura=f_asignatura,
                                docente=f_docente,
                                horario="Jornada Escolar",
                                fecha=str(f_fecha),
                                tipo_falta=f_falta,
                                contenido=f_texto,
                            )
                            st.success(
                                f"Documento oficial de {f_nombre} publicado con"
                                " éxito."
                            )
                            st.rerun()
        else:
            st.warning(
                "Acceso Restringido: Ingrese la clave correcta en la barra"
                " lateral para habilitar el módulo directivo."
            )

    st.markdown("---")

    anios_disponibles = consultar_anios_db()
    if not anios_disponibles:
        anios_disponibles = [str(datetime.now().year)]
    lista_anios_select = ["Todos"] + sorted(
        list(set(anios_disponibles)), reverse=True
    )

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        filtro_a = st.selectbox("Filtrar por Año:", lista_anios_select)
    with col_f2:
        filtro_g = st.selectbox(
            "Filtrar por Grado:",
            ["Todos", "6°", "7°", "8°", "9°", "10°", "11°"],
        )
    with col_f3:
        filtro_j = st.selectbox(
            "Filtrar por Jornada:", ["Todas", "Mañana", "Tarde"]
        )
    with col_f4:
        filtro_b = st.text_input(
            "Buscar (Estudiante / Documento / Docente):", ""
        )

    # Consulta unificada compatible con ambos flujos de datos
    df_registros = consultar_registros_db(
        busqueda=filtro_b,
        an=filtro_a if filtro_a != "Todos" else None,
        grado=filtro_g if filtro_g != "Todos" else None,
        jornada=filtro_j if filtro_j != "Todas" else None,
    )

    if not df_registros.empty:
        st.info(
            f"Se encontraron {len(df_registros)} registro(s) en el repositorio."
        )
        for index, row in df_registros.iterrows():
            with st.expander(
                f"Grado {row['grado']} ({row['jornada']}) —"
                f" {row['estudiante_nombre']} ({row['estudiante_documento']})"
                f" [Fecha: {row['fecha_hechos']}]"
            ):
                st.write(
                    f"**Docente:** {row['docente']} | **Asignatura:**"
                    f" {row['asignatura']}"
                )
                st.markdown("---")
                st.markdown(row["contenido_texto"])
                if HAS_DOCX:
                    doc_bytes = generar_documento_word(row["contenido_texto"])
                    st.download_button(
                        label="Descargar Documento Oficial en Word (.docx)",
                        data=doc_bytes,
                        file_name=(
                            f"Acta_{str(row['estudiante_documento']).strip()}.docx"
                        ),
                        key=f"rep_dl_{row['id']}_{index}",
                    )
    else:
        st.warning(
            "No hay registros guardados en la base de datos para los filtros de"
            " año, grado, jornada o búsqueda seleccionados."
        )
