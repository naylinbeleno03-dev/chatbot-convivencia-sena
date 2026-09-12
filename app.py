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

# CLAVE MAESTRA PARA DOCENTES / DIRECTIVOS
CLAVE_DIRECTIVA_CORRECTA = st.secrets.get("CLAVE_DOCENTE", "1ntesacSOLEDADgrupo1")

# CORREO INSTITUCIONAL DE RECEPCIÓN DE DOCUMENTOS
CORREO_INSTITUCIONAL = "naylinbeleno03@gmail.com"

# ==============================================================================
# BASE DE DATOS SQLITE - REPOSITORIO VIRTUAL
# ==============================================================================
DB_NAME = "repositorio_convivencia.db"


def init_db():
    """Inicializa la base de datos SQLite."""
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
            ano_lectivo INTEGER NOT NULL,
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
    """Guarda un expediente oficial validado por un directivo."""
    try:
        ano_actual = datetime.now().year
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


def consultar_registros_db(
    busqueda="", ano=None, grado=None, jornada=None, doc_exacto=None
):
    """Consulta registros en la biblioteca virtual."""
    conn = sqlite3.connect(DB_NAME)
    query = "SELECT id, estudiante_nombre, estudiante_documento, grado, jornada, asignatura, docente, fecha_hechos, tipo_falta, an_lectivo, contenido_texto, fecha_registro FROM registros_convivencia WHERE 1=1"
    params = []

    if doc_exacto:
        query += " AND estudiante_documento = ?"
        params.append(str(doc_exacto).strip())
    else:
        if busqueda:
            query += " AND (estudiante_nombre LIKE ? OR estudiante_documento LIKE ? OR docente LIKE ?)"
            params.extend([f"%{busqueda}%", f"%{busqueda}%", f"%{busqueda}%"])

        if an and an != "Todos":
            query += " AND an_lectivo = ?"
            params.append(int(ano))

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
URL_LEY_1620 = "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=52287"
URL_MANUAL_CONVIVENCIA = "https://drive.google.com/file/d/10WqGY5EvXzCMPROBZB6Ga6J4zjrEzmOG/view?usp=sharing"

st.set_page_config(
    page_title="Sistema Integral de Convivencia Escolar - INTESAC",
    layout="wide",
)

# Estilos CSS
st.markdown(
    """
    <style>
    .stApp { background-color: #F3F4F6; color: #1F2937; font-family: 'Segoe UI', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #475569; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] a { color: #FFFFFF !important; }
    .header-box { background-color: #0F172A; padding: 22px; border-radius: 12px; border-left: 6px solid #D4AF37; margin-bottom: 20px; }
    .header-box h1 { color: #FFFFFF !important; margin: 0 !important; font-size: 1.65rem !important; }
    .header-box p { color: #E2E8F0 !important; margin-top: 6px !important; }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="header-box">
        <h1>Sistema Integral de Convivencia Escolar</h1>
        <p>Institución Educativa Técnica Sagrado Corazón de Soledad (6° a 11°)</p>
    </div>
""",
    unsafe_allow_html=True,
)

api_key = st.secrets.get("GEMINI_API_KEY", "")


def limpiar_texto_para_word(texto: str) -> str:
    if not texto:
        return ""
    texto = re.sub(r"^[-\*_]{3,}\s*$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"^\s*[-\*]\s+", "", texto, flags=re.MULTILINE)
    return texto.strip()


def extraer_solo_documento(texto_contenido: str) -> str:
    if not texto_contenido:
        return ""
    patrones = [
        r"(ACTA DE COMPROMISO.*)",
        r"(MODELO DE CARTA.*)",
        r"(REGISTRO EN EL OBSERVADOR.*)",
    ]
    for patron in patrones:
        match = re.search(patron, texto_contenido, re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1).strip()
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

        posibles_nombres_imagen = ["escudo.png", "escudo.jpg", "logo.png"]
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
        r_head.font.bold = True

    lineas = texto_documento.split("\n")
    for linea in lineas:
        linea_limpia = limpiar_texto_para_word(linea)
        if not linea_limpia:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        if "ACTA DE COMPROMISO" in linea.upper() or "REGISTRO" in linea.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(linea_limpia)
            run.bold = True
            run.font.size = Pt(12)
        elif "________________" in linea or "FIRMA" in linea.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(linea_limpia)
            run.font.size = Pt(10)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = p.add_run(linea_limpia)
            run.font.size = Pt(10)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# BARRA LATERAL CON CONTROL DE ACCESO
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
            "🔒 Clave de Acceso Directivo:", type="password"
        )
        if clave_ingresada == CLAVE_DIRECTIVA_CORRECTA:
            autenticado_directivo = True
            st.success("Acceso Directivo Autorizado")
        elif clave_ingresada != "":
            st.error("Clave incorrecta")

    st.markdown("---")
    with st.expander("Ver Marco Legal e Institucional"):
        st.markdown(f"* [Constitución Política]({URL_CONSTITUCION_POLITICA})")
        st.markdown(f"* [Ley 115 de 1994]({URL_LEY_115})")
        st.markdown(f"* [Ley 1098 de 2006]({URL_LEY_1098})")
        st.markdown(f"* [Ley 1620 de 2013]({URL_LEY_1620})")
        st.markdown(f"* [Manual de Convivencia]({URL_MANUAL_CONVIVENCIA})")

    st.markdown("---")
    if st.button("Reiniciar consulta"):
        st.session_state.messages = []
        st.rerun()

SYSTEM_PROMPT = f"""
Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón de Soledad (Grados 6° a 11°).
Perfil actual: {user_role}.

Tu propósito es orientar al usuario en la redacción de sus actas y descargos.
IMPORTANTE: Aclara al usuario que las actas generadas en este chat son borradores orientativos que deben enviarse al correo institucional ({CORREO_INSTITUCIONAL}) para ser analizadas y posteriormente publicadas por un directivo en la biblioteca oficial.

DATOS OBLIGATORIOS REQUERIDOS (6° a 11°):
1. Nombre completo del estudiante
2. Tarjeta de identidad / Documento del estudiante
3. Grado (6° a 11°) y Jornada (Mañana o Tarde)
4. Asignatura / Clase
5. Nombre del docente a cargo / reportante
6. Horario y Fecha exacta de los hechos

Si faltan datos, solicítalos amablemente. Cuando estén completos, genera el reporte con el documento digital formal al final.
"""

tab_chat, tab_repositorio = st.tabs(
    ["💬 Asistente y Generador de Actas", "📚 Biblioteca / Repositorio Virtual"]
)

# PESTAÑA 1: CHATBOT (SOLO ASISTENTE Y GENERADOR DE BORRADORES)
with tab_chat:
    WELCOME_MESSAGE = f"Saludos. Bienvenido(a) al Sistema Integral de Convivencia Escolar de INTESAC.\n\nSesión iniciada como: **{user_role}**.\n\n*Nota: Este chat orienta y genera su borrador de acta. Para su registro oficial, descargue el archivo Word y envíelo al correo **{CORREO_INSTITUCIONAL}** para validación directiva.*"

    if "messages" not in st.session_state or len(st.session_state.messages) == 0:
        st.session_state.messages = [
            {"role": "assistant", "content": WELCOME_MESSAGE}
        ]

    for idx, msg in enumerate(st.session_state.messages):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if (
                msg["role"] == "assistant"
                and idx > 0
                and HAS_DOCX
                and "5." in msg["content"]
            ):
                docx_bytes = generar_documento_word(msg["content"])
                st.download_button(
                    label="📄 Descargar Borrador de Acta en Word (.docx)",
                    data=docx_bytes,
                    file_name=f"Borrador_Acta_INTESAC_{idx}.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    ),
                    key=f"dl_word_{idx}",
                )
                st.info(
                    f"📩 Una vez completado, envíe este documento al correo"
                    f" **{CORREO_INSTITUCIONAL}** para revisión por parte de"
                    " Coordinación."
                )

    user_input = st.chat_input("Escriba aquí la situación o los datos...")
    if user_input:
        if not api_key:
            st.warning("Ingrese la Clave API para continuar.")
            st.stop()

        st.session_state.messages.append({"role": "user", "content": user_input})

        try:
            client = genai.Client(api_key=api_key)
            contents = [
                types.Content(
                    role="user" if m["role"] == "user" else "model",
                    parts=[types.Part.from_text(text=m["content"])],
                )
                for m in st.session_state.messages
            ]

            with st.spinner("Procesando información..."):
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
            st.error(f"Error en la consulta: {e}")

# PESTAÑA 2: REPOSITORIO VIRTUAL (PUBLICACIÓN EXCLUSIVA DE DIRECTIVOS)
with tab_repositorio:
    st.subheader("📚 Repositorio Digital de Seguimiento Disciplinario")

    # PERMISO DOCENTE / DIRECTIVO (PUBLICACIÓN Y GESTIÓN OFICIAL)
    if user_role == "Docente / Directivo":
        if autenticado_directivo:
            st.info(
                "🔓 **Modulo de Publicación Directiva:** Revisa, aprueba y"
                " publica actas oficiales en la biblioteca institucional."
            )

            # Formulario exclusivo de publicación oficial
            with st.expander(
                "➕ Publicar Acta Oficial Analizada en la Biblioteca"
            ):
                with st.form("form_registro"):
                    col_a, col_b = st.columns(2)
                    with col_a:
                        f_nombre = st.text_input(
                            "Nombre Completo del Estudiante"
                        )
                        f_doc = st.text_input(
                            "Tarjeta de Identidad / Documento"
                        )
                        f_grado = st.selectbox(
                            "Grado (6° a 11°)",
                            ["6°", "7°", "8°", "9°", "10°", "11°"],
                        )
                        f_jornada = st.selectbox(
                            "Jornada Escolar", ["Mañana", "Tarde"]
                        )
                    with col_b:
                        f_asignatura = st.text_input(
                            "Asignatura / Clase", "Convivencia"
                        )
                        f_docente = st.text_input(
                            "Docente / Coordinador Evaluador"
                        )
                        f_fecha = st.date_input("Fecha del Incidente")
                        f_falta = st.selectbox(
                            "Clasificación de Falta Definitiva",
                            [
                                "Situación Tipo I (Leve)",
                                "Situación Tipo II (Grave)",
                                "Situación Tipo III (Gravísima)",
                            ],
                        )

                    f_texto = st.text_area(
                        "Contenido definitivo del Acta / Resolución"
                    )

                    if st.form_submit_button(
                        "📌 Publicar Oficialmente en Biblioteca"
                    ):
                        if f_nombre and f_doc and f_texto:
                            guardar_registro_db(
                                f_nombre,
                                f_doc,
                                f_grado,
                                f_jornada,
                                f_asignatura,
                                f_docente,
                                "Jornada Escolar",
                                str(f_fecha),
                                f_falta,
                                f_texto,
                            )
                            st.success(
                                f"✅ Acta de {f_nombre} publicada"
                                " exitosamente en la biblioteca digital."
                            )
                            st.rerun()

            st.markdown("---")

            # Filtros de Búsqueda
            col_b1, col_b2, col_b3, col_b4 = st.columns([2, 1, 1, 1])
            with col_b1:
                search_txt = st.text_input(
                    "🔍 Buscar por Estudiante, Documento o Docente:"
                )
            with col_b2:
                f_ano = st.selectbox(
                    "Año Lectivo:", ["Todos", "2026", "2025", "2024"], index=1
                )
            with col_b3:
                f_grado_sel = st.selectbox(
                    "Grado:", ["Todos", "6°", "7°", "8°", "9°", "10°", "11°"]
                )
            with col_b4:
                f_jornada_sel = st.selectbox(
                    "Jornada:", ["Todas", "Mañana", "Tarde"]
                )

            df_registros = consultar_registros_db(
                busqueda=search_txt,
                ano=f_ano,
                grado=f_grado_sel,
                jornada=f_jornada_sel,
            )

            if not df_registros.empty:
                st.markdown(
                    f"**Se encontraron {len(df_registros)} acta(s) publicada(s):**"
                )
                for index, row in df_registros.iterrows():
                    with st.expander(
                        f"📄 [{row['ano_lectivo']}] Grado {row['grado']} ("
                        f"Jornada {row['jornada']}) — {row['estudiante_nombre']}"
                        f" ({row['estudiante_documento']})"
                    ):
                        st.write(f"**Evaluador/Docente:** {row['docente']}")
                        st.write(
                            f"**Asignatura:** {row['asignatura']} | **Fecha:**"
                            f" {row['fecha_hechos']}"
                        )
                        st.markdown(row["contenido_texto"])
                        if HAS_DOCX:
                            doc_bytes = generar_documento_word(
                                row["contenido_texto"]
                            )
                            st.download_button(
                                label="📄 Descargar Acta Word (.docx)",
                                data=doc_bytes,
                                file_name=(
                                    f"Acta_Oficial_{row['estudiante_documento']}.docx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            ),
                                key=f"dl_db_{row['id']}",
                            )
            else:
                st.info(
                    "No hay actas publicadas que coincidan con los filtros."
                )
        else:
            st.warning(
                "🔒 **Acceso Restringido:** Ingrese la clave de acceso"
                " directivo en la barra lateral para gestionar la biblioteca"
                " oficial."
            )

    # PERMISO ESTUDIANTE / ACUDIENTE (CONSULTA DE ACTAS PUBLICADAS)
    else:
        st.info(
            "🔎 **Consulta de Estado de Convivencia:** Ingrese su Tarjeta de"
            " Identidad para consultar si la coordinación ha publicado un acta"
            " oficial correspondiente a su caso."
        )

        doc_estudiante_input = st.text_input(
            "Número de Tarjeta de Identidad del Estudiante:",
            key="doc_est_search",
        )

        if doc_estudiante_input:
            df_mis_registros = consultar_registros_db(
                doc_exacto=doc_estudiante_input
            )

            if not df_mis_registros.empty:
                st.warning(
                    f"⚠️ Se encontraron {len(df_mis_registros)} acta(s) oficiales"
                    " publicadas en el sistema:"
                )

                for index, row in df_mis_registros.iterrows():
                    with st.expander(
                        f"📄 Acta Oficial [{row['ano_lectivo']}] — Grado"
                        f" {row['grado']} ({row['jornada']}) | Fecha:"
                        f" {row['fecha_hechos']}"
                    ):
                        st.write(f"**Estudiante:** {row['estudiante_nombre']}")
                        st.write(f"**Asignatura:** {row['asignatura']}")
                        st.write(f"**Docente/Evaluador:** {row['docente']}")
                        st.markdown("---")
                        st.markdown(row["contenido_texto"])

                        if HAS_DOCX:
                            doc_bytes = generar_documento_word(
                                row["contenido_texto"]
                            )
                            st.download_button(
                                label=(
                                    "📄 Descargar Acta Oficial en Word (.docx)"
                                ),
                                data=doc_bytes,
                                file_name=(
                                    f"Acta_Oficial_{row['estudiante_documento']}.docx"
                                ),
                                mime=(
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                ),
                                key=f"dl_my_db_{row['id']}",
                            )
            else:
                st.success(
                    f"✅ **Sin registros pendientes:** No se encontraron actas"
                    " oficiales publicadas asociadas a la Tarjeta de Identidad"
                    f" N.° **{doc_estudiante_input}**. Tu historial de"
                    " convivencia escolar se encuentra totalmente limpio."
                )
