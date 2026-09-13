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

# ==============================================================================
# CONFIGURACIÓN Y CONSTANTES
# ==============================================================================
CLAVE_DIRECTIVA_CORRECTA = st.secrets.get("CLAVE_DOCENTE", "INTESAC2026")
CORREO_INSTITUCIONAL = "convivencia@intesac.edu.co"
DB_NAME = "repositorio_convivencia.db"

URL_CONSTITUCION_POLITICA = "https://www.registraduria.gov.co/IMG/pdf/constitucio-politica-colombia-1991.pdf"
URL_LEY_115 = "https://www.mineducacion.gov.co/1621/articles-85906_archivo_pdf.pdf"
URL_LEY_1098 = "https://www.icbf.gov.co/sites/default/files/codigoinfancialey1098.pdf"
URL_LEY_1620 = "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=52287"
URL_DECRETO_1965 = "https://www.funcionpublica.gov.co/eva/gestornormativo/norma_pdf.php?i=54537"
URL_MANUAL_CONVIVENCIA = "https://drive.google.com/file/d/10WqGY5EvXzCMPROBZB6Ga6J4zjrEzmOG/view?usp=sharing"

st.set_page_config(
    page_title="Sistema Integral de Convivencia Escolar - INTESAC",
    layout="wide",
)

# ==============================================================================
# GESTIÓN DE BASE DE DATOS SQLITE (REPOSITORIO)
# ==============================================================================
def init_db():
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

def guardar_registro_db(nombre, documento, grado, jornada, asignatura, docente, horario, fecha, tipo_falta, contenido):
    try:
        an_actual = datetime.now().year
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO registros_convivencia 
            (estudiante_nombre, estudiante_documento, grado, jornada, asignatura, docente, horario, fecha_hechos, tipo_falta, an_lectivo, contenido_texto)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (nombre, str(documento).strip(), grado, jornada, asignatura, docente, horario, fecha, tipo_falta, an_actual, contenido))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error al guardar en la base de datos: {e}")
        return False

def consultar_registros_db(busqueda="", an=None, grado=None, jornada=None, doc_exacto=None):
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
# ESTILOS CSS PERSONALIZADOS
# ==============================================================================
st.markdown("""
    <style>
    .stApp { background-color: #F3F4F6; color: #1F2937; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    section[data-testid="stSidebar"] { background-color: #475569; border-right: 1px solid #334155; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] li, section[data-testid="stSidebar"] a { color: #FFFFFF !important; }
    .header-box { background-color: #0F172A; padding: 22px; border-radius: 12px; border-left: 6px solid #D4AF37; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2); margin-bottom: 24px; }
    .header-box h1 { color: #FFFFFF !important; margin: 0 !important; font-size: 1.65rem !important; }
    .header-box p { color: #E2E8F0 !important; margin-top: 6px !important; margin-bottom: 0 !important; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="header-box">
        <h1>Sistema Integral de Convivencia Escolar</h1>
        <p>Institución Educativa Técnica Sagrado Corazón</p>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# BARRA LATERAL Y NAVEGACIÓN PRINCIPAL
# ==============================================================================
api_key = st.secrets.get("GEMINI_API_KEY", "")

with st.sidebar:
    st.header("Configuración del Sistema")
    if not api_key:
        api_key = st.text_input("Clave API de Gemini", type="password")

    user_role = st.selectbox("Perfil del Consultante:", ["Estudiante", "Acudiente / Padre de Familia", "Docente / Directivo"])
    
    st.markdown("---")
    st.markdown("### 🧭 Menú de Navegación")
    seccion_actual = st.radio("Ir a:", ["🤖 Asistente de Convivencia", "📚 Repositorio y Archivo Histórico"])

    st.markdown("---")
    with st.expander("Ver Marco Legal e Institucional"):
        st.markdown(f"* **[Constitución Política]({URL_CONSTITUCION_POLITICA})**: Art. 29.")
        st.markdown(f"* **[Ley 115]({URL_LEY_115})**: Ley General de Educación.")
        st.markdown(f"* **[Ley 1098]({URL_LEY_1098})**: Código de Infancia.")
        st.markdown(f"* **[Ley 1620]({URL_LEY_1620})**: Convivencia Escolar.")
        st.markdown(f"* **[Decreto 1965]({URL_DECRETO_1965})**: Reglamentación Ley 1620.")
        st.markdown(f"* **[Manual de Convivencia]({URL_MANUAL_CONVIVENCIA})**: Manual Institucional.")

# ==============================================================================
# SECCIÓN 1: ASISTENTE DE CONVIVENCIA (CHAT)
# ==============================================================================
if seccion_actual == "🤖 Asistente de Convivencia":
    SYSTEM_PROMPT = f"""
    Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón de Soledad.
    Perfil del usuario actual: {user_role}.
    Asegura el cumplimiento del debido proceso (Art. 29 Constitución Política), Ley 115, Ley 1098, Ley 1620 y el Manual de Convivencia.
    Datos requeridos obligatoriamente: 1. Nombre completo, 2. Documento de identidad, 3. Grado/curso, 4. Asignatura, 5. Docente a cargo, 6. Horario, 7. Fecha exacta.
    Si faltan datos, solicítalos amablemente. Si están completos, genera la asesoría estructurada en 5 puntos y la plantilla formal.
    """

    WELCOME_MESSAGE = f"""
    Bienvenido(a) al Asistente Digital de Convivencia Escolar. Sesión como: **{user_role}**.
    Seleccione una situación rápida abajo o escriba su caso en el chat.
    """

    if "messages" not in st.session_state or len(st.session_state.messages) == 0:
        st.session_state.messages = [{"role": "assistant", "content": WELCOME_MESSAGE}]

    for idx, msg in enumerate(st.session_state.messages):
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if msg["role"] == "assistant" and idx > 0 and HAS_DOCX:
                # Opcional: Lógica de descarga de Word
                pass

    user_input = st.chat_input("Escriba los detalles de la situación...")
    if user_input:
        if not api_key:
            st.warning("Se requiere Clave API.")
            st.stop()
        st.session_state.messages.append({"role": "user", "content": user_input})
        try:
            client = genai.Client(api_key=api_key)
            contents = [types.Content(role="user" if m["role"] == "user" else "model", parts=[types.Part.from_text(text=m["content"])]) for m in st.session_state.messages]
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
                config=types.GenerateContentConfig(system_instruction=SYSTEM_PROMPT, temperature=0.2),
            )
            if response and response.text:
                st.session_state.messages.append({"role": "assistant", "content": response.text})
                st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")

# ==============================================================================
# SECCIÓN 2: REPOSITORIO Y BIBLIOTECA VIRTUAL (CON AÑOS LECTIVOS)
# ==============================================================================
elif seccion_actual == "📚 Repositorio y Archivo Histórico":
    st.subheader("📁 Repositorio Digital y Expedientes de Convivencia")
    st.markdown("Consulte los registros históricos de seguimiento por **año lectivo**, grado o documento del estudiante.")

    # Filtros de búsqueda para el repositorio
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        filtro_texto = st.text_input("Buscar por nombre, documento o docente:", "")
    with col_f2:
        # Extraer años disponibles en la BD
        conn_db = sqlite3.connect(DB_NAME)
        anos_disponibles = pd.read_sql_query("SELECT DISTINCT an_lectivo FROM registros_convivencia ORDER BY an_lectivo DESC", conn_db)
        conn_db.close()
        lista_anos = ["Todos"] + [str(an) for an in anos_disponibles["an_lectivo"].tolist()] if not anos_disponibles.empty else ["Todos", str(datetime.now().year)]
        filtro_ano = st.selectbox("Año Lectivo:", lista_anos)
    with col_f3:
        filtro_grado = st.selectbox("Grado:", ["Todos", "6°", "7°", "8°", "9°", "10°", "11°"])

    # Consulta a la base de datos con los filtros aplicados
    df_registros = consultar_registros_db(
        busqueda=filtro_texto, 
        an=filtro_ano if filtro_ano != "Todos" else None, 
        grado=filtro_grado if filtro_grado != "Todos" else None
    )

    if not df_registros.empty:
        st.markdown(f"**Registros encontrados:** {len(df_registros)}")
        for index, row in df_registros.iterrows():
            with st.expander(f"📌 {row['estudiante_nombre']} (Doc: {row['estudiante_documento']}) - Grado: {row['grado']} | Año: {row['an_lectivo']} | Fecha: {row['fecha_hechos']}"):
                st.write(f"**Docente Reportante:** {row['docente']}")
                st.write(f"**Asignatura / Jornada:** {row['asignatura']} ({row['jornada']})")
                st.write(f"**Tipo de Falta:** {row['tipo_falta']}")
                st.markdown("---")
                st.text_area("Contenido del Acta / Expediente:", row['contenido_texto'], height=150, key=f"txt_repo_{row['id']}")
    else:
        st.info("No se encontraron registros en el repositorio con los filtros seleccionados.")
