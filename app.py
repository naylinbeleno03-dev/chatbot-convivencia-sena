import time
import streamlit as st
from google import genai
from google.genai import types

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema Digital de Convivencia - INTESAC",
    layout="centered"
)

# Estilos CSS personalizados
st.markdown("""
    <style>
    .stApp {
        background-color: #EBF4FC;
        color: #1E293B;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #3B729F;
        border-right: 1px solid #2C597D;
    }
    section[data-testid="stSidebar"] *, 
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label {
        color: #FFFFFF !important;
    }
    .header-box {
        background-color: #3B729F;
        padding: 22px;
        border-radius: 12px;
        border-left: 6px solid #D4AF37;
        box-shadow: 0 4px 12px rgba(59, 114, 159, 0.2);
        margin-bottom: 24px;
    }
    .header-box h1 {
        color: #FFFFFF !important;
        margin: 0 !important;
        font-size: 1.75rem !important;
        font-weight: 600 !important;
    }
    .header-box p {
        color: #F0F7FF !important;
        margin-top: 6px !important;
        margin-bottom: 0 !important;
        font-size: 0.95rem !important;
    }
    .stButton>button {
        background-color: #FFFFFF !important;
        color: #2C597D !important;
        border: 1px solid #FFFFFF !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #D4AF37 !important;
        color: #FFFFFF !important;
        border-color: #D4AF37 !important;
    }
    .stTextInput>div>div>input {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 6px !important;
    }
    [data-testid="stChatMessage"] {
        background-color: #FFFFFF !important;
        border: 1px solid #DCE7F3 !important;
        border-left: 5px solid #3B729F !important;
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
    [data-testid="stChatInput"] {
        background-color: #FFFFFF !important;
        border: 2px solid #3B729F !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05) !important;
    }
    [data-testid="stChatInput"] textarea {
        color: #000000 !important;
        background-color: #FFFFFF !important;
    }
    hr {
        border-color: #538BB8;
    }
    </style>
""", unsafe_allow_html=True)

# Encabezado superior
st.markdown("""
    <div class="header-box">
        <h1>Sistema Digital de Llamados de Atención y Convivencia Escolar</h1>
        <p>Institución Educativa Técnica Sagrado Corazón de Soledad — Proyecto SENA</p>
    </div>
""", unsafe_allow_html=True)

# Intentar obtener la clave automáticamente desde Secrets o permitir ingreso manual opcional
api_key = st.secrets.get("GEMINI_API_KEY", "")

# Barra lateral izquierda
with st.sidebar:
    st.header("Configuración del Sistema")
    
    if api_key:
        st.success("Conectado con clave API guardada.")
    else:
        api_key = st.text_input("Clave API de Gemini", type="password")
    
    st.markdown("---")
    st.markdown("**Guía de consulta:**")
    st.markdown("1. Ingrese los detalles de la situación acontecida.")
    st.markdown("2. El sistema categorizará el hecho de acuerdo con la Ley 1620 de 2013 y el Manual de Convivencia de INTESAC.")
    st.markdown("3. Se estructurará el procedimiento a seguir y la guía de descargos correspondiente.")
    
    if st.button("Reiniciar consulta"):
        st.session_state.messages = []
        st.rerun()

SYSTEM_PROMPT = """
Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón INTESAC de Soledad.

Tu propósito es asesorar formal y pedagógicamente a la comunidad educativa ante situaciones disciplinarias, asegurando el cumplimiento de la Ley 1620 de 2013, el Decreto 1965 de 2013 y la garantía del debido proceso (Artículo 29 de la Constitución Política).

Importante: Bajo ninguna circunstancia utilices emojis, emoticones ni símbolos gráficos decorativos en tus respuestas. Mantén un formato sobrio, formal y estructurado.

Ante cada caso expuesto por el usuario:
1. Resumen de la situación: Presenta una síntesis objetiva de los hechos reportados.
2. Clasificación de la falta (Según el Manual de Convivencia de INTESAC):
   - Situación Tipo I (Leve): Conflictos manejados inadecuadamente o faltas menores a los deberes.
   - Situación Tipo II (Grave): Situaciones de acoso escolar (bullying), ciberacoso o agresiones físicas/verbales sin incapacidad médica.
   - Situación Tipo III (Gravísima): Presuntos delitos penales, agresiones físicas con incapacidad o porte de elementos prohibidos.
3. Procedimiento institucional: Detalla el protocolo a aplicar según el nivel de falta (llamado de atención verbal, registro en el observador, citación a acudientes o remisión al Comité de Convivencia).
4. Garantías y Debido Proceso: Indica los derechos del estudiante (derecho a ser escuchado, presunción de inocencia, presentación de pruebas y descargos).
5. Modelo de Carta de Descargos: Redacta una plantilla formal de descargos dirigida a la Coordinación o Rectoría de INTESAC.
"""

if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if user_input := st.chat_input("Escriba aquí los hechos de la situación a evaluar..."):
    if not api_key:
        st.warning("Se requiere ingresar la Clave API para continuar.")
        st.stop()

    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    try:
        client = genai.Client(api_key=api_key)
        
        contents = []
        for m in st.session_state.messages:
            role = "user" if m["role"] == "user" else "model"
            contents.append(
                types.Content(
                    role=role,
                    parts=[types.Part.from_text(text=m["content"])]
                )
            )

        with st.chat_message("assistant"):
            with st.spinner("Procesando información institucional..."):
                response = client.models.generate_content(
                    model="gemini-3.6-flash",
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT,
                        temperature=0.2
                    )
                )
                if response and response.text:
                    st.markdown(response.text)
                    st.session_state.messages.append({"role": "assistant", "content": response.text})

    except Exception as e:
        st.error(f"Error de comunicación con el servicio: {e}")
