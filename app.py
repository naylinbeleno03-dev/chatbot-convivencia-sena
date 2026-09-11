import streamlit as st
from google import genai
from google.genai import types

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema Digital de Convivencia - INTESAC",
    layout="centered"
)

# Inyección de estilos CSS personalizados (Azul, Dorado y Negro)
st.markdown("""
    <style>
    /* Fondo principal y tipografía general */
    .stApp {
        background-color: #050B14;
        color: #E2E8F0;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Barra lateral */
    section[data-testid="stSidebar"] {
        background-color: #02050A;
        border-right: 1px solid #D4AF37;
    }

    /* Encabezados y títulos */
    h1 {
        color: #D4AF37 !important;
        font-weight: 600 !important;
        letter-spacing: 0.5px;
        border-bottom: 2px solid #1E3A8A;
        padding-bottom: 10px;
    }
    
    h2, h3, .stSidebar h2 {
        color: #D4AF37 !important;
        font-weight: 500 !important;
    }

    /* Subtítulos y texto secundario */
    .stCaption {
        color: #94A3B8 !important;
        font-size: 0.95rem !important;
    }

    /* Botones principales */
    .stButton>button {
        background-color: #0F172A;
        color: #D4AF37;
        border: 1px solid #D4AF37;
        border-radius: 4px;
        font-weight: 500;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton>button:hover {
        background-color: #D4AF37;
        color: #050B14;
        border-color: #D4AF37;
    }

    /* Campos de texto de entrada */
    .stTextInput>div>div>input {
        background-color: #0F172A;
        color: #FFFFFF;
        border: 1px solid #1E3A8A;
        border-radius: 4px;
    }

    .stTextInput>div>div>input:focus {
        border-color: #D4AF37;
        box-shadow: 0 0 5px rgba(212, 175, 55, 0.4);
    }

    /* Cuadro de chat del usuario y asistente */
    .stChatMessage {
        background-color: #0F172A;
        border: 1px solid #1E3A8A;
        border-radius: 6px;
        margin-bottom: 12px;
    }

    /* Ajustes del separador horizontal */
    hr {
        border-color: #1E3A8A;
    }
    </style>
""", unsafe_allow_html=True)

st.title("Sistema Digital de Llamados de Atención y Convivencia Escolar")
st.caption("Institución Educativa Técnica Sagrado Corazón INTESAC de Soledad — Proyecto SENA")

# Barra lateral de configuración
with st.sidebar:
    st.header("Configuración del Sistema")
    api_key = st.text_input("Clave API de Gemini", type="password")
    
    st.markdown("---")
    st.markdown("**Guía de consulta:**")
    st.markdown("1. Ingrese los detalles de la situación acontecida.")
    st.markdown("2. El sistema categorizará el hecho de acuerdo con la Ley 1620 de 2013 y el Manual de Convivencia de INTESAC.")
    st.markdown("3. Se estructurará el procedimiento a seguir y la guía de descargos correspondiente.")
    
    if st.button("Reiniciar consulta"):
        st.session_state.messages = []
        st.rerun()

# Configuración del prompt de sistema
SYSTEM_PROMPT = """
Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón INTESAC de Soledad[span_0](start_span)[span_0](end_span).

Tu propósito es asesorar formal y pedagógicamente a la comunidad educativa ante situaciones disciplinarias, asegurando el cumplimiento de la Ley 1620 de 2013, el Decreto 1965 de 2013 y la garantía del debido proceso (Artículo 29 de la Constitución Política)[span_1](start_span)[span_1](end_span).

Importante: Bajo ninguna circunstancia utilices emojis, emoticones ni símbolos gráficos decorativos en tus respuestas. Mantén un formato sobrio, formal y estructurado.

Ante cada caso expuesto por el usuario:
1. Resumen de la situación: Presenta una síntesis objetiva de los hechos reportados.
2. Clasificación de la falta (Según el Manual de Convivencia de INTESAC)[span_2](start_span)[span_2](end_span):
   - Situación Tipo I (Leve): Conflictos manejados inadecuadamente o faltas menores a los deberes[span_3](start_span)[span_3](end_span).
   - Situación Tipo II (Grave): Situaciones de acoso escolar (bullying), ciberacoso o agresiones físicas/verbales sin incapacidad médica[span_4](start_span)[span_4](end_span).
   - Situación Tipo III (Gravísima): Presuntos delitos penales, agresiones físicas con incapacidad o porte de elementos prohibidos[span_5](start_span)[span_5](end_span).
3. Procedimiento institucional: Detalla el protocolo a aplicar según el nivel de falta (llamado de atención verbal, registro en el observador, citación a acudientes o remisión al Comité de Convivencia)[span_6](start_span)[span_6](end_span).
4. Garantías y Debido Proceso: Indica los derechos del estudiante (derecho a ser escuchado, presunción de inocencia, presentación de pruebas y descargos)[span_7](start_span)[span_7](end_span).
5. Modelo de Carta de Descargos: Redacta una plantilla formal de descargos dirigida a la Coordinación o Rectoría de INTESAC.
"""

# Inicialización del historial
if "messages" not in st.session_state:
    st.session_state.messages = []

# Despliegue de mensajes en el chat
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Captura de entrada del usuario
if user_input := st.chat_input("Escriba aquí los hechos de la situación a evaluar..."):
    if not api_key:
        st.warning("Se requiere ingresar la Clave API en la barra lateral para continuar.")
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
                    model="gemini-2.5-flash",
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_PROMPT,
                        temperature=0.2
                    )
                )
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})

    except Exception as e:
        st.error(f"Error de comunicación con el servicio: {e}")