import time
import streamlit as st
from google import genai
from google.genai import types

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema Integral de Convivencia Escolar - INTESAC",
    layout="centered"
)

# Estilos CSS personalizados
st.markdown("""
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
    section[data-testid="stSidebar"] li {
        color: #FFFFFF !important;
    }

    /* 3. Encabezado superior en Azul más oscuro con letras blancas */
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
        font-size: 1.75rem !important;
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

    /* Botón de Enviar en Negro */
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
""", unsafe_allow_html=True)

# Encabezado superior
st.markdown("""
    <div class="header-box">
        <h1>Sistema Integral de Convivencia Escolar</h1>
        <p>Institución Educativa Técnica Sagrado Corazón de Soledad</p>
    </div>
""", unsafe_allow_html=True)

# Obtener clave API automáticamente si está guardada en Secrets o pedirla
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

# Prompt de sistema institucional
SYSTEM_PROMPT = """
Eres el asistente institucional del Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de la Institución Educativa Técnica Sagrado Corazón INTESAC de Soledad.

Tu propósito es asesorar formal y pedagógicamente a la comunidad educativa ante situaciones disciplinarias, asegurando el cumplimiento de la Ley 1620 de 2013, el Decreto 1965 de 2013 y la garantía del debido proceso (Artículo 29 de la Constitución Política).

Importante: Bajo ninguna circunstancia utilices emojis, emoticones ni símbolos gráficos decorativos en tus respuestas. Mantén un formato sobrio, formal y estructurado.

Ante cada caso expuesto por el usuario:
1. Resumen de la situación: Presenta una síntesis objetiva de los hechos reportados.
2. Clasificación de la falta (Según el Manual de Convivencia de INTESAC):
   - Situación Tipo I (Leve): Conflictos manejados inadecuadamente o faltas menores a los deberes (incluye presentación personal, exceso de maquillaje, corte de cabello no acorde al manual, uso de accesorios no permitidos, impuntualidad).
   - Situación Tipo II (Grave): Situaciones de acoso escolar (bullying), ciberacoso o agresiones físicas/verbales sin incapacidad médica.
   - Situación Tipo III (Gravísima): Presuntos delitos penales, agresiones físicas con incapacidad o porte de elementos prohibidos.
3. Procedimiento institucional: Detalla el protocolo a aplicar según el nivel de falta (llamado de atención verbal, registro en el observador, citación a acudientes o remisión al Comité de Convivencia).
4. Garantías y Debido Proceso: Indica los derechos del estudiante (derecho a ser escuchado, presunción de inocencia, presentación de pruebas y descargos).
5. Modelo de Carta de Descargos: Redacta una plantilla formal de descargos dirigida a la Coordinación o Rectoría de INTESAC.
"""

# Mensaje automático de bienvenida
WELCOME_MESSAGE = """
Saludos. Bienvenido al Sistema Digital de Llamados de Atención y Seguimiento de Convivencia Escolar de INTESAC.

Este portal brinda orientación sobre el protocolo disciplinario institucional y el debido proceso de acuerdo con el Manual de Convivencia y la Ley 1620 de 2013.

Por favor, seleccione una de las situaciones predeterminadas a continuación o redacte detalladamente lo sucedido en la casilla de texto inferior.
"""

# Inicialización del historial de chat con saludo automático
if "messages" not in st.session_state or len(st.session_state.messages) == 0:
    st.session_state.messages = [
        {"role": "assistant", "content": WELCOME_MESSAGE}
    ]

# Despliegue de mensajes anteriores
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

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
    with col2:
        if st.button("Corte de cabello / Uniforme"):
            selected_option = "Se presenta un llamado de atención por corte de cabello inadecuado o porte incorrecto del uniforme institucional."
        if st.button("Incumplimiento de deberes / Asistencia"):
            selected_option = "Se presentó un incumplimiento en los deberes académicos, faltas de asistencia o impuntualidad."
        if st.button("Uso no autorizado de celular/equipos"):
            selected_option = "Se reporta el uso no autorizado de teléfono celular o dispositivos electrónicos durante la jornada escolar."

# Captura de mensaje del usuario
user_input = st.chat_input("Escriba aquí los hechos de la situación a evaluar...")
prompt = selected_option or user_input

if prompt:
    if not api_key:
        st.warning("Se requiere ingresar la Clave API para continuar.")
        st.stop()

    st.session_state.messages.append({"role": "user", "content": prompt})
    st.rerun()

# Generación de respuesta utilizando únicamente gemini-3.6-flash
if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
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
                    st.rerun()

    except Exception as e:
        st.error(f"Error de comunicación con Gemini 3.6 Flash: {e}")
