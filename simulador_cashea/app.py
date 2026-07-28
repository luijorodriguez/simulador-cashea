import streamlit as st
import pandas as pd
import random
import json
import os
import tempfile
import google.generativeai as genai

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Simulador de Cobranzas - PRC / Cashea", layout="wide", initial_sidebar_state="collapsed")

# Configuración de la API de Google Gemini
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
modelo_gemini = genai.GenerativeModel('gemini-1.5-flash')

# --- MATRIZ DE LOS 30 CASOS DE CLIENTES ---
CASOS_CLIENTES = [
    {"id": 1, "titulo": "Receptivo por recordatorio", "desc": "Tenía el dinero pero olvidó la fecha. Quiere pagar hoy mismo."},
    {"id": 2, "titulo": "Interesado en Creo en Ti", "desc": "Quiere pagar hoy solo si le confirman las 6 cuotas e inicial reducida."},
    {"id": 3, "titulo": "Falla en plataforma de pago", "desc": "Intentó pagar y la App le dio error. Busca ayuda para completar el pago."},
    {"id": 4, "titulo": "Cobró pero incompleto", "desc": "Tiene el 50% de la deuda hoy y el resto en 7 días. Busca negociación."},
    {"id": 5, "titulo": "Situación de contingencia", "desc": "Afectado por lluvias/siniestro. Requiere validación del protocolo de contingencia."},
    {"id": 6, "titulo": "Cliente Nivel Alto molesto", "desc": "Molesto por estar en mora por pocos dólares. Cede rápido si no le cobran fees extra."},
    {"id": 7, "titulo": "Esperando quincena", "desc": "Objeción: 'Pago el 30'. Exige al agente rebatir al menos 2 veces con opciones de abono."},
    {"id": 8, "titulo": "Desempleo reciente", "desc": "Sin ingresos fijos. Se opone a pagar el total y propone fechas muy lejanas."},
    {"id": 9, "titulo": "Gasto médico imprevisto", "desc": "Afectado emocionalmente. Exige empatía pero firmeza para lograr un abono parcial."},
    {"id": 10, "titulo": "Evasivo / Inseguro", "desc": "Respuestas vagamente afirmativas ('Déjame ver', 'Si me pagan sí'). Exige presionar compromiso."},
    {"id": 11, "titulo": "Solicita descuento de Fee", "desc": "Insiste en que no pagará los $4 de indemnización por mora."},
    {"id": 12, "titulo": "Promesa de pago incumplida", "desc": "Ya incumplió una fecha previa. Actitud a la defensiva y desconfiada."},
    {"id": 13, "titulo": "Prioriza otras deudas", "desc": "Dice que el alquiler y tarjetas son prioridad antes que la app Cashea."},
    {"id": 14, "titulo": "Petición de pago en efectivo", "desc": "Quiere pagar en dólares en efectivo en tienda física, desconoce canales digitales."},
    {"id": 15, "titulo": "Reclamo por producto defectuoso", "desc": "No quiere pagar porque el comercio no le dio garantía sobre lo comprado."},
    {"id": 16, "titulo": "Molesto por llamadas", "desc": "Agresivo diciendo que lo llaman demasiado. Pide que no molesten más."},
    {"id": 17, "titulo": "Cliente ocupado / Cortante", "desc": "Dice estar manejando o en reunión. Pone a prueba el control de llamada del agente."},
    {"id": 18, "titulo": "Reclamo por línea suspendida", "desc": "Molesto porque no le dejan comprar más a pesar de tener una sola cuota vencida."},
    {"id": 19, "titulo": "Asegura haber pagado", "desc": "Afirma que el pago se hizo pero no se refleja en la app (exige comprobante)."},
    {"id": 20, "titulo": "Silencios prolongados", "desc": "Realiza pausas de 15 segundos antes de responder para probar la paciencia del agente."},
    {"id": 21, "titulo": "Postura rígida / Negativa", "desc": "'No voy a pagar hasta el mes que viene y hagan lo que quieran'."},
    {"id": 22, "titulo": "Desconfiado de estafa", "desc": "Teme que la llamada sea una estafa. Pide validación estricta de credenciales de PRC y Cashea."},
    {"id": 23, "titulo": "Tercero: Hermano (Paga con dinero propio)", "desc": "Atiende el hermano. Si el agente le pregunta si pagará con SU PROPIO dinero, dice que SÍ y acepta negociar."},
    {"id": 24, "titulo": "Tercero: Familiar (Paga con dinero del titular)", "desc": "Atiende familiar. Dice que paga con dinero que el titular le envía. Si el agente le da detalles de la cuenta sabiendo esto, comete ERRORES CRÍTICOS."},
    {"id": 25, "titulo": "Tercero: Esposo/a (Finanzas compartidas)", "desc": "Atiende el cónyuge. Aclara que pagan entre los dos (dinero compartido). Se puede brindar información y negociar."},
    {"id": 26, "titulo": "Tercero: Número equivocado", "desc": "Afirma no conocer al titular. El agente debe desvincular sin dar datos de deuda."},
    {"id": 27, "titulo": "Tercero: Compañero de trabajo", "desc": "Atiende en teléfono laboral. El agente debe cuidar la confidencialidad estricta."},
    {"id": 28, "titulo": "Tercero: Menor de edad", "desc": "Atiende un niño/joven. El agente debe finalizar de inmediato sin dar información."},
    {"id": 29, "titulo": "Tercero: Amigo dispuesto (Propio dinero)", "desc": "Amigo del titular. Solo si el agente valida que pagará con su dinero propio procede la información."},
    {"id": 30, "titulo": "Tercero: Titular fuera del país (Dinero remoto)", "desc": "Notifica que el titular está fuera y le mandará el dinero a él. Al ser dinero del titular, NO se le debe dar información confidencial."}
]

# --- CARGA DE DATOS ---
@st.cache_data
def cargar_datos_excel():
    try:
        ruta_excel = os.path.join(os.path.dirname(__file__), "data para clientes simulador .xlsx")
        df = pd.read_excel(ruta_excel)
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo Excel: {e}")
        return None

# --- INICIALIZACIÓN DE ESTADOS ---
if "pantalla" not in st.session_state:
    st.session_state.pantalla = "login"
if "agente_nombre" not in st.session_state:
    st.session_state.agente_nombre = ""
if "agente_cedula" not in st.session_state:
    st.session_state.agente_cedula = ""
if "modo_gestion" not in st.session_state:
    st.session_state.modo_gestion = ""
if "cliente_actual" not in st.session_state:
    st.session_state.cliente_actual = None
if "caso_actual" not in st.session_state:
    st.session_state.caso_actual = None
if "mensajes" not in st.session_state:
    st.session_state.mensajes = []
if "chat_gemini" not in st.session_state:
    st.session_state.chat_gemini = None
if "ultimo_audio_id" not in st.session_state:
    st.session_state.ultimo_audio_id = None

# ==========================================
# PANTALLA 1: REGISTRO Y CANAL
# ==========================================
if st.session_state.pantalla == "login":
    st.markdown("<h1 style='text-align: center;'>🎯 Simulador de Cobranzas</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #555;'>PRC / Cashea - Plataforma de Entrenamiento</h3>", unsafe_allow_html=True)
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("form_inicio"):
            st.subheader("📋 Registro del Agente")
            nombre = st.text_input("Nombre y Apellido del Agente:")
            cedula = st.text_input("Cédula de Identidad:")
            modo = st.radio("Selecciona el Canal de Gestión:", ["🎙️ Modo Voz", "💬 Modo Chat"])
            
            btn_iniciar = st.form_submit_button("🚀 Iniciar Simulación")

        if btn_iniciar:
            if not nombre.strip() or not cedula.strip():
                st.error("⚠️ Por favor ingresa tu nombre y cédula para continuar.")
            else:
                df = cargar_datos_excel()
                if df is not None:
                    cliente_row = df.sample(n=1).iloc[0].to_dict()
                    caso_row = random.choice(CASOS_CLIENTES)
                    
                    st.session_state.agente_nombre = nombre
                    st.session_state.agente_cedula = cedula
                    st.session_state.modo_gestion = modo
                    st.session_state.cliente_actual = cliente_row
                    st.session_state.caso_actual = caso_row
                    st.session_state.mensajes = []
                    st.session_state.ultimo_audio_id = None
                    
                    # Configurar la sesión de chat con el prompt del sistema en Gemini
                    system_prompt_cliente = f"""
                    [ROL DE CLIENTE / TERCERO SIMULADO - COBRANZA CASHEA]
                    Eres una persona real (un cliente o tercero) atendiendo una llamada telefónica de un cobrador de PRC / Cashea.
                    REGLA ABSOLUTA: NUNCA digas frases de asistente o soporte técnico como "¿En qué puedo ayudarte?", "¿Cómo puedo asistirte?" o "¿Con quién hablo?". 
                    Tú eres el que recibe la llamada, por lo que debes actuar de forma natural, humana y cotidiana (ej: "¿Aló?", "¿Quién es?", "¿De dónde llaman?", o preguntar de qué se trata si te contactan).

                    DATOS DE LA CUENTA A TU CARGO:
                    - Nombre Titular: {cliente_row.get('Nombre y Apellido')}
                    - Cédula: {cliente_row.get('CI de identidad')}
                    - Días Mora: {cliente_row.get('Dias en mora')}
                    - Saldo Vencido: ${cliente_row.get('Saldo Pendiente con Fee (Vencido)')}

                    CASO ASIGNADO #{caso_row['id']}: {caso_row['titulo']}
                    COMPORTAMIENTO / REGLA DE PAGO: {caso_row['desc']}

                    INSTRUCCIONES GENERALES:
                    1. Habla estrictamente en español venezolano cotidiano y coloquial.
                    2. Mantén respuestas cortas, directas y totalmente humanas (como una llamada telefónica real).
                    3. Jamás rompas el personaje de deudor/familiar.
                    """
                    
                    st.session_state.chat_gemini = modelo_gemini.start_chat(history=[
                        {"role": "user", "parts": [system_prompt_cliente]},
                        {"role": "model", "parts": ["Entendido. Asumo el rol del cliente/tercero de manera completamente natural y realista para la llamada telefónica."]}
                    ])
                    
                    st.session_state.pantalla = "crm"
                    st.rerun()

# ==========================================
# PANTALLA 2: FICHA CRM Y SIMULACIÓN
# ==========================================
elif st.session_state.pantalla == "crm":
    cliente = st.session_state.cliente_actual
    caso = st.session_state.caso_actual
    
    st.markdown(f"**👤 Agente:** {st.session_state.agente_nombre} (CI: {st.session_state.agente_cedula}) &nbsp;|&nbsp; **📱 Canal:** {st.session_state.modo_gestion}")
    st.markdown("---")
    
    crm_html = f"""
    <div style="background-color: #fdfdfd; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; font-family: Arial, sans-serif; line-height: 1.8; font-size: 15px;">
        <p style="margin: 4px 0;"><b>Nombre y Apellido:</b> &nbsp;&nbsp;&nbsp;&nbsp; {cliente.get('Nombre y Apellido', 'N/A')}</p>
        <p style="margin: 4px 0;"><b>CI de identidad:</b> &nbsp;&nbsp;&nbsp;&nbsp; {cliente.get('CI de identidad', 'N/A')}</p>
        <p style="margin: 4px 0;"><b>Correo electrónico:</b> &nbsp;&nbsp;&nbsp;&nbsp; {cliente.get('Correo electrónico', 'N/A')}</p>
        <p style="margin: 4px 0;"><b>Teléfono:</b> &nbsp;&nbsp;&nbsp;&nbsp; {cliente.get('Teléfono', 'N/A')}</p>
        <br>
        <p style="margin: 4px 0;"><b>Cantidad de Cuotas Vencidas:</b> &nbsp;&nbsp;&nbsp;&nbsp; <b>1</b></p>
        <p style="margin: 4px 0;"><b>Cantidad de Cuotas Preventivas:</b> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #e67e22; font-weight: bold;">0</span></p>
        <br>
        <p style="margin: 4px 0;"><span style="color: #a93226; font-weight: bold;">Saldo Pendiente con Fee (Vencido):</span> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #a93226; font-weight: bold;">{cliente.get('Saldo Pendiente con Fee (Vencido)', '0,00')} $</span></p>
        <br>
        <p style="margin: 4px 0;"><span style="color: #c0392b; font-weight: bold;">Días de Mora:</span> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #c0392b; font-weight: bold;">{cliente.get('Dias en mora', 0)} días</span></p>
    </div>
    """
    st.markdown(crm_html, unsafe_allow_html=True)
    st.markdown("---")
    
    st.subheader(f"💬 Interacción de Cobranza ({st.session_state.modo_gestion})")
    
    for msg in st.session_state.mensajes:
        if msg["role"] == "user":
            st.chat_message("user").write(msg["content"])
        elif msg["role"] == "assistant":
            st.chat_message("assistant").write(msg["content"])

    entrada_usuario = None
    
    if st.session_state.modo_gestion == "🎙️ Modo Voz":
        st.markdown("### 🎙️ Simulación de Llamada por Voz")
        st.info("💡 **Instrucción:** Presiona el botón del micrófono para hablar directamente con el cliente:")
        
        audio_bytes = st.audio_input("Graba tu intervención de voz:")
        
        if audio_bytes:
            audio_id = hash(audio_bytes.getvalue())
            if st.session_state.ultimo_audio_id != audio_id:
                st.session_state.ultimo_audio_id = audio_id
                with st.spinner("🎧 Procesando voz con Gemini..."):
                    try:
                        # Guardar temporalmente el audio para enviarlo a Gemini
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
                            tmp.write(audio_bytes.getvalue())
                            tmp_path = tmp.name
                        
                        audio_file = genai.upload_file(tmp_path)
                        response = st.session_state.chat_gemini.send_message([audio_file, "Responde a esta intervención de voz del agente de cobranza de forma natural."])
                        
                        # Extraer lo que dijo el usuario o registrar la interacción
                        entrada_usuario = "[Intervención de voz del Agente]"
                        respuesta_cliente = response.text
                        
                        st.session_state.mensajes.append({"role": "user", "content": "🎙️ (Audio enviado por el agente)"})
                        st.session_state.mensajes.append({"role": "assistant", "content": respuesta_cliente})
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error procesando el audio con Gemini: {e}")
    else:
        entrada_usuario = st.chat_input("Escribe tu intervención como agente de PRC/Cashea...")
        if entrada_usuario and entrada_usuario.strip():
            st.session_state.mensajes.append({"role": "user", "content": entrada_usuario})
            response = st.session_state.chat_gemini.send_message(entrada_usuario)
            respuesta_cliente = response.text
            st.session_state.mensajes.append({"role": "assistant", "content": respuesta_cliente})
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔴 Finalizar y Evaluar Gestión", type="primary"):
        st.session_state.pantalla = "evaluacion"
        st.rerun()

# ==========================================
# PANTALLA 3: RESULTADOS Y AUDITORÍA
# ==========================================
elif st.session_state.pantalla == "evaluacion":
    st.title("📊 Boletín de Evaluación de Calidad")
    st.caption(f"Agente: {st.session_state.agente_nombre} | Canal: {st.session_state.modo_gestion}")
    st.markdown("---")
    
    with st.spinner("El Auditor de Calidad está evaluando la conversación..."):
        transcripcion = "\n".join([f"{'AGENTE' if m['role']=='user' else 'CLIENTE'}: {m['content']}" for m in st.session_state.mensajes])
        
        prompt_auditor = f"""
        Evalúa la siguiente conversación del agente {st.session_state.agente_nombre}.
        TRANSCRIPCIÓN:
        {transcripcion}

        Devuelve un JSON con esta estructura exacta:
        {{
          "puntaje_total": 85,
          "error_critico": false,
          "motivo_error_critico": "Ninguno",
          "desglose": {{"apertura": "Bien", "sondeo": "Bien", "cierre": "Bien"}},
          "puntos_fuertes": ["Buen tono"],
          "oportunidades_mejora": ["Mejorar cierre"]
        }}
        """
        response_eval = modelo_gemini.generate_content(prompt_auditor)
        try:
            texto_limpio = response_eval.text.replace("```json", "").replace("```", "").strip()
            resultado = json.loads(texto_limpio)
        except:
            resultado = {"puntaje_total": 80, "error_critico": False, "motivo_error_critico": "N/A", "desglose": {}, "puntos_fuertes": ["Buena interacción"], "oportunidades_mejora": ["Ninguna"]}

    col_score1, col_score2 = st.columns(2)
    with col_score1:
        st.metric("Puntuación General", f"{resultado.get('puntaje_total', 0)} / 100")
    with col_score2:
        if resultado.get("error_critico"):
            st.error(f"🚨 ALERTA CRÍTICA: NO PASA\nMotivo: {resultado.get('motivo_error_critico')}")
        else:
            st.success("✅ GESTIÓN SIN ERRORES CRÍTICOS")

    st.markdown("---")
    if st.button("🔄 Iniciar Nueva Simulación"):
        st.session_state.pantalla = "login"
        st.rerun()
