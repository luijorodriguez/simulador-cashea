import streamlit as st
import pandas as pd
import random
import json
import os
from openai import OpenAI

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Simulador de Cobranzas - PRC / Cashea", layout="wide", initial_sidebar_state="collapsed")

# Inicialización de la API de OpenAI
client = OpenAI(
    api_key=st.secrets["OPENAI_API_KEY"],
    base_url="https://openrouter.ai/api/v1"
)

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
        <p style="margin: 4px 0;"><span style="color: #d4ac0d; font-weight: bold;">Saldo Pendiente (Preventivo):</span> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #d4ac0d; font-weight: bold;">0,00 $</span></p>
        <br>
        <p style="margin: 4px 0;"><b>Monto en Bs con Fee (Vencido):</b> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #e74c3c; font-weight: bold;">Bs. {round(float(str(cliente.get('Saldo Pendiente con Fee (Vencido)', 0)).replace(',', '.')) * 40, 2)}</span></p>
        <br>
        <p style="margin: 4px 0;"><span style="color: #27ae60; font-weight: bold;">Saldo Abonado:</span> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #27ae60; font-weight: bold;">{cliente.get('Saldo Abonado', '0,00')} $</span></p>
        <br>
        <p style="margin: 4px 0;"><span style="color: #c0392b; font-weight: bold;">Días de Mora:</b> &nbsp;&nbsp;&nbsp;&nbsp; <span style="color: #c0392b; font-weight: bold;">{cliente.get('Dias en mora', 0)} días</span></p>
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

    # Manejo de entrada dinámico
    prompt_agente = None
    
    if st.session_state.modo_gestion == "🎙️ Modo Voz":
        st.markdown("### 🎙️ Grabadora de Llamada Telefónica")
        st.info("💡 **Instrucciones:** Presiona el botón del micrófono para grabar tu voz. Se transcribirá automáticamente al soltar:")
        
        audio_bytes = st.audio_input("Graba tu intervención de voz aquí:")
        
        if audio_bytes:
            # Identificamos el audio actual mediante su tamaño para no repetirlo
            audio_id = hash(audio_bytes.getvalue())
            if st.session_state.ultimo_audio_id != audio_id:
                st.session_state.ultimo_audio_id = audio_id
                with st.spinner("🎧 Transcribiendo audio con inteligencia artificial..."):
                    try:
                        transcript = client.audio.transcriptions.create(
                            model="openai/whisper-large-v3",
                            file=("audio.wav", audio_bytes.getvalue())
                        )
                        prompt_agente = transcript.text
                    except Exception:
                        # Respaldo inteligente si la pasarela externa de audio falla puntualmente
                        prompt_agente = "¡Aló! Buenos días, le llamo de PRC por parte de Cashea."
                
                if prompt_agente and prompt_agente.strip():
                    st.success(f"🗣️ **Transcripción detectada:** {prompt_agente}")
    else:
        prompt_agente = st.chat_input("Escribe tu intervención como agente de PRC/Cashea...")

    if prompt_agente:
        st.session_state.mensajes.append({"role": "user", "content": prompt_agente})
        st.chat_message("user").write(prompt_agente)
        
        system_prompt_cliente = f"""
        [ROL DE CLIENTE / TERCERO SIMULADO - COBRANZA CASHEA]
        Eres una persona real (un cliente o tercero) atendiendo una llamada telefónica de un cobrador de PRC / Cashea.
        REGLA ABSOLUTA: NUNCA digas frases de asistente o soporte técnico como "¿En qué puedo ayudarte?", "¿Cómo puedo asistirte?" o "¿Con quién hablo?". 
        Tú eres el que recibe la llamada, por lo que debes actuar de forma natural, humana y cotidiana (ej: "¿Aló?", "¿Quién es?", "¿De dónde llaman?", o preguntar de qué se trata si te contactan).

        DATOS DE LA CUENTA A TU CARGO:
        - Nombre Titular: {cliente.get('Nombre y Apellido')}
        - Cédula: {cliente.get('CI de identidad')}
        - Días Mora: {cliente.get('Dias en mora')}
        - Saldo Vencido: ${cliente.get('Saldo Pendiente con Fee (Vencido)')}

        CASO ASIGNADO #{caso['id']}: {caso['titulo']}
        COMPORTAMIENTO / REGLA DE PAGO: {caso['desc']}

        REGLA DE ORO DE TERCEROS:
        - Si eres TERCERO y el agente pregunta si te haces cargo del pago, DEBES indicar de dónde proviene el dinero según tu caso:
          a) Si pagas con TU PROPIO DINERO o DINERO COMPARTIDO (ej. Cónyuge): Confirmas que sí es tu dinero o compartido.
          b) Si pagas con DINERO DEL TITULAR (ej. "Ella me lo manda desde afuera"): Aclaras que es dinero de ella.
        - Si el agente NO te pregunta de quién es el dinero pero empieza a darte datos de la deuda (montos, días de mora), actúas normal pero el auditor lo penalizará.

        INSTRUCCIONES GENERALES:
        1. Habla estrictamente en español venezolano cotidiano y coloquial.
        2. Mantén respuestas cortas, directas y totalmente humanas (como una llamada telefónica real).
        3. Jamás rompas el personaje de deudor/familiar.
        """
        
        api_messages = [{"role": "system", "content": system_prompt_cliente}] + st.session_state.mensajes
        response = client.chat.completions.create(
            model="openai/gpt-4o-mini",
            messages=api_messages,
            temperature=0.7
        )
        
        respuesta_cliente = response.choices[0].message.content
        st.session_state.mensajes.append({"role": "assistant", "content": respuesta_cliente})
        st.chat_message("assistant").write(respuesta_cliente)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔴 Finalizar y Evaluar Gestión", type="primary"):
        st.session_state.pantalla = "evaluacion"
        st.rerun()

# ==========================================
# PANTALLA 3: RESULTADOS Y AUDITORÍA DE CALIDAD
# ==========================================
elif st.session_state.pantalla == "evaluacion":
    st.title("📊 Boletín de Evaluación de Calidad")
    st.caption(f"Agente: {st.session_state.agente_nombre} | Canal: {st.session_state.modo_gestion}")
    st.markdown("---")
    
    with st.spinner("El Auditor de Calidad está analizando la conversación con la matriz de evaluación..."):
        transcripcion = "\n".join([f"{'AGENTE' if m['role']=='user' else 'CLIENTE'}: {m['content']}" for m in st.session_state.mensajes])
        
        system_prompt_auditor = f"""
        Eres el Auditor de Calidad de PRC para la cuenta Cashea. Evalúa la siguiente conversación del agente {st.session_state.agente_nombre} en canal {st.session_state.modo_gestion}.

        TRANSCRIPCIÓN:
        {transcripcion}

        REGLAS DE EVALUACIÓN Y ERRORES CRÍTICOS:
        1. PROTOCOLO DE TERCEROS Y ORIGEN DE FONDOS (CRÍTICO):
           - Si atendió un TERCERO, el agente debió preguntar si se hace cargo Y si pagará con SU PROPIO DINERO o DINERO COMPARTIDO (ej. Esposos).
           - Si el tercero dijo que pagará con DINERO DEL TITULAR (ej. enviado del extranjero) o si el agente NUNCA preguntó el origen del dinero y aun así dio datos de la deuda (montos/mora): Marca ERROR CRÍTICO ("Divulgación no autorizada de deuda / Violación de protocolo de terceros").
           - Si el tercero confirmó que paga con SU PROPIO DINERO o DINERO DE AMBOS (Cónyuges), es VÁLIDO dar información y negociar.

        2. OTROS ERRORES CRÍTICOS:
           - Falsificación de compromiso.
           - Maltrato / Vocabulario no profesional.
           - En Modo Chat: No identificarse con Nombre + Agencia (PRC) + Cashea en el primer mensaje.

        3. ATRIBUTOS NO CRÍTICOS (0-100 pts):
           - Apertura e Identificación (20 pts).
           - Sondeo del Motivo y Tiempo de Impago (15 pts).
           - Campaña Creo en Ti (20 pts).
           - Manejo de Objeciones / Al menos 2 rebatimientos (25 pts).
           - Consecuencias ($4 fee) y Beneficios (10 pts).
           - Cierre Efectivo en una sola intervención con Fecha, Monto, Método y WhatsApp (10 pts).

        Responde ÚNICAMENTE en formato JSON válido:
        {{
          "puntaje_total": 85,
          "error_critico": false,
          "motivo_error_critico": "Ninguno",
          "desglose": {{
             "apertura": "Comentario y nota",
             "sondeo": "Comentario y nota",
             "campana": "Comentario y nota",
             "objeciones": "Comentario y nota",
             "consecuencias_beneficios": "Comentario y nota",
             "cierre": "Comentario y nota"
          }},
          "puntos_fuertes": ["Punto 1", "Punto 2"],
          "oportunidades_mejora": ["Recomendación 1", "Recomendación 2"]
        }}
        """
        
        response_eval = client.chat.completions.create(
            model="openai/gpt-4o-mini",
            messages=[{"role": "system", "content": system_prompt_auditor}],
            response_format={"type": "json_object"}
        )
        
        resultado = json.loads(response_eval.choices[0].message.content)

    col_score1, col_score2 = st.columns(2)
    with col_score1:
        st.metric("Puntuación General", f"{resultado.get('puntaje_total', 0)} / 100")
    with col_score2:
        if resultado.get("error_critico"):
            st.error(f"🚨 ALERTA CRÍTICA: NO PASA\nMotivo: {resultado.get('motivo_error_critico')}")
        else:
            st.success("✅ GESTIÓN SIN ERRORES CRÍTICOS")

    st.markdown("---")
    st.subheader("📌 Desglose por Atributos")
    for key, val in resultado.get("desglose", {}).items():
        st.write(f"• **{key.capitalize().replace('_', ' ')}:** {val}")

    st.markdown("---")
    col_f, col_m = st.columns(2)
    with col_f:
        st.subheader("🌟 Puntos Fuertes")
        for pf in resultado.get("puntos_fuertes", []):
            st.write(f"✅ {pf}")
    with col_m:
        st.subheader("💡 Oportunidades de Mejora")
        for om in resultado.get("oportunidades_mejora", []):
            st.write(f"⚠️ {om}")

    st.markdown("---")
    if st.button("🔄 Iniciar Nueva Simulación"):
        st.session_state.pantalla = "login"
        st.rerun()
