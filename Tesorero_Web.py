import streamlit as st
import gspread
from gspread.utils import rowcol_to_a1  # 🛠️ Traduce coordenadas a celdas Excel (Ej: B5)
from google.oauth2.service_account import Credentials
from datetime import datetime
from zoneinfo import ZoneInfo  
import unicodedata
import pandas as pd
import cloudinary
import cloudinary.uploader

# ==========================================
# 1. CONFIGURACIÓN DE IDENTIDAD (DINÁMICA)
# ==========================================
NOM_UNIDAD = st.secrets["config"].get("nombre_unidad", "Manada")

# ☁️ Inicializamos Cloudinary con los secretos guardados
cloudinary.config(
    cloud_name = st.secrets["cloudinary"]["cloud_name"],
    api_key = st.secrets["cloudinary"]["api_key"],
    api_secret = st.secrets["cloudinary"]["api_secret"],
    secure = True
)

# Definition de Emojis
EMOJI_NORMAL = "🐾" 
emoji_secret = st.secrets["config"].get("emoji", "")

if emoji_secret.startswith("data:image"):
    repo_name = f"tesoreria-{NOM_UNIDAD.lower().strip()}"
    img_name = "lobatos.png" if NOM_UNIDAD.lower() == "manada" else "insignia.png"
    url_imagen = f"https://github.com/Gaspar-collab/{repo_name}/blob/main/{img_name}"
    EMOJI_ORIGINAL = f'<img src="{url_imagen}?raw=true" width="26" style="vertical-align: middle;">'
else:
    EMOJI_ORIGINAL = emoji_secret 

EMOJIS_TITULO_DERECHA = f"{EMOJI_NORMAL} {EMOJI_ORIGINAL}"
EMOJIS_TITULO_IZQUIERDA = f"{EMOJI_ORIGINAL} {EMOJI_NORMAL}"

st.set_page_config(page_title=f"Tesorería {NOM_UNIDAD}", layout="centered")
titulo_html = f"## {EMOJIS_TITULO_DERECHA} Tesorería de la {NOM_UNIDAD} {EMOJIS_TITULO_IZQUIERDA}"

if NOM_UNIDAD.lower() == "manada":
    TEXTO_INDIVIDUAL = "niño"
    TEXTO_PLURAL = "niños"
else:
    TEXTO_INDIVIDUAL = "joven"
    TEXTO_PLURAL = "jóvenes"

SPREADSHEET_ID = st.secrets["config"]["spreadsheet_id"]

# 🔐 Permisos para Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

ahora_cl = datetime.now(ZoneInfo("America/Santiago"))

# ========================================================
# 🔐 ESCUDO DE SEGURIDAD (LOGIN PARA DIRIGENTES)
# ========================================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.markdown(f"## {EMOJIS_TITULO_IZQUIERDA} Acceso Restringido - Tesorería", unsafe_allow_html=True)
    st.write(f"Ingresa la contraseña de la {NOM_UNIDAD} para registrar movimientos.")

    clave_maestra = st.secrets["credenciales"]["clave_compartida"]
    pass_input = st.text_input("Contraseña de acceso", type="password")

    if st.button("Ingresar al Sistema"):
        if pass_input == clave_maestra:
            st.session_state["autenticado"] = True
            st.success(f"¡Bienvenido! Cargando panel de la {NOM_UNIDAD}...")
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta. ¡Inténtalo de nuevo!")
            
    st.stop()

# ========================================================
# 🛠️ BARRA LATERAL - ACCIONES ADMINISTRATIVAS
# ========================================================
with st.sidebar:
    st.header("⚙️ Panel de Control")
    st.write("Si cambiaste datos directo en el Excel o quieres ver las actualizaciones de inmediato, usa este botón:")
    if st.button("🔄 FORZAR RECARGA DE DATOS", use_container_width=True):
        st.cache_data.clear()
        st.success("¡Memoria borrada! Cargando datos fresquitos... 🚀")
        st.rerun()

# ==========================================
# 2. AUTENTICACIÓN DE GOOGLE Y FUNCIONES
# ==========================================
def autenticar():
    info_credenciales = dict(st.secrets["token_json"])
    if "private_key" in info_credenciales:
        info_credenciales["private_key"] = info_credenciales["private_key"].replace("\\n", "\n")
    
    creds = Credentials.from_service_account_info(info_credenciales, scopes=SCOPES)
    cliente_sheets = gspread.authorize(creds)
    return cliente_sheets

def quitar_tildes(texto):
    texto = str(texto).lower().strip()
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')

@st.cache_data(ttl=600)
def cargar_lista_lobatos():
    try:
        cliente_sheets = autenticar()
        sheet = cliente_sheets.open_by_key(SPREADSHEET_ID)
        hoja_principal = sheet.worksheet("Mensualidades")
        registros = hoja_principal.get_all_values()
        
        nombres = [fila[1] for fila in registros[2:] if len(fila) > 1 and fila[1].strip()]
        return nombres, registros
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return [], []

@st.cache_data(ttl=600)
def obtener_datos_busqueda():
    cliente_sheets = autenticar()
    sheet = cliente_sheets.open_by_key(SPREADSHEET_ID)
    meses_map = ['Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'ene']
    data_total = []
    for mes in meses_map:
        try:
            hoja = sheet.worksheet(mes)
            rows = hoja.get_all_values()
            for row in rows[1:]:
                if len(row) > 1 and str(row[1]).strip().lower() != 'total':
                    row_data = list(row) + [mes] 
                    data_total.append(row_data)
        except: continue
    return data_total

# ==========================================
# 3. CREACIÓN DE PESTAÑAS (TABS)
# ==========================================
st.markdown(titulo_html, unsafe_allow_html=True)
tab1, tab2 = st.tabs(["📝 Registrar Transacción", "📊 Estadísticas y Buscador"])

# ==========================================
# PESTAÑA 1: REGISTRAR TRANSACCIÓN
# ==========================================
with tab1:
    st.write("Registra ingresos y egresos de forma rápida para el Excel.")

    if 'form_id' not in st.session_state:
        st.session_state.form_id = 0
    if 'ultima_fecha' not in st.session_state:
        st.session_state['ultima_fecha'] = ahora_cl.date()
    if 'ultimo_tipo_trans' not in st.session_state:
        st.session_state['ultimo_tipo_trans'] = "Cuota"

    if 'ultimo_registro' in st.session_state:
        st.markdown(f"✅ **Último registro:** {st.session_state['ultimo_registro']}", unsafe_allow_html=True)
        if st.session_state.get('mostrar_globos', False):
            st.balloons()
            st.session_state['mostrar_globos'] = False

    lista_nombres_lobatos, todos_los_registros = cargar_lista_lobatos()

    col1, col2 = st.columns(2)
    with col1:
        lista_meses = ["Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Enero"]
        meses_espanol = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        
        mes_actual_nombre = meses_espanol[ahora_cl.month]
        default_index = lista_meses.index(mes_actual_nombre) if mes_actual_nombre in lista_meses else 0

        mes_seleccionado = st.selectbox("📅 Mes del registro:", lista_meses, index=default_index, key=f'mes_sel_{st.session_state.form_id}')
    with col2:
        tipo_movimiento = st.radio("💰 Tipo de movimiento:", ["Ingreso", "Egreso"], horizontal=True, key=f'tipo_mov_{st.session_state.form_id}')

    lista_tipos_transaccion = ["Cuota", "Inscripción", "Cuota e Inscripción", "Transferencia", "Devolución", "Compra", "Depósito", "Donación"]
    if st.session_state['ultimo_tipo_trans'] in lista_tipos_transaccion:
        idx_tipo_defecto = lista_tipos_transaccion.index(st.session_state['ultimo_tipo_trans'])
    else:
        idx_tipo_defecto = 0

    tipo_transaccion = st.selectbox(
        "📌 Tipo de Transacción:", 
        lista_tipos_transaccion, 
        index=idx_tipo_defecto, 
        key=f'tipo_trans_{st.session_state.form_id}'
    )

    motivo_especifico = ""
    objeto_comprado = ""
    evento_compra = ""
    quien_transfiere = ""
    quien_recibe = ""

    if tipo_transaccion == "Transferencia":
        col_tf1, col_tf2 = st.columns(2)
        with col_tf1:
            quien_transfiere = st.text_input("👤 ¿Quién hace la transferencia? (Origen):", placeholder="Ej: Apoderado Juan Pérez o Unidad", key=f'q_transfiere_{st.session_state.form_id}')
        with col_tf2:
            quien_recibe = st.text_input("👤 ¿Quién recibe la transferencia? (Destino):", placeholder="Ej: Caja Chica o Proveedor", key=f'q_recibe_{st.session_state.form_id}')
        motivo_especifico = st.text_input("🎯 Motivo específico (Ej: Cuota Rifa, Materiales Campamento):", key=f'mot_esp_tf_{st.session_state.form_id}')
        
    elif tipo_transaccion == "Devolución":
        motivo_especifico = st.text_input("🎯 Motivo específico (Ej: Compra campamento de verano):", key=f'mot_esp_dev_{st.session_state.form_id}')
        
    elif tipo_transaccion == "Compra":
        col_compra1, col_compra2 = st.columns(2)
        with col_compra1:
            objeto_comprado = st.text_input("📦 Qué se compró (Ej: Cartulinas y plumones):", key=f'obj_comp_{st.session_state.form_id}')
        with col_compra2:
            evento_compra = st.text_input("📅 Día o Evento (Ej: Consejo de Sábado):", key=f'evt_comp_{st.session_state.form_id}')

    es_pago_lobato = tipo_transaccion in ["Cuota", "Inscripción", "Cuota e Inscripción"]
    nombre_final = ""
    tiene_hermanos = False
    ya_pago_inscripcion = False

    if es_pago_lobato:
        nombre_final = st.selectbox(f"👦 Selecciona al {TEXTO_INDIVIDUAL} de la {NOM_UNIDAD}:", ["-- Selecciona una opción --"] + lista_nombres_lobatos, key=f'nom_lobato_{st.session_state.form_id}')
        if nombre_final != "-- Selecciona una opción --":
            nombre_input_limpio = quitar_tildes(nombre_final)
            headers = todos_los_registros[0]
            for fila in todos_los_registros:
                if len(fila) > 1 and quitar_tildes(fila[1]) == nombre_input_limpio:
                    if len(fila) > 3 and str(fila[3]).strip() not in ['0', '', '0%']:
                        tiene_hermanos = True
                    if 'Inscr.' in headers:
                        idx_inscr = headers.index('Inscr.')
                        if idx_inscr < len(fila) and fila[idx_inscr].strip().upper() in ["TRUE", "1"]:
                            ya_pago_inscripcion = True
                    break
    elif tipo_transaccion == "Transferencia":
        nombre_final = f"{quien_transfiere} a {quien_recibe}"
    else:
        nombre_final = st.text_input("🏢 Nombre de la persona o entidad (Ej: Librería Central u Olave):", key=f'nom_entidad_{st.session_state.form_id}')

    monto = st.number_input("💵 Monto ($):", min_value=0, step=500, format="%d", key=f'monto_val_{st.session_state.form_id}')

    cuotas_calculadas = 0
    num_cuotas_final = 0
    detalles_extra = ""
    minimo_requerido = 0
    monto_base_esperado = 0

    if es_pago_lobato and nombre_final != "-- Selecciona una opción --":
        value_cuota = 11000 if tiene_hermanos else 12000
        
        if tipo_transaccion == "Cuota":
            minimo_requerido = value_cuota
        elif tipo_transaccion == "Inscripción":
            minimo_requerido = 21000
        elif tipo_transaccion == "Cuota e Inscripción":
            minimo_requerido = 21000 + value_cuota

        if tipo_transaccion in ["Inscripción", "Cuota e Inscripción"] and ya_pago_inscripcion:
            st.error(f"❌ **¡Paren todo!** Este {TEXTO_INDIVIDUAL} ya tiene registrada su inscripción como pagada en el sistema.")

        if monto > 0 and monto < minimo_requerido:
            st.error(f"❌ **Monto insuficiente:** El valor ingresado (\${monto:,}) ni siquiera llega al valor mínimo requerido (\${minimo_requerido:,}).")
        else:
            if tipo_transaccion in ["Cuota", "Cuota e Inscripción"] and monto >= minimo_requerido:
                monto_restante_calc = monto
                if tipo_transaccion == "Cuota e Inscripción":
                    monto_restante_calc -= 21000
                    
                cuotas_calculadas = max(0, monto_restante_calc // value_cuota)
                st.info(f"✨ **Cálculo Automático:** El programa detecta que este monto equivale a **{cuotas_calculadas} cuotas**.")
                num_cuotas_final = st.number_input("⚙️ Número de cuotas final:", value=int(cuotas_calculadas), min_value=0, step=1, key=f'num_cuotas_{st.session_state.form_id}')

                if tipo_transaccion == "Cuota":
                    monto_base_esperado = num_cuotas_final * value_cuota
                else:  
                    monto_base_esperado = 21000 + (num_cuotas_final * value_cuota)

            elif tipo_transaccion == "Inscripción" and monto >= minimo_requerido:
                st.info("✨ **Cálculo Automático:** Inscripción detectada ($21,000).")
                num_cuotas_final = 0
                monto_base_esperado = 21000

            if monto > monto_base_esperado and monto > 0:
                st.warning("⚠️ El monto ingresado supera el valor de lo cubierto por cuotas/inscripción.")
                detalles_extra = st.text_input("🔍 Detalla qué más transfirieron (Ej: los curantos, rifa):", placeholder="Ej: los curantos", key=f'detalles_extra_{st.session_state.form_id}')

    fecha = st.date_input(
        "📆 Fecha de la transacción:", 
        value=st.session_state['ultima_fecha'], 
        key=f'fecha_val_{st.session_state.form_id}'
    )

    texto_persona = nombre_final if (nombre_final and nombre_final != "-- Selecciona una opción --") else "[Nombre]"
    texto_motivo = motivo_especifico if motivo_especifico.strip() else "[Motivo]"
    texto_objeto = objeto_comprado if objeto_comprado.strip() else "[Objeto]"
    texto_evento = evento_compra if evento_compra.strip() else "[Día/Evento]"

    predeterminado_final = ""

    if tipo_transaccion == "Devolución":
        predeterminado_final = f"Devolución de dinero por parte de la unidad a {texto_persona} por {texto_motivo}"
    elif tipo_transaccion == "Transferencia":
        texto_transfiere = quien_transfiere if quien_transfiere.strip() else "[Quién envía]"
        texto_recibe = quien_recibe if quien_recibe.strip() else "[Quién recibe]"
        predeterminado_final = f"Transferencia por concepto de {texto_motivo} realizada por {texto_transfiere} a {texto_recibe}"
    elif tipo_transaccion == "Donación":
        predeterminado_final = f"Donación por parte de {texto_persona}"
    elif tipo_transaccion == "Compra":
        sug_objeto = f"Pago por parte de la unidad para comprar {texto_objeto}"
        sug_evento = f"Pago por parte de la unidad para compras de {texto_evento}"
        
        st.write("📋 **Selecciona el comentario predeterminado para usar si dejas el campo vacío:**")
        seleccion_sugerida = st.radio("Opciones disponibles:", [sug_objeto, sug_evento], key=f'radio_compra_{st.session_state.form_id}')
        predeterminado_final = seleccion_sugerida
    elif es_pago_lobato:
        if detalles_extra.strip():
            if tipo_transaccion == "Cuota e Inscripción":
                predeterminado_final = f"Pago de cuota, inscripción y {detalles_extra.strip()} por parte del apoderado/a de {texto_persona}"
            elif tipo_transaccion == "Inscripción":
                predeterminado_final = f"Pago de inscripción y {detalles_extra.strip()} por parte del apoderado/a de {texto_persona}"
            else:
                predeterminado_final = f"Pago de cuota y {detalles_extra.strip()} por parte del apoderado/a de {texto_persona}"
        else:
            predeterminado_final = f"Pago de {tipo_transaccion.lower()} por parte del apoderado/a de {texto_persona}"
    else:
        predeterminado_final = f"Pago por parte de {texto_persona} para {tipo_transaccion.lower()}"

    if tipo_transaccion != "Compra":
        st.info(f"💡 **Comentario:** Si lo dejas en blanco se escribirá:\n\n*{predeterminado_final}*")

    comentario_usuario = st.text_input("📝 Escribe un comentario personalizado si deseas cambiar el predeterminado:", key=f'com_user_{st.session_state.form_id}')

    st.write("---")
    st.subheader("📸 Comprobantes de Pago")
    archivos_comprobantes = st.file_uploader(
        "Toma una foto, sube un pantallazo o un PDF:", 
        type=["png", "jpg", "jpeg", "pdf"],
        accept_multiple_files=True,
        key=f"comprobantes_{st.session_state.form_id}"
    )

    if archivos_comprobantes:
        st.write("👀 **Previsualización de comprobantes:**")
        cols = st.columns(min(len(archivos_comprobantes), 3)) 
        for i, archivo in enumerate(archivos_comprobantes):
            with cols[i % 3]:
                if archivo.type.startswith("image"):
                    st.image(archivo, use_container_width=True)
                else:
                    st.info(f"📄 {archivo.name[:15]}...")
                    
    st.write("---")
    if st.button("🚀 REGISTRAR TRANSACCIÓN", use_container_width=True):
        if not archivos_comprobantes:
            st.error("❌ ¡Paren las prensas! Debes adjuntar al menos un comprobante para registrar la transacción.")
        elif es_pago_lobato and nombre_final == "-- Selecciona una opción --":
            st.error(f"❌ Por favor, selecciona un {TEXTO_INDIVIDUAL} válido de la lista.")
        elif es_pago_lobato and tipo_transaccion in ["Inscripción", "Cuota e Inscripción"] and ya_pago_inscripcion:
            st.error("❌ Operación rechazada: La inscripción de este {TEXTO_INDIVIDUAL} ya figura como pagada.")
        elif es_pago_lobato and monto < minimo_requerido:
            st.error(f"❌ Error en el monto: Debe ingresar al menos el valor mínimo requerido (${minimo_requerido:,}).")
        elif tipo_transaccion == "Transferencia" and (not quien_transfiere.strip() or not quien_recibe.strip()):
            st.error("❌ Por favor, rellena quién realiza y quién recibe la transferencia.")
        elif not es_pago_lobato and tipo_transaccion != "Transferencia" and not nombre_final.strip():
            st.error("❌ Por favor, escribe el nombre de la entidad o persona.")
        elif tipo_transaccion in ["Devolución", "Transferencia"] and not motivo_especifico.strip():
            st.error(f"❌ Por favor, ingresa el motivo específico de la {tipo_transaccion.lower()}.")
        elif tipo_transaccion == "Compra" and not objeto_comprado.strip() and not evento_compra.strip():
            st.error("❌ Por favor, ingresa qué se compró o para qué evento se realizó la compra.")
        elif monto <= 0:
            st.error("❌ El monto debe ser mayor a $0.")
        elif es_pago_lobato and monto > monto_base_esperado and not detalles_extra.strip():
            st.error("❌ Detectamos dinero extra. Por favor, detalla qué más están pagando en la casilla correspondiente.")
        else:
            with st.spinner("Procesando... Subiendo archivos a Cloudinary y actualizando planilla... ⏳"):
                try:
                    cliente_sheets = autenticar()
                    sheet = cliente_sheets.open_by_key(SPREADSHEET_ID)
                    
                    meses_map = {'Abril': 'Abr', 'Mayo': 'May', 'Junio': 'Jun', 'Julio': 'Jul',
                                 'Agosto': 'Ago', 'Septiembre': 'Sep', 'Octubre': 'Oct',
                                 'Noviembre': 'Nov', 'Diciembre': 'Dic', 'Enero': 'ene'}
                    hoja_mes_nombre = meses_map[mes_seleccionado]
                    
                    fecha_str = fecha.strftime("%d/%m/%Y")
                    
                    motivo_limpio = quitar_tildes(tipo_transaccion).replace(' ', '_').title()
                    persona_limpia = quitar_tildes(nombre_final).replace(' ', '_').title()
                    fecha_limpia = fecha_str.replace('/', '-')
                    
                    # 🕒 DETECTAMOS LA HORA EXACTA (HHMM) SIN SEGUNDOS
                    # Ej: 1732 (5 de la tarde con 32 min)
                    hora_exacta_str = datetime.now(ZoneInfo("America/Santiago")).strftime("%H%M")
                    
                    links_comprobantes = []
                    for index, archivo in enumerate(archivos_comprobantes):
                        # Si suben más de un archivo en el mismo formulario, mantenemos el diferenciador de lote
                        sufijo_lote = f"_N{index + 1}" if len(archivos_comprobantes) > 1 else ""
                        
                        # 🚀 CONSTRUIMOS EL NOMBRE INYECTANDO SOLO HORA Y MINUTOS
                        nombre_archivo_cloudinary = f"Comprobante_{motivo_limpio}_{persona_limpia}_{fecha_limpia}_{hora_exacta_str}{sufijo_lote}"
                        
                        resultado_upload = cloudinary.uploader.upload(
                            archivo,
                            public_id=nombre_archivo_cloudinary,
                            folder=f"comprobantes_{NOM_UNIDAD.lower().strip()}",
                            resource_type="auto"
                        )
                        
                        links_comprobantes.append(resultado_upload.get('secure_url'))
                    
                    links_comprobantes_final = "\n".join(links_comprobantes)

                    if not comentario_usuario.strip():
                        comentario_final = predeterminado_final
                    else:
                        comentario_final = comentario_usuario.strip()

                    hoja_mes = sheet.worksheet(hoja_mes_nombre)
                    ingreso = monto if tipo_movimiento == "Ingreso" else 0
                    egreso = monto if tipo_movimiento == "Egreso" else 0
                    
                    nueva_fila = [fecha_str, tipo_transaccion, ingreso, egreso, "", links_comprobantes_final, comentario_final]
                    
                    todas_las_filas = hoja_mes.get_all_values()
                    index_a_insertar = 4  
                    for i, fila in enumerate(todas_las_filas[1:]):
                        try:
                            if len(fila) > 0 and fila[0].strip():
                                fecha_fila_dt = datetime.strptime(fila[0], "%d/%m/%Y").date()
                                if fecha >= fecha_fila_dt:
                                    index_a_insertar = i + 3  
                        except (ValueError, IndexError):
                            continue
                    
                    hoja_mes.insert_row(nueva_fila, index=index_a_insertar, value_input_option='USER_ENTERED')

                    if es_pago_lobato and todos_los_registros:
                        hoja_principal = sheet.worksheet("Mensualidades")
                        headers = todos_los_registros[0]
                        
                        fila_lobato = None
                        datos_lobato = []
                        nombre_input_limpio = quitar_tildes(nombre_final)
                        
                        for i, fila in enumerate(todos_los_registros):
                            if len(fila) > 1 and quitar_tildes(fila[1]) == nombre_input_limpio:
                                fila_lobato = i + 1
                                datos_lobato = fila
                                break
                        
                        if fila_lobato:
                            celdas_a_actualizar = []

                            paga_inscripcion = tipo_transaccion in ["Inscripción", "Cuota e Inscripción"]
                            if paga_inscripcion: 
                                if 'Inscr.' in headers:
                                    idx_inscr = headers.index('Inscr.') + 1
                                    rango_inscr = rowcol_to_a1(fila_lobato, idx_inscr)
                                    celdas_a_actualizar.append({'range': rango_inscr, 'values': [[True]]})
                            
                            if num_cuotas_final > 0:
                                meses_orden = ['Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'ene']
                                cuotas_assigned = 0
                                for m_abrev in meses_orden:
                                    if cuotas_assigned >= num_cuotas_final: 
                                        break
                                    try:
                                        idx_mes = headers.index(m_abrev)
                                        valor_actual = datos_lobato[idx_mes].strip() if idx_mes < len(datos_lobato) else ""
                                        if valor_actual in ["", "0", "$0", "-", "FALSE"]:
                                            rango_mes = rowcol_to_a1(fila_lobato, idx_mes + 1)
                                            celdas_a_actualizar.append({'range': rango_mes, 'values': [[1]]})
                                            cuotas_assigned += 1
                                    except ValueError:
                                        continue
                            
                            if celdas_a_actualizar:
                                hoja_principal.batch_update(celdas_a_actualizar, value_input_option='USER_ENTERED')
                    
                    st.session_state['ultima_fecha'] = fecha
                    st.session_state['ultimo_tipo_trans'] = tipo_transaccion
                    
                    st.session_state['ultimo_registro'] = comentario_final
                    st.session_state['mostrar_globos'] = True
                    st.session_state.form_id += 1
                    
                    st.cache_data.clear() 
                    st.rerun()
                    
                except Exception as err:
                    st.error(f"❌ Ocurrió un error al guardar los datos: {err}")

# ==========================================
# PESTAÑA 2: ESTADÍSTICAS Y BUSCADOR
# ==========================================
with tab2:
    st.header("📊 Panel de Consultas y Estadísticas")
    st.write("Sapea los datos acumulados y encuentra información al toque.")

    if not todos_los_registros:
        st.warning("No hay datos disponibles para mostrar estadísticas. Revisa la conexión.")
    else:
        def limpiar_plata(valor):
            val_str = str(valor).replace('$', '').strip()
            if val_str in ['', '-', 'nan', 'None']:
                return 0
            for terminacion in [',00', '.00', ',0', '.0']:
                if val_str.endswith(terminacion):
                    val_str = val_str[:-len(terminacion)]
            val_str = val_str.replace('.', '').replace(',', '')
            try:
                return int(val_str)
            except:
                return 0

        st.subheader("📈 Balance de Caja Acumulado (Evolución de Saldos)")
        datos_anuales = obtener_datos_busqueda()
        if datos_anuales:
            cleaned_data = []
            for fila in datos_anuales:
                if len(fila) >= 7:
                    fecha = fila[0]
                    tipo = fila[1]
                    ingreso = fila[2]
                    egreso = fila[3]
                    saldo = fila[4] if len(fila) > 4 else "0"
                    comentario = fila[6] if len(fila) > 6 else ""
                    mes = fila[-1]
                    cleaned_data.append([fecha, tipo, ingreso, egreso, saldo, comentario, mes])
                    
            df_anual = pd.DataFrame(cleaned_data, columns=["Fecha", "Tipo", "Ingreso", "Egreso", "Saldo", "Comentario", "Mes"])
            
            for col in ['Ingreso', 'Egreso', 'Saldo']:
                df_anual[col] = df_anual[col].apply(limpiar_plata)
            
            orden_meses = ['Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'ene']
            mes_grafico_sel = st.selectbox("📅 Selecciona el mes para desplegar el gráfico:", orden_meses, index=1, key="mes_grafico_principal")
            
            df_mes = df_anual[df_anual['Mes'] == mes_grafico_sel].copy().reset_index(drop=True)
            
            if not df_mes.empty:
                es_inicial = df_mes['Tipo'].str.lower().str.contains('inicial') | df_mes['Comentario'].str.lower().str.contains('inicial') | df_mes['Tipo'].str.lower().str.contains('inicio')
                df_ini_row = df_mes[es_inicial]
                
                if not df_ini_row.empty:
                    saldo_inicial = df_ini_row.iloc[0]['Ingreso'] if df_ini_row.iloc[0]['Ingreso'] > 0 else df_ini_row.iloc[0]['Saldo']
                    df_calc = df_mes[~es_inicial].copy().reset_index(drop=True)
                else:
                    saldo_inicial = 0
                    df_calc = df_mes.copy()
                
                saldos_linea = [saldo_inicial]
                monto_actual = saldo_inicial
                
                for idx, row in df_calc.iterrows():
                    monto_actual = monto_actual + row['Ingreso'] - row['Egreso']
                    saldos_linea.append(monto_actual)
                
                fechas_tooltip = ["Inicio"] + list(df_calc['Fecha'])
                detalles_tooltip = ["Saldo Inicial"] + list(df_calc['Tipo'] + " - " + df_calc['Comentario'].str.slice(0, 30))
                
                df_grafico_final = pd.DataFrame({
                    "Saldo en Caja": saldos_linea,
                    "Fecha Mov.": fechas_tooltip,
                    "Detalle": detalles_tooltip
                })
                
                st.line_chart(df_grafico_final, y="Saldo en Caja", x=None) 
                st.caption("💡 El eje X representa el N° correlativo de movimientos en orden cronológico (0 = Inicio de mes).")
                
                nombre_visible_metric = "Enero" if mes_grafico_sel == "ene" else mes_grafico_sel
                st.metric(label=f"💰 Saldo Final en Planilla ({nombre_visible_metric})", value=f"${monto_actual:,.0f}".replace(",", "."))
            else:
                nombre_visible_info = "Enero" if mes_grafico_sel == "ene" else mes_grafico_sel
                st.info(f"Aún no hay transacciones registradas en la planilla para el mes de {nombre_visible_info}.")

        st.write("---")
        headers_m = todos_los_registros[0]
        
        metodo_busqueda = st.selectbox("🔍 Selecciona el método de búsqueda:", [
            "Buscar Miembro", 
            "Buscar Tipo de Transacción",
            "Buscar por Motivo/Comentario"
        ], key="metodo_busqueda_key")

        if metodo_busqueda == "Buscar Miembro":
            lobato_seleccionado = st.selectbox(f"👦 Elige un {TEXTO_INDIVIDUAL} de la {NOM_UNIDAD}:", ["-- Selecciona una opción --"] + lista_nombres_lobatos, key="busqueda_lobato_stats")
            
            if lobato_seleccionado != "-- Selecciona una opción --":
                st.subheader(f"📋 Estado de Cuenta: {lobato_seleccionado}")
                
                fila_n = None
                nombre_clean = quitar_tildes(lobato_seleccionado)
                for fila in todos_los_registros:
                    if len(fila) > 1 and quitar_tildes(fila[1]) == nombre_clean:
                        fila_n = fila
                        break
                
                if fila_n:
                    if 'Inscr.' in headers_m:
                        idx_i = headers_m.index('Inscr.')
                        status_inscr = fila_n[idx_i].strip().upper() if idx_i < len(fila_n) else ""
                        if status_inscr in ["TRUE", "1"]:
                            st.markdown("**Inscripción Inicial:** 🟢 Registrada / Al día")
                        else:
                            st.markdown("**Inscripción Inicial:** 🔴 Pendiente / No Registrada")
                    
                    st.write("#### Visualización de Cuotas:")
                    meses_a_mostrar = ['Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'ene']
                    cols_meses = st.columns(5)
                    
                    for idx_m, m_abrev in enumerate(meses_a_mostrar):
                        col_target = cols_meses[idx_m % 5]
                        with col_target:
                            nombre_visible = "Ene" if m_abrev == "ene" else m_abrev
                            st.write(f"**{nombre_visible}**")
                            if m_abrev in headers_m:
                                idx_cell = headers_m.index(m_abrev)
                                idx_monto = idx_cell + 1
                                val_monto = fila_n[idx_monto].strip() if idx_monto < len(fila_n) else ""
                                if val_monto and val_monto not in ["", "0", "$0", "-", "$ -"]:
                                    st.markdown("🟢 Al día")
                                else:
                                    st.markdown("🔴 Pendiente")
                            else:
                                st.markdown("🔴 No Disp.")

        elif metodo_busqueda == "Buscar Tipo de Transacción":
            tipo_sel = st.selectbox("💰 Selecciona el tipo de transacción:", 
                                    ["Cuota", "Inscripción", "Transferencia", "Devolución", "Compra", "Depósito", "Donación"], 
                                    key="tipo_trans_stats")
            
            if tipo_sel == "Cuota":
                st.subheader(f"🗓️ Cuotas de la {NOM_UNIDAD} por Mes")
                meses_a_contar = ['Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'ene']
                total_lobatos = len(lista_nombres_lobatos) 
                
                for m_abrev in meses_a_contar:
                    nombre_visible = "Ene" if m_abrev == "ene" else m_abrev
                    if m_abrev in headers_m:
                        idx_mes = headers_m.index(m_abrev)
                        idx_monto = idx_mes + 1
                        contador_al_dia = sum(1 for fila in todos_los_registros[2:] 
                                              if idx_monto < len(fila) and fila[idx_monto].strip() and fila[idx_monto].strip() not in ["", "0", "$0", "-", "$ -"])
                        st.write(f"🔹 **{nombre_visible}**: {contador_al_dia} de {total_lobatos} {TEXTO_PLURAL} al día.")
            else:
                st.subheader(f"📋 Listado de: {tipo_sel}")
                registros_encontrados = []
                cliente_sheets = autenticar()
                sheet = cliente_sheets.open_by_key(SPREADSHEET_ID)
                meses_map = ['Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic', 'ene']
                
                for mes in meses_map:
                    try:
                        hoja = sheet.worksheet(mes)
                        data = hoja.get_all_values()
                        for fila in data[1:]:
                            if tipo_sel.lower() in fila[1].lower():
                                registros_encontrados.append({"Mes": "Enero" if mes == "ene" else mes, "Fecha": fila[0], "Ingreso": fila[2], "Egreso": fila[3], "Comentario": fila[6]})
                    except: continue
                
                if registros_encontrados:
                    df_res_b = pd.DataFrame(registros_encontrados)
                    st.table(df_res_b)
                else:
                    st.info(f"No se encontraron registros que contengan '{tipo_sel}'.")
                    
        elif metodo_busqueda == "Buscar por Motivo/Comentario":
            keyword = st.text_input("📝 Escribe la palabra a buscar (ej: curanto):").strip().lower()
            
            if keyword:
                st.subheader(f"🔎 Resultados para: '{keyword}'")
                todos_los_datos = obtener_datos_busqueda()
                resultados_busqueda = []
                
                for fila in todos_los_datos:
                    comentario_fila = str(fila[6]).lower() if len(fila) > 6 else ""
                    if keyword in comentario_fila:
                        resultados_busqueda.append({
                            "Mes": "Enero" if fila[7] == "ene" else fila[7], 
                            "Fecha": fila[0],
                            "Tipo": fila[1],
                            "Ingreso": fila[2],
                            "Egreso": fila[3],
                            "Comentario": fila[6]
                        })
                
                if len(resultados_busqueda) > 0:
                    df_res_c = pd.DataFrame(resultados_busqueda)
                    df_res_c['Ingreso'] = df_res_c['Ingreso'].apply(limpiar_plata)
                    df_res_c['Egreso'] = df_res_c['Egreso'].apply(limpiar_plata)
                    st.table(df_res_c)
                else:
                    st.info("No se encontraron registros que contengan esa palabra.")