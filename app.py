import streamlit as st
import pandas as pd
from supabase import create_client, Client
import os
import time
import json
import io
import textwrap
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv

# Importación segura de librerías opcionales
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# Helper para zona horaria (Lima/Bogotá UTC-5)
def get_lima_now():
    return datetime.now(timezone.utc) - timedelta(hours=5)

def generate_word_letter(texto_completo, firma_resp):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    # Membrete simple
    header = doc.sections[0].header
    htable = header.add_table(1, 2, width=Inches(6))
    htable.autofit = False
    htable.columns[0].width = Inches(3)
    htable.columns[1].width = Inches(3)
    
    # Contenido del cuerpo
    for paragraph in texto_completo.split('\n'):
        if paragraph.strip():
            p = doc.add_paragraph(paragraph.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    # Firma
    doc.add_paragraph("\n\n")
    p_firma = doc.add_paragraph(firma_resp)
    p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cargo = doc.add_paragraph("Responsable")
    p_cargo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Cargar variables del archivo .env
env_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(env_path):
    load_dotenv(env_path)
else:
    load_dotenv()

# Configuración de la página
st.set_page_config(
    page_title="Control Horas - ER",
    page_icon="⏱️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inicialización de Supabase
@st.cache_resource
def get_supabase():
    url = st.secrets.get("SUPABASE_URL") or os.getenv("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_KEY") or os.getenv("SUPABASE_KEY")
    service_key = st.secrets.get("SUPABASE_SERVICE_KEY") or os.getenv("SUPABASE_SERVICE_KEY")

    if not url or not key:
        st.error("❌ Configuración incompleta. Revisa los Secrets de Streamlit.")
        st.stop()

    def clean(v):
        return str(v).strip().strip('"').strip("'").strip() if v else None

    url, key, service_key = map(clean, [url, key, service_key])
    return create_client(url, service_key if service_key else key)

supabase = get_supabase()

# Estilos premium
st.markdown("""
<style>
    .main { background-color: #f5f3ff; }
    .stButton>button { background-color: #7c3aed; color: white; border-radius: 8px; font-weight: bold; width: 100%; }
    .stButton>button:hover { background-color: #6d28d9; border-color: #6d28d9; }
    .sidebar .sidebar-content { background-color: #ede9fe; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { 
        background-color: #f3f4f6; padding: 10px 20px; border-radius: 8px 8px 0 0;
    }
    .stTabs [aria-selected="true"] { background-color: #7c3aed !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# Lógica de Login
def login_user(email, password):
    try:
        response = supabase.auth.sign_in_with_password({"email": email, "password": password})
        if response.user:
            profile = supabase.table("profiles").select("*, roles(name)").eq("id", response.user.id).single().execute()
            p_data = profile.data
            if not p_data.get('is_active', False):
                st.error("🚫 Usuario desactivado. Por favor contacte al administrador.")
                return
            st.session_state.user = response.user
            st.session_state.profile = p_data
            is_admin_check = p_data.get('is_admin', False)
            acc_type = p_data.get('account_type', '')
            st.session_state.is_admin = is_admin_check or (acc_type == "Administrador")
            st.rerun()
    except Exception as e:
        st.error("Error de acceso: Verifica tus datos.")

# Sidebar y Título
st.title("💜 Control Horas - ER")

def mostrar_registro_tiempos():
    st.header("Registro de Tiempos")
    
    # Manejo de mensajes persistentes tras rerun
    if 'success_msg' in st.session_state:
        st.toast(st.session_state.success_msg, icon="✅")
        del st.session_state.success_msg
    
    # Manejo de keys para borrado
    if 'form_key_suffix' not in st.session_state: st.session_state.form_key_suffix = 0
    
    # 1. Selección de Cliente y Proyecto
    clientes = supabase.table("clients").select("id, name").order("name").execute()
    if not clientes.data:
        st.info("Aún no hay clientes registrados.")
        return
        
    client_map = {c['name']: c['id'] for c in clientes.data}
    
    # Recuperar cliente del timer activo si existe
    index_cliente = 0
    if 'active_client_name' in st.session_state and st.session_state.active_client_name in client_map:
        index_cliente = (list(client_map.keys()).index(st.session_state.active_client_name)) + 1

    cliente_sel = st.selectbox("Seleccionar Cliente", ["---"] + list(client_map.keys()), index=index_cliente, key=f"cli_{st.session_state.form_key_suffix}")
    
    if cliente_sel == "---": return

    proyectos = supabase.table("projects").select("id, name, currency").eq("client_id", client_map[cliente_sel]).order("name").execute()
    if not proyectos.data:
        st.warning(f"Sin proyectos para {cliente_sel}.")
        return
        
    proj_map = {p['name']: p['id'] for p in proyectos.data}
    proj_currency = {p['id']: p['currency'] for p in proyectos.data}
    
    index_proj = 0
    if 'active_project_name' in st.session_state and st.session_state.active_project_name in proj_map:
        index_proj = list(proj_map.keys()).index(st.session_state.active_project_name)

    proyecto_sel = st.selectbox("Seleccionar Proyecto", list(proj_map.keys()), index=index_proj, key=f"pro_{st.session_state.form_key_suffix}")
    target_user_id = st.session_state.user.id
    fecha_sel = get_lima_now()
    p_id = proj_map[proyecto_sel]
    moneda = proj_currency[p_id]
    
    st.info(f"Proyecto: **{proyecto_sel}** | Moneda: **{moneda}**")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        fecha_sel = st.date_input("Fecha", value=get_lima_now(), max_value=get_lima_now(), key=f"fec_{st.session_state.form_key_suffix}")
    with col_u2:
        if st.session_state.is_admin:
            usuarios_res = supabase.table("profiles").select("id, full_name").eq("is_active", True).execute()
            user_map = {u['full_name']: u['id'] for u in usuarios_res.data}
            usuario_para = st.selectbox("Registrar para", list(user_map.keys()), index=list(user_map.values()).index(st.session_state.user.id) if st.session_state.user.id in user_map.values() else 0, key=f"user_sel_{st.session_state.form_key_suffix}")
            target_user_id = user_map[usuario_para]
        else:
            target_user_id = st.session_state.user.id
        st.write(f"Usuario: **{st.session_state.profile['full_name']}**")

    # VALIDACIÓN DE TARIFA
    if not target_user_id:
        st.warning("Sesión no válida o usuario no detectado. Por favor, inicie sesión nuevamente.")
        return

    profile_info = supabase.table("profiles").select("role_id").eq("id", target_user_id).single().execute()
    role_id = profile_info.data['role_id']
    rate_q = supabase.table("project_rates").select("rate").eq("project_id", p_id).eq("role_id", role_id).execute()
    
    current_rate_val = float(rate_q.data[0]['rate']) if rate_q.data else 0.0
    
    if current_rate_val <= 0:
        if st.session_state.is_admin:
            st.warning(f"⚠️ **Atención**: No se han definido tarifas para el rol en este proyecto.")
        can_register = True
    else:
        if st.session_state.is_admin:
            st.success(f"Tarifa detectada: **{current_rate_val} {moneda}/h**")
        can_register = True

    # Valor por defecto para descripción y facturabilidad
    def_desc = st.session_state.get('active_timer_description', '')
    def_fact = st.session_state.get('active_timer_billable', True)

    descripcion = st.text_area("Detalle del trabajo", value=def_desc, placeholder="¿Qué hiciste?", key=f"desc_{st.session_state.form_key_suffix}")
    es_facturable = st.checkbox("¿Es facturable?", value=def_fact, key=f"fact_{st.session_state.form_key_suffix}")
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    # 2. INGRESO MANUAL
    with col1:
        st.subheader("📝 Ingreso Manual")
        t_inicio_str = st.text_input("Hora Inicio (HH:mm)", value="08:00", key=f"hi_{st.session_state.form_key_suffix}")
        t_fin_str = st.text_input("Hora Final (HH:mm)", value="09:00", key=f"hf_{st.session_state.form_key_suffix}")
        
        if st.button("Registrar Manualmente", disabled=not can_register):
            try:
                t1_dt = datetime.strptime(t_inicio_str, "%H:%M")
                t2_dt = datetime.strptime(t_fin_str, "%H:%M")
                tz_local = timezone(timedelta(hours=-5))
                t1 = datetime.combine(fecha_sel, t1_dt.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                t2 = datetime.combine(fecha_sel, t2_dt.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                
                if t2 <= t1:
                    st.error("La hora final debe ser posterior a la inicial.")
                else:
                    total_min = int((t2 - t1).total_seconds() / 60)
                    supabase.table("time_entries").insert({
                        "profile_id": target_user_id,
                        "project_id": p_id,
                        "description": descripcion,
                        "start_time": t1.isoformat(),
                        "end_time": t2.isoformat(),
                        "total_minutes": total_min,
                        "is_billable": es_facturable
                    }).execute()
                    
                    # LIMPIEZA TOTAL TRAS GUARDAR MANUAL
                    st.session_state.active_timer_description = ''
                    st.session_state.active_timer_billable = True
                    st.session_state.active_client_name = None
                    st.session_state.active_project_name = None
                    
                    st.session_state.success_msg = f"✅ Guardado con éxito ({t_inicio_str} a {t_fin_str})."
                    st.session_state.form_key_suffix += 1
                    st.rerun()
            except ValueError:
                st.error("Formato inválido. Use HH:mm (ej: 08:33)")

    # 3. CRONÓMETRO (Persistente)
    with col2:
        st.subheader("⏱️ Cronómetro")
        if 'timer_running' not in st.session_state: st.session_state.timer_running = False
        if 'timer_start' not in st.session_state: st.session_state.timer_start = None
        if 'total_elapsed' not in st.session_state: st.session_state.total_elapsed = 0
        if 'active_timer_id' not in st.session_state: st.session_state.active_timer_id = None

        # Sincronización inicial
        if st.session_state.active_timer_id is None and st.session_state.user:
            try:
                timer_q = supabase.table("active_timers").select("*, projects(name, client_id, clients(name))").eq("user_id", st.session_state.user.id).execute()
                if timer_q.data:
                    t_data = timer_q.data[0]
                    st.session_state.active_timer_id = t_data['id']
                    st.session_state.timer_running = t_data['is_running']
                    st.session_state.timer_start = pd.to_datetime(t_data['start_time']).replace(tzinfo=None)
                    st.session_state.total_elapsed = t_data['total_elapsed_seconds']
                    # Persistencia de formulario
                    st.session_state.active_project_id = t_data['project_id']
                    st.session_state.active_project_name = t_data['projects']['name']
                    st.session_state.active_client_name = t_data['projects']['clients']['name']
                    st.session_state.active_timer_description = t_data.get('description', '')
                    st.session_state.active_timer_billable = t_data.get('is_billable', True)
                    st.rerun()
            except Exception as e:
                st.error(f"Error sincronizando cronómetro: {str(e)}")

        # CONTROL DE VISIBILIDAD DE CRONÓMETRO ACTIVO
        timer_is_for_current_proj = False
        if st.session_state.active_timer_id and st.session_state.get('active_project_id') == p_id:
            timer_is_for_current_proj = True

        if st.session_state.timer_running and timer_is_for_current_proj:
            now_lima = get_lima_now().replace(tzinfo=None)
            actual_elapsed = st.session_state.total_elapsed + (now_lima - st.session_state.timer_start).total_seconds()
            
            hrs, rem = divmod(int(actual_elapsed), 3600)
            mins, secs = divmod(rem, 60)
            st.metric("⏱️ EN VIVO (Cronometrando)", f"{hrs:02d}:{mins:02d}:{secs:02d}")

            col_t1, col_t2 = st.columns(2)
            with col_t1:
                if st.button("⏸️ Pausar"):
                    new_elapsed = st.session_state.total_elapsed + (get_lima_now().replace(tzinfo=None) - st.session_state.timer_start).total_seconds()
                    st.session_state.total_elapsed = new_elapsed
                    st.session_state.timer_running = False
                    if st.session_state.active_timer_id:
                        supabase.table("active_timers").update({
                            "is_running": False,
                            "total_elapsed_seconds": int(new_elapsed),
                            "description": descripcion,
                            "is_billable": es_facturable
                        }).eq("id", st.session_state.active_timer_id).execute()
                    st.rerun()
            
            with col_t2:
                if st.button("⏹️ Finalizar", disabled=not can_register):
                    t_now = get_lima_now().replace(tzinfo=None)
                    total_sec = st.session_state.total_elapsed + (t_now - st.session_state.timer_start).total_seconds()
                    total_min = int(total_sec // 60) + (1 if total_sec % 60 > 0 else 0)
                    
                    tz_local = timezone(timedelta(hours=-5))
                    t_start_local = st.session_state.timer_start - timedelta(seconds=st.session_state.total_elapsed)
                    start_dt = datetime.combine(fecha_sel, t_start_local.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                    end_dt = start_dt + timedelta(minutes=total_min)
                    
                    supabase.table("time_entries").insert({
                        "profile_id": target_user_id,
                        "project_id": p_id,
                        "description": descripcion,
                        "start_time": start_dt.isoformat(),
                        "end_time": end_dt.isoformat(),
                        "total_minutes": total_min,
                        "is_billable": es_facturable
                    }).execute()
                    
                    if st.session_state.active_timer_id:
                        supabase.table("active_timers").delete().eq("id", st.session_state.active_timer_id).execute()

                    # LIMPIEZA TOTAL
                    st.session_state.timer_running = False
                    st.session_state.total_elapsed = 0
                    st.session_state.timer_start = None
                    st.session_state.active_timer_id = None
                    st.session_state.active_project_id = None
                    st.session_state.active_project_name = None
                    st.session_state.active_client_name = None
                    st.session_state.active_timer_description = ''
                    st.session_state.active_timer_billable = True
                    
                    st.session_state.success_msg = "✅ Registro con cronómetro guardado y pantalla limpiada."
                    st.session_state.form_key_suffix += 1
                    st.rerun()
        else:
            if st.session_state.total_elapsed > 0 and timer_is_for_current_proj:
                hrs, rem = divmod(int(st.session_state.total_elapsed), 3600)
                mins, secs = divmod(rem, 60)
                st.metric("Pausado", f"{hrs:02d}:{mins:02d}:{secs:02d}")
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    if st.button("▶️ Continuar"):
                        st.session_state.timer_start = get_lima_now().replace(tzinfo=None)
                        st.session_state.timer_running = True
                        if st.session_state.active_timer_id:
                            supabase.table("active_timers").update({
                                "is_running": True,
                                "start_time": st.session_state.timer_start.isoformat(),
                                "description": descripcion,
                                "is_billable": es_facturable
                            }).eq("id", st.session_state.active_timer_id).execute()
                        st.rerun()
                with col_p2:
                    if st.button("🗑️ Descartar"):
                        if st.session_state.active_timer_id:
                            supabase.table("active_timers").delete().eq("id", st.session_state.active_timer_id).execute()
                        
                        st.session_state.timer_running = False
                        st.session_state.total_elapsed = 0
                        st.session_state.timer_start = None
                        st.session_state.active_timer_id = None
                        st.session_state.active_project_id = None
                        st.session_state.active_project_name = None
                        st.session_state.active_client_name = None
                        st.session_state.active_timer_description = ''
                        st.session_state.active_timer_billable = True
                        
                        st.session_state.form_key_suffix += 1
                        st.rerun()
            elif st.session_state.active_timer_id and not timer_is_for_current_proj:
                st.warning(f"⚠️ Tienes un cronómetro activo en: **{st.session_state.get('active_project_name', 'otro proyecto')}**")
                if st.button("Ir al proyecto activo"):
                    st.session_state.active_client_name = st.session_state.get('active_client_name')
                    st.session_state.active_project_name = st.session_state.get('active_project_name')
                    st.rerun()
            else:
                if st.button("▶️ Iniciar Cronómetro", disabled=not can_register):
                    st.session_state.timer_start = get_lima_now().replace(tzinfo=None)
                    st.session_state.timer_running = True
                    try:
                        resp = supabase.table("active_timers").insert({
                            "user_id": st.session_state.user.id,
                            "project_id": p_id,
                            "start_time": st.session_state.timer_start.isoformat(),
                            "description": descripcion,
                            "is_billable": es_facturable,
                            "is_running": True
                        }).execute()
                        if resp.data:
                            st.session_state.active_timer_id = resp.data[0]['id']
                            st.session_state.active_project_id = p_id
                    except Exception as e:
                        st.error(f"Error iniciando cronómetro: {str(e)}")
                    st.rerun()

    # 4. TABLA DE HISTORIAL
    st.markdown("---")
    st.subheader("📋 Historial de Horas")
    
    # Sidebar filtros
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros de Historial")
    hoy = get_lima_now().date()
    lunes = hoy - timedelta(days=hoy.weekday())
    
    date_range = st.sidebar.date_input("Rango de fechas", value=(lunes, hoy))
    
    query = supabase.table("time_entries").select("*, projects(name, currency), profiles(full_name)").eq("profile_id" if not st.session_state.is_admin else "profile_id", target_user_id)
    
    if len(date_range) == 2:
        start_date = datetime.combine(date_range[0], datetime.min.time()).replace(tzinfo=timezone.utc).isoformat()
        end_date = datetime.combine(date_range[1], datetime.max.time()).replace(tzinfo=timezone.utc).isoformat()
        query = query.gte("start_time", start_date).lte("start_time", end_date)
    
    # ORDEN DESCENDENTE: Más reciente primero
    res = query.order("start_time", desc=True).execute()
    
    if res.data:
        df = pd.DataFrame(res.data)
        df['Fecha'] = pd.to_datetime(df['start_time']).dt.strftime('%d/%m/%Y')
        df['Proyecto'] = df['projects'].apply(lambda x: x['name'])
        df['Consultor'] = df['profiles'].apply(lambda x: x['full_name'])
        df['Horas'] = (df['total_minutes'] / 60).apply(lambda x: f"{int(x)}h {int((x*60)%60)}m")
        
        cols = ['Fecha', 'Consultor', 'Proyecto', 'description', 'Horas', 'is_billable']
        st.dataframe(df[cols], use_container_width=True)
    else:
        st.write("No hay registros en este rango.")

def main():
    if 'user' not in st.session_state:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st.subheader("Iniciar Sesión")
            email = st.text_input("Correo electrónico")
            pw = st.text_input("Contraseña", type="password")
            if st.button("Entrar"):
                login_user(email, pw)
    else:
        # Menú lateral
        with st.sidebar:
            st.write(f"Conectado como: **{st.session_state.profile['full_name']}**")
            menu = ["Cronómetro / Manual", "Reportes / Liquidación", "Carga Masiva"]
            choice = st.radio("Menú", menu)
            
            if st.button("Cerrar Sesión"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()

        if choice == "Cronómetro / Manual":
            mostrar_registro_tiempos()
        elif choice == "Reportes / Liquidación":
            mostrar_reportes()
        elif choice == "Carga Masiva":
            mostrar_carga_masiva()

def mostrar_carga_masiva():
    st.header("Carga Masiva de Datos")
    tabs = st.tabs(["Registros de Tiempo", "Clientes", "Proyectos", "Tarifas"])
    
    with tabs[0]:
        st.subheader("Carga Masiva de Registros de Tiempo")
        st.info("Formato requerido: Fecha | Responsable | Cliente | Proyecto | Detalle | Hora Inicio | Hora Final")
        
        if HAS_OPENPYXL:
            template_time = pd.DataFrame({
                'Fecha': ['06.02-2026', '06.02-2026'],
                'Responsable': ['Juan Pérez', 'María García'],
                'Cliente': ['Cliente A', 'Cliente B'],
                'Proyecto': ['Proyecto X', 'Proyecto Y'],
                'Detalle': ['Reunión de planificación', 'Desarrollo de módulo'],
                'Hora Inicio': ['09:00', '14:00'],
                'Hora Final': ['11:30', '17:00']
            })
            buffer_template = io.BytesIO()
            with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
                template_time.to_excel(writer, index=False, sheet_name='Registros')
            st.download_button("📥 Descargar Template", data=buffer_template.getvalue(), file_name="template_registros.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ La función de descarga de templates requiere 'openpyxl'. Por favor, instálela en requirements.txt y reinicie la app.")

        uploaded_file = st.file_uploader("Seleccionar archivo Excel", type=['xlsx'])
        if uploaded_file and HAS_OPENPYXL:
            try:
                df_upload = pd.read_excel(uploaded_file, engine='openpyxl')
                
                # Cargar mapas para validación
                users_res = supabase.table("profiles").select("id, full_name").execute()
                clients_res = supabase.table("clients").select("id, name").execute()
                
                users_map = {u['full_name']: u['id'] for u in users_res.data}
                clients_map = {c['name']: c['id'] for c in clients_res.data}
                
                errors = []
                valid_entries = []
                
                for idx, row in df_upload.iterrows():
                    try:
                        resp_name = row.get('Responsable')
                        if resp_name not in users_map:
                            errors.append(f"Fila {idx+2}: Usuario '{resp_name}' no encontrado")
                            continue
                        u_id = users_map[resp_name]
                        
                        cli_name = row.get('Cliente')
                        if cli_name not in clients_map:
                            errors.append(f"Fila {idx+2}: Cliente '{cli_name}' no encontrado")
                            continue
                        c_id = clients_map[cli_name]
                        
                        proj_name = row.get('Proyecto')
                        # Encontrar proyecto por nombre y cliente_id
                        p_res = supabase.table("projects").select("id").eq("client_id", c_id).eq("name", proj_name).execute()
                        if not p_res.data:
                            errors.append(f"Fila {idx+2}: Proyecto '{proj_name}' no coincide con el cliente")
                            continue
                        p_id = p_res.data[0]['id']
                        
                        # Fechas y Horas
                        fecha = pd.to_datetime(row.get('Fecha'))
                        tz_local = timezone(timedelta(hours=-5))
                        h1 = datetime.strptime(str(row.get('Hora Inicio')), "%H:%M")
                        h2 = datetime.strptime(str(row.get('Hora Final')), "%H:%M")
                        
                        t1 = datetime.combine(fecha.date(), h1.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                        t2 = datetime.combine(fecha.date(), h2.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                        
                        total_min = int((t2 - t1).total_seconds() / 60)
                        
                        valid_entries.append({
                            "profile_id": u_id,
                            "project_id": p_id,
                            "description": row.get('Detalle', ''),
                            "start_time": t1.isoformat(),
                            "end_time": t2.isoformat(),
                            "total_minutes": total_min,
                            "is_billable": True
                        })
                    except Exception as fe:
                        errors.append(f"Fila {idx+2}: Error de formato - {str(fe)}")

                if errors:
                    st.error("Errores encontrados:")
                    for e in errors[:10]: st.write(f"- {e}")
                    if len(errors) > 10: st.write("...")
                
                if valid_entries:
                    st.success(f"Se encontraron {len(valid_entries)} registros válidos.")
                    if st.button("🚀 Confirmar e Importar"):
                        with st.spinner("Importando..."):
                            for entry in valid_entries:
                                supabase.table("time_entries").insert(entry).execute()
                        st.success("¡Importación completada!")
                        st.rerun()
            except Exception as e:
                st.error(f"Error procesando archivo: {str(e)}")


# ... (Aquí irían mostrar_reportes y otras funciones, pero las restauraré de a pocos si son muy grandes) ...

def mostrar_reportes():
    st.header("Reportes y Liquidación")
    
    # 1. Selección de Cliente
    clientes = supabase.table("clients").select("id, name").order("name").execute()
    if not clientes.data:
        st.info("Aún no hay clientes registrados.")
        return
    cli_map = {c['name']: c['id'] for c in clientes.data}
    cli_name_sel = st.selectbox("Seleccionar Cliente (Reportes)", ["---"] + list(cli_map.keys()))
    
    if cli_name_sel == "---":
        st.info("Seleccione un cliente para ver reportes.")
        return
    
    cli_id = cli_map[cli_name_sel]

    # Rango de fechas
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Rango para Reporte")
    hoy = get_lima_now().date()
    ayer = hoy - timedelta(days=30)
    range_rep = st.sidebar.date_input("Periodo", value=(ayer, hoy))

    if len(range_rep) == 2:
        t1 = datetime.combine(range_rep[0], datetime.min.time()).replace(tzinfo=timezone.utc).isoformat()
        t2 = datetime.combine(range_rep[1], datetime.max.time()).replace(tzinfo=timezone.utc).isoformat()

        # Query de registros para este cliente
        res_rep = supabase.table("time_entries").select(
            "*, projects!inner(name, client_id, currency), profiles(full_name, role_id)"
        ).eq("projects.client_id", cli_id).gte("start_time", t1).lte("start_time", t2).order("start_time").execute()

        if res_rep.data:
            df_rep = pd.DataFrame(res_rep.data)
            df_rep['Fecha_dt'] = pd.to_datetime(df_rep['start_time'])
            df_rep['Fecha_str'] = df_rep['Fecha_dt'].dt.strftime('%d/%m/%Y')
            df_rep['Horas_num'] = df_rep['total_minutes'] / 60
            
            # Obtener tarifas para calcular montos
            projs_ids = df_rep['project_id'].unique().tolist()
            roles_ids = df_rep['profiles'].apply(lambda x: x['role_id']).unique().tolist()
            rates_res = supabase.table("project_rates").select("*").in_("project_id", projs_ids).in_("role_id", roles_ids).execute()
            rates_df = pd.DataFrame(rates_res.data) if rates_res.data else pd.DataFrame()

            def calc_monto(row):
                if not rates_df.empty:
                    match = rates_df[(rates_df['project_id'] == row['project_id']) & (rates_df['role_id'] == row['profiles']['role_id'])]
                    if not match.empty:
                        return row['Horas_num'] * float(match.iloc[0]['rate'])
                return 0.0

            df_rep['Total_Monto'] = df_rep.apply(calc_monto, axis=1)

            st.subheader("Control de Liquidación")
            
            # Pestañas: Carta, Anexo, Dashboard
            tab1, tab2, tab3 = st.tabs(["📄 Carta", "📊 Anexo Detallado", "📈 Dashboard"])
            
            with tab1:
                monedas_disp = df_rep['projects'].apply(lambda x: x['currency']).unique()
                moneda_liq = st.selectbox("Moneda para Liquidar", monedas_disp)
                
                # Filtrar por moneda
                df_liq = df_rep[df_rep['projects'].apply(lambda x: x['currency']) == moneda_liq].copy()
                total_monto_liq = df_liq['Total_Monto'].sum()
                total_horas_liq = df_liq['Horas_num'].sum()
                
                # Buscar liquidación existente
                liq_q = supabase.table("liquidations").select("*").eq("client_id", cli_id).eq("currency", moneda_liq).eq("status", "draft").execute()
                liq_data = liq_q.data[0] if liq_q.data else {}
                
                liquidation_number = liq_data.get('liquidation_number', '')
                liquidation_status = liq_data.get('status', 'nuevo')
                liquidation_id = liq_data.get('id')

                if liq_data:
                    st.write(f"**Liquidación Existente (Borrador)**: {liquidation_number}")
                
                # Notas especiales
                notas_especiales = st.text_area("Notas / Detractiones (Opcional)", value=liq_data.get('special_notes', ''))
                
                # Generar texto de la carta
                fecha_carta = get_lima_now().strftime("%d de %B de %Y")
                # Traducción simple de meses
                meses = {'January': 'Enero', 'February': 'Febrero', 'March': 'Marzo', 'April': 'Abril', 'May': 'Mayo', 'June': 'Junio', 'July': 'Julio', 'August': 'Agosto', 'September': 'Septiembre', 'October': 'Octubre', 'November': 'Noviembre', 'December': 'Diciembre'}
                for en, es in meses.items(): fecha_carta = fecha_carta.replace(en, es)
                
                letter_template = f"""San Isidro, {fecha_carta}

Señores:
{cli_name_sel}
Atención: Administración / Finanzas

Ref: Liquidación de Servicios por Honorarios Profesionales

De nuestra consideración:

Por intermedio de la presente, les hacemos llegar la liquidación correspondiente a los servicios profesionales prestados durante el periodo seleccionado.

El detalle del monto a liquidar es el siguiente:

- Total Horas: {total_horas_liq:.2f} h
- Moneda: {moneda_liq}
- Monto Total: {moneda_liq} {total_monto_liq:,.2f}

{notas_especiales if notas_especiales else ''}

Agradeciendo de antemano su atención, quedamos a su disposición para cualquier consulta.

Atentamente,

ERH Abogados"""
                
                txt_carta = st.text_area("Editar Contenido de la Carta", value=letter_template, height=300)
                firma_def = st.text_input("Firma Responsable", value=st.session_state.profile.get('full_name', ''))

                col_save1, col_save2, col_save3 = st.columns(3)
                
                with col_save1:
                    if st.button("💾 Guardar Borrador"):
                        try:
                            data = {
                                "client_id": cli_id,
                                "currency": moneda_liq,
                                "total_amount": float(total_monto_liq),
                                "total_hours": float(total_horas_liq),
                                "status": "draft",
                                "special_notes": notas_especiales,
                                "generated_by": st.session_state.user.id
                            }
                            if not liquidation_id:
                                # Generar número auto-incremental simple
                                last = supabase.table("liquidations").select("liquidation_number").order("id", desc=True).limit(1).execute()
                                next_num = 1
                                if last.data:
                                    try: next_num = int(last.data[0]['liquidation_number'].split('-')[1]) + 1
                                    except: next_num = 1
                                data["liquidation_number"] = f"LIQ-{next_num:04d}"
                                supabase.table("liquidations").insert(data).execute()
                            else:
                                supabase.table("liquidations").update(data).eq("id", liquidation_id).execute()
                            st.success("Borrador guardado.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
                            
                with col_save2:
                    if liquidation_id and st.button("📧 Marcar como Enviada"):
                        supabase.table("liquidations").update({"status": "sent", "sent_at": get_lima_now().isoformat()}).eq("id", liquidation_id).execute()
                        st.success("Estatus: ENVIADO")
                        st.rerun()
                
                with col_save3:
                    if liquidation_id and st.button("💰 Marcar como Pagada"):
                        supabase.table("liquidations").update({"status": "paid", "paid_at": get_lima_now().isoformat()}).eq("id", liquidation_id).execute()
                        st.success("Estatus: PAGADO")
                        st.rerun()

                st.markdown("---")
                if HAS_DOCX:
                    docx_bytes = generate_word_letter(txt_carta, firma_def)
                    st.download_button("📄 Descargar Word (.docx)", data=docx_bytes, file_name=f"Liquidacion_{cli_name_sel}_{moneda_liq}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.warning("Instale 'python-docx' para descargar en Word.")

            with tab2:
                st.subheader(f"Anexo: Detalle ({moneda_liq})")
                df_anexo = df_rep[df_rep['projects'].apply(lambda x: x['currency']) == moneda_liq].copy()
                if not df_anexo.empty:
                    df_view = df_anexo[['Fecha_str', 'projects', 'description', 'Horas_num', 'Total_Monto']]
                    df_view.columns = ['Fecha', 'Proyecto', 'Descripción', 'Horas', 'Monto']
                    df_view['Proyecto'] = df_view['Proyecto'].apply(lambda x: x['name'])
                    st.dataframe(df_view, use_container_width=True)
                    
                    if HAS_OPENPYXL:
                        buffer_xls = io.BytesIO()
                        with pd.ExcelWriter(buffer_xls, engine='openpyxl') as writer:
                            df_view.to_excel(writer, index=False, sheet_name='Anexo')
                        st.download_button("📥 Descargar Anexo (Excel)", data=buffer_xls.getvalue(), file_name=f"Anexo_{cli_name_sel}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.info("No hay datos para esta moneda.")

            with tab3:
                st.subheader("Métricas por Proyecto")
                sum_proj = df_rep.groupby(df_rep['projects'].apply(lambda x: x['name'])).agg({'Horas_num': 'sum', 'Total_Monto': 'sum'})
                st.bar_chart(sum_proj['Horas_num'])
                st.write(sum_proj)

        else:
            st.info("No hay registros en el periodo seleccionado.")


if __name__ == "__main__":
    main()

# --- REFRESH DINÁMICO ---
if st.session_state.get('timer_running'):
    time.sleep(1)
    st.rerun()
