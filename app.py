import streamlit as st
import pandas as pd
from supabase import create_client, Client
import os
import time
import base64
import json
import io
import textwrap
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv

# Cargar variables del archivo .env buscando el archivo en la misma carpeta que este script
env_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(env_path):
    load_dotenv(env_path)
else:
    load_dotenv() # Fallback por si acaso

# Configuración de la página
st.set_page_config(
    page_title="Control Horas - ER",
    page_icon="⏱️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inicialización de Supabase con soporte para Nube
@st.cache_resource
def get_supabase():
    # 1. Intentar cargar desde Secrets o entorno
    url = st.secrets.get("SUPABASE_URL") or os.getenv("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_KEY") or os.getenv("SUPABASE_KEY")
    service_key = st.secrets.get("SUPABASE_SERVICE_KEY") or os.getenv("SUPABASE_SERVICE_KEY")

    if not url or not key:
        st.error("❌ Configuración incompleta. Revisa los Secrets de Streamlit.")
        st.stop()

    # 2. Limpieza de llaves
    def clean(v):
        return str(v).strip().strip('"').strip("'").strip() if v else None

    url, key, service_key = map(clean, [url, key, service_key])
    
    # Priorizar Service Key para administración
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
            # Fetch profile with account_type and role
            profile = supabase.table("profiles").select("*, roles(name)").eq("id", response.user.id).single().execute()
            p_data = profile.data
            
            # Validar si está activo
            if not p_data.get('is_active', False):
                st.error("🚫 Usuario desactivado. Por favor contacte al administrador.")
                return

            st.session_state.user = response.user
            st.session_state.profile = p_data
            # Administrador si tiene check o por fallback de tipo de cuenta
            is_admin_check = p_data.get('is_admin', False)
            acc_type = p_data.get('account_type', '')
            st.session_state.is_admin = is_admin_check or (acc_type == "Administrador")
            st.rerun()
    except Exception as e:
        st.error("Error de acceso: Verifica tus datos.")

# Sidebar y Título
st.title("💜 Control Horas - ER")

if 'user' not in st.session_state:
    st.session_state.user = None

# Función reutilizable para el Registro de Tiempos
def mostrar_registro_tiempos():
    st.header("⏳ Registro de Tiempos")
    
    # Manejo de mensajes persistentes tras rerun
    if 'success_msg' in st.session_state:
        st.success(st.session_state.success_msg)
        del st.session_state.success_msg
    
    # Manejo de keys para borrado
    if 'form_key_suffix' not in st.session_state: st.session_state.form_key_suffix = 0
    
    # 1. Selección de Cliente y Proyecto
    clientes = supabase.table("clients").select("id, name").order("name").execute()
    if not clientes.data:
        st.info("Aún no hay clientes registrados.")
        return
        
    client_map = {c['name']: c['id'] for c in clientes.data}
    cliente_sel = st.selectbox("Seleccionar Cliente", ["---"] + list(client_map.keys()), key=f"cli_{st.session_state.form_key_suffix}")
    
    if cliente_sel == "---": return

    proyectos = supabase.table("projects").select("id, name, currency").eq("client_id", client_map[cliente_sel]).order("name").execute()
    if not proyectos.data:
        st.warning(f"Sin proyectos para {cliente_sel}.")
        return
        
    proj_map = {p['name']: p['id'] for p in proyectos.data}
    proj_currency = {p['id']: p['currency'] for p in proyectos.data}
    proyecto_sel = st.selectbox("Seleccionar Proyecto", list(proj_map.keys()), key=f"pro_{st.session_state.form_key_suffix}")
    # Variables para alcance (Scope)
    target_user_id = st.session_state.user.id
    fecha_sel = datetime.today()
    p_id = proj_map[proyecto_sel]
    moneda = proj_currency[p_id]
    
    st.info(f"Proyecto: **{proyecto_sel}** | Moneda: **{moneda}**")
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        fecha_sel = st.date_input("Fecha", value=datetime.today(), max_value=datetime.today(), key=f"fec_{st.session_state.form_key_suffix}")
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
    profile_info = supabase.table("profiles").select("role_id").eq("id", target_user_id).single().execute()
    role_id = profile_info.data['role_id']
    rate_q = supabase.table("project_rates").select("rate").eq("project_id", p_id).eq("role_id", role_id).execute()
    
    current_rate_val = float(rate_q.data[0]['rate']) if rate_q.data else 0.0
    
    if current_rate_val <= 0:
        if st.session_state.is_admin:
            st.warning(f"⚠️ **Atención**: No se han definido tarifas para el rol en este proyecto.")
        can_register = True # Permitir registrar incluso sin tarifa (será 0)
    else:
        if st.session_state.is_admin:
            st.success(f"Tarifa detectada: **{current_rate_val} {moneda}/h**")
        can_register = True

    descripcion = st.text_area("Detalle del trabajo", placeholder="¿Qué hiciste?", key=f"desc_{st.session_state.form_key_suffix}")
    es_facturable = st.checkbox("¿Es facturable?", value=True, key=f"fact_{st.session_state.form_key_suffix}")
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    # 2. INGRESO MANUAL (Restaurado a TEXTO para precisión HH:mm)
    with col1:
        st.subheader("📝 Ingreso Manual")
        t_inicio_str = st.text_input("Hora Inicio (HH:mm)", value="08:00", key=f"hi_{st.session_state.form_key_suffix}")
        t_fin_str = st.text_input("Hora Final (HH:mm)", value="09:00", key=f"hf_{st.session_state.form_key_suffix}")
        
        if st.button("Registrar Manualmente", disabled=not can_register):
            try:
                # Validar formatos
                t1_dt = datetime.strptime(t_inicio_str, "%H:%M")
                t2_dt = datetime.strptime(t_fin_str, "%H:%M")
                
                # Considerar UTC-5 (Bogotá/Lima) para el ingreso manual
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
                    st.session_state.success_msg = f"✅ Guardado con éxito ({t_inicio_str} a {t_fin_str})."
                    st.session_state.form_key_suffix += 1
                    st.rerun()
            except ValueError:
                st.error("Formato inválido. Use HH:mm (ej: 08:33)")

    # 3. CRONÓMETRO
    with col2:
        st.subheader("⏱️ Cronómetro")
        if 'timer_running' not in st.session_state: st.session_state.timer_running = False
        if 'timer_start' not in st.session_state: st.session_state.timer_start = None
        if 'total_elapsed' not in st.session_state: st.session_state.total_elapsed = 0

        if st.session_state.timer_running:
            actual_elapsed = st.session_state.total_elapsed + (datetime.now() - st.session_state.timer_start).total_seconds()
            hrs, rem = divmod(int(actual_elapsed), 3600)
            mins, secs = divmod(rem, 60)
            st.metric("En vivo", f"{hrs:02d}:{mins:02d}:{secs:02d}")

            if st.button("⏸️ Pausar"):
                st.session_state.total_elapsed += (datetime.now() - st.session_state.timer_start).total_seconds()
                st.session_state.timer_running = False
                st.rerun()
            
            if st.button("⏹️ Finalizar y Guardar", disabled=not can_register):
                t_start = st.session_state.timer_start
                t_now = datetime.now()
                total_sec = st.session_state.total_elapsed + (t_now - t_start).total_seconds()
                total_min = int(total_sec // 60) + (1 if total_sec % 60 > 0 else 0)
                
                tz_local = timezone(timedelta(hours=-5))
                start_dt = datetime.combine(fecha_sel, t_start.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                end_dt = datetime.combine(fecha_sel, t_now.time()).replace(tzinfo=tz_local).astimezone(timezone.utc)
                
                if total_min == 1 and start_dt.strftime('%H:%M') == end_dt.strftime('%H:%M'):
                    end_dt = start_dt + timedelta(minutes=1)
                
                supabase.table("time_entries").insert({
                    "profile_id": target_user_id,
                    "project_id": p_id,
                    "description": descripcion,
                    "start_time": start_dt.isoformat(),
                    "end_time": end_dt.isoformat(),
                    "total_minutes": total_min,
                    "is_billable": es_facturable
                }).execute()
                
                st.session_state.timer_running = False
                st.session_state.total_elapsed = 0
                st.session_state.timer_start = None
                st.session_state.success_msg = "✅ Registro con cronómetro guardado exitosamente."
                st.session_state.form_key_suffix += 1
                st.rerun()
        else:
            if st.session_state.total_elapsed > 0:
                hrs, rem = divmod(int(st.session_state.total_elapsed), 3600)
                mins, secs = divmod(rem, 60)
                st.metric("Pausado", f"{hrs:02d}:{mins:02d}:{secs:02d}")
                if st.button("▶️ Continuar"):
                    st.session_state.timer_start = datetime.now()
                    st.session_state.timer_running = True
                    st.rerun()
            else:
                if st.button("▶️ Iniciar", disabled=not can_register):
                    st.session_state.timer_start = datetime.now()
                    st.session_state.timer_running = True
                    st.rerun()

    # 4. TABLA DE HISTORIAL (Diferenciada por rol)
    st.markdown("---")
    st.subheader("📋 Historial de Horas")
    
    # Query base
    query = supabase.table("time_entries").select("*, profiles(full_name, role_id, roles(name)), projects(name, currency, clients(name))").order("created_at", desc=True)
    if not st.session_state.is_admin:
        query = query.eq("profile_id", st.session_state.user.id)
    
    entries_resp = query.execute()
    
    if entries_resp.data:
        df = pd.json_normalize(entries_resp.data)
        
        # SANEAMIENTO HORARIO GLOBAL (Garantizar UTC-5 Lima/Bogotá)
        def to_local_manual(s):
            if pd.isna(s) or s == 'nan' or not s: return None
            try:
                # Parsear como UTC, convertir a Lima, y quitar info de TZ para Streamlit
                return pd.to_datetime(s, utc=True).tz_convert('America/Lima').tz_localize(None)
            except:
                try:
                    # Fallback manual si tz_convert falla (Shift -5h)
                    return pd.to_datetime(s).replace(tzinfo=None) - pd.Timedelta(hours=5)
                except:
                    return None

        df['dt_ref'] = df['start_time'].fillna(df['created_at'])
        df['dt_start_naive'] = df['dt_ref'].apply(to_local_manual)
        df['dt_end_naive'] = df['end_time'].apply(to_local_manual)
        
        df['Inicio'] = df['dt_start_naive'].dt.strftime('%H:%M').fillna('---')
        df['Fin'] = df['dt_end_naive'].dt.strftime('%H:%M').fillna('---')
        df['Fecha'] = df['dt_start_naive'].dt.strftime('%d.%m-%Y').fillna('---')
        df['Tiempo (hh:mm)'] = df['total_minutes'].apply(lambda x: f"{int(x)//60:02d}:{int(x)%60:02d}")
        
        # Mapeo de nombres seguro
        df['Cliente'] = df['projects.clients.name'].fillna('Desconocido')
        df['Proyecto'] = df['projects.name'].fillna('Desconocido')
        df['Usuario_Nombre'] = df['profiles.full_name'].fillna('...')
        
        if st.session_state.is_admin:
            rates_resp = supabase.table("project_rates").select("*").execute()
            rates_df = pd.DataFrame(rates_resp.data)
            
            def calc_billing(row):
                rate = 0.0
                if not rates_df.empty:
                    r = rates_df[(rates_df['project_id'] == row['project_id']) & (rates_df['role_id'] == row['profiles.role_id'])]
                    rate = float(r['rate'].iloc[0]) if not r.empty else 0.0
                c_total = (row['total_minutes'] / 60) * rate
                return pd.Series([rate, c_total, c_total if row['is_billable'] else 0.0])

            df[['Costo Hora', 'Costo Total', 'Costo Facturable']] = df.apply(calc_billing, axis=1)
            # Redondeo físico
            df['Costo Hora'] = df['Costo Hora'].fillna(0).round(2)
            df['Costo Total'] = df['Costo Total'].fillna(0).round(2)
            df['Costo Facturable'] = df['Costo Facturable'].fillna(0).round(2)
            
            display_cols = ['Fecha', 'Usuario_Nombre', 'Cliente', 'Proyecto', 'description', 'Inicio', 'Fin', 'Tiempo (hh:mm)', 'Costo Hora', 'is_billable', 'Costo Total', 'Costo Facturable', 'is_paid', 'invoice_number']
            col_names = ['Fecha', 'Usuario', 'Cliente', 'Proyecto', 'Detalle', 'Hora Inicio', 'Hora Final', 'Tiempo', 'Costo Hora', 'Facturable', 'Costo Total', 'Costo Facturable', '¿Cobrado?', 'Factura #']
            
            # Configuración uniforme
            col_cfg_hist = {
                "Costo Hora": st.column_config.NumberColumn(format="%,.2f"),
                "Costo Total": st.column_config.NumberColumn(format="%,.2f"),
                "Costo Facturable": st.column_config.NumberColumn(format="%,.2f"),
                "Facturable": st.column_config.CheckboxColumn(label="✅")
            }

            edited_df = st.data_editor(
                df[display_cols].rename(columns=dict(zip(display_cols, col_names))),
                column_config=col_cfg_hist,
                use_container_width=True, hide_index=True,
                disabled=[c for c in col_names if c not in ['¿Cobrado?', 'Factura #', 'Facturable']]
            )
            
            if st.button("Guardar Cambios Administrativos"):
                for i, row in edited_df.iterrows():
                    orig = df.iloc[i]
                    if row['¿Cobrado?'] != orig['is_paid'] or row['Factura #'] != orig['invoice_number'] or row['Facturable'] != orig['is_billable']:
                        supabase.table("time_entries").update({
                            "is_paid": row['¿Cobrado?'], "invoice_number": row['Factura #'], "is_billable": row['Facturable']
                        }).eq("id", orig['id']).execute()
                st.success("✅ Cambios guardados.")
                st.rerun()
        else:
            # Vista simplificada para Usuario
            display_cols = ['Fecha', 'Cliente', 'Proyecto', 'description', 'Inicio', 'Fin', 'Tiempo (hh:mm)']
            col_names = ['Fecha', 'Cliente', 'Proyecto', 'Detalle', 'Hora Inicio', 'Hora Final', 'Tiempo']
            
            st.dataframe(
                df[display_cols].rename(columns=dict(zip(display_cols, col_names))),
                use_container_width=True, hide_index=True
            )

# --- CUERPO PRINCIPAL ---
if not st.session_state.user:
    st.subheader("Acceso al Sistema")
    with st.form("login_form"):
        email = st.text_input("Correo electrónico")
        password = st.text_input("Contraseña", type="password")
        if st.form_submit_button("Entrar"):
            login_user(email, password)
else:
    with st.sidebar:
        st.write(f"👤 **{st.session_state.profile['full_name']}**")
        st.write(f"🏷️ Rol: {st.session_state.profile['roles']['name']}")
        st.write(f"🔑 Tipo: {'Administrador' if st.session_state.is_admin else 'Usuario'}")
        if st.button("Cerrar Sesión"):
            st.session_state.user = None
            st.rerun()

    if st.session_state.is_admin:
        menu = ["Panel General", "Registro de Tiempos", "Clientes", "Proyectos", "Usuarios", "Roles y Tarifas", "Facturación y Reportes"]
        choice = st.sidebar.selectbox("Seleccione Módulo", menu)

        if choice == "Panel General":
            st.header("📊 Panel General de Horas")
            
            # Query base (Admin ve todo)
            entries_q = supabase.table("time_entries").select("*, profiles(full_name, role_id, roles(name)), projects(name, currency, clients(name))").order("created_at", desc=True)
            entries = entries_q.execute()
            rates = supabase.table("project_rates").select("*").execute()
            
            if entries.data:
                df = pd.json_normalize(entries.data)
                rates_df = pd.DataFrame(rates.data)
                
                # Conversión horaria manual garantizada (UTC-5)
                df['dt_ref'] = df['start_time'].fillna(df['created_at'])
                df['dt_start'] = df['dt_ref'].apply(lambda x: pd.to_datetime(x, utc=True).tz_convert('America/Lima').tz_localize(None) if pd.notna(x) and x != 'nan' else None)
                df['dt_end'] = df['end_time'].apply(lambda x: pd.to_datetime(x, utc=True).tz_convert('America/Lima').tz_localize(None) if pd.notna(x) and x != 'nan' else None)
                
                df['Hora Inicio'] = df['dt_start'].dt.strftime('%H:%M').fillna('---')
                df['Hora Final'] = df['dt_end'].dt.strftime('%H:%M').fillna('---')
                df['Tiempo (hh:mm)'] = df['total_minutes'].apply(lambda x: f"{int(x)//60:02d}:{int(x)%60:02d}")
                df['Fecha'] = df['dt_start'].dt.strftime('%d.%m-%Y').fillna('---')
                
                # Respaldo si falló el apply (si resultaron nulos pero no deberían)
                if not df.empty and df['Hora Inicio'].iloc[0] == '---' and not df['dt_ref'].isnull().all():
                     df['dt_start'] = (pd.to_datetime(df['dt_ref'], utc=True) - pd.Timedelta(hours=5)).dt.tz_localize(None)
                     df['dt_end'] = (pd.to_datetime(df['end_time'], utc=True) - pd.Timedelta(hours=5)).dt.tz_localize(None)
                     df['Hora Inicio'] = df['dt_start'].dt.strftime('%H:%M').fillna('---')
                     df['Hora Final'] = df['dt_end'].dt.strftime('%H:%M').fillna('---')
                     df['Fecha'] = df['dt_start'].dt.strftime('%d.%m-%Y').fillna('---')
                
                def get_cost(row):
                    if not rates_df.empty:
                        r = rates_df[(rates_df['project_id'] == row['project_id']) & (rates_df['role_id'] == row['profiles.role_id'])]
                        return float(r['rate'].iloc[0]) if not r.empty else 0.0
                    return 0.0

                df['Costo Hora'] = df.apply(get_cost, axis=1)
                df['Valor Total'] = (df['total_minutes'] / 60) * df['Costo Hora']
                df['Costo Facturable'] = df.apply(lambda r: r['Valor Total'] if r['is_billable'] else 0.0, axis=1)
                
                # Renombrar para visualización
                df = df.rename(columns={
                    'profiles.full_name': 'Usuario',
                    'profiles.roles.name': 'Rol',
                    'projects.clients.name': 'Cliente',
                    'projects.name': 'Proyecto',
                    'is_billable': 'Facturable'
                })
                
                # Filtros
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    f_user = st.multiselect("Filtrar por Usuario", df['Usuario'].unique())
                with col_f2:
                    f_client = st.multiselect("Filtrar por Cliente", df['Cliente'].unique())
                
                filtered_df = df.copy()
                if f_user: filtered_df = filtered_df[filtered_df['Usuario'].isin(f_user)]
                if f_client: filtered_df = filtered_df[filtered_df['Cliente'].isin(f_client)]
                
                # Columnas finales (Admin ve todo y puede editar)
                display_cols = ['id', 'Fecha', 'Usuario', 'Rol', 'Cliente', 'Proyecto', 'Hora Inicio', 'Hora Final', 'Tiempo (hh:mm)', 'Costo Hora', 'Valor Total', 'Costo Facturable', 'Facturable']
                
                # FORZAR REDONDEO FÍSICO EN EL DF PARA EVITAR DECIMALES LARGOS
                filtered_df['Costo Hora'] = filtered_df['Costo Hora'].fillna(0).round(2)
                filtered_df['Valor Total'] = filtered_df['Valor Total'].fillna(0).round(2)
                filtered_df['Costo Facturable'] = filtered_df['Costo Facturable'].fillna(0).round(2)

                # Configuración de columnas para alineación y formato
                col_config = {
                    "id": None, # Habilitar ocultamiento real sin error
                    "Costo Hora": st.column_config.NumberColumn(format="%,.2f"),
                    "Valor Total": st.column_config.NumberColumn(format="%,.2f"),
                    "Costo Facturable": st.column_config.NumberColumn(format="%,.2f"),
                    "Facturable": st.column_config.CheckboxColumn(label="✅")
                }
                
                edited_gen = st.data_editor(
                    filtered_df[display_cols], 
                    column_config=col_config,
                    use_container_width=True, hide_index=True,
                    disabled=['Rol', 'Cliente', 'Proyecto', 'Tiempo (hh:mm)', 'Costo Hora', 'Valor Total', 'Costo Facturable'] # Solo lo básico y Facturable es editable
                )
                
                # El desmarcado de "Facturable" se refleja en el editor. Recalcular métricas dinámicas para visualización rápida:
                billable_total_live = edited_gen[edited_gen['Facturable'] == True]['Costo Facturable'].sum()
                st.info(f"💰 **Total Facturable Proyectado (en esta vista): {billable_total_live:,.2f}**")
                
                col_btn1, col_btn2 = st.columns([1, 1])
                with col_btn1:
                    if st.button("Guardar cambios en Panel General"):
                        for i, row in edited_gen.iterrows():
                            # Encontrar la fila original por ID
                            orig_id = row['id']
                            orig_row = df[df['id'] == orig_id].iloc[0]
                            
                            updates = {}
                            if row['Facturable'] != orig_row['Facturable']: updates["is_billable"] = row['Facturable']
                            if row['Fecha'] != orig_row['Fecha']: 
                                try:
                                    # Intentar parsear fecha editada
                                    new_d = datetime.strptime(row['Fecha'], '%d.%m-%Y')
                                    # Mantener la hora original
                                    old_dt = pd.to_datetime(orig_row['dt_ref'])
                                    new_dt = old_dt.replace(year=new_d.year, month=new_d.month, day=new_d.day)
                                    updates["start_time"] = new_dt.isoformat()
                                except: pass
                            
                            if updates:
                                supabase.table("time_entries").update(updates).eq("id", orig_id).execute()
                        st.success("✅ Cambios administrativos guardados.")
                        st.rerun()

                with col_btn2:
                    # Exportar a Excel con manejo de errores si la biblioteca no ha cargado
                    try:
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            filtered_df[display_cols].to_excel(writer, index=False, sheet_name='Historial')
                        st.download_button(
                            label="Excel 📥",
                            data=output.getvalue(),
                            file_name=f"historial_horas_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except ModuleNotFoundError:
                        st.warning("⚠️ El módulo de descarga Excel se está inicializando. Por favor, intente de nuevo en unos minutos.")
                    except Exception as e:
                        st.error(f"Error al generar Excel: {e}")
                
                # Calcular inversión por moneda
                st.subheader("Inversión Total por Divisa")
                if not filtered_df.empty:
                    # Agrupar por la moneda del proyecto (que sacamos del join)
                    # El campo en el df normalizado es 'projects.currency'
                    if 'projects.currency' in filtered_df:
                        metrics_cols = st.columns(len(filtered_df['projects.currency'].unique()))
                        for i, (curr, group) in enumerate(filtered_df.groupby('projects.currency')):
                            with metrics_cols[i]:
                                total_curr = group['Valor Total'].sum()
                                st.metric(f"Total {curr}", f"{curr} {total_curr:,.2f}")
                    else:
                        st.metric("Inversión Total", f"${filtered_df['Valor Total'].sum():,.2f}")
            else:
                st.info("No hay registros de tiempo aún.")

        elif choice == "Registro de Tiempos":
            mostrar_registro_tiempos()

        elif choice == "Clientes":
            st.header("🏢 Gestión de Clientes")
            with st.expander("➕ Crear Nuevo Cliente", expanded=True):
                with st.form("form_cliente"):
                    nombre = st.text_input("Nombre o Razón Social")
                    doi_type = st.selectbox("Tipo DOI", ["RUC", "DNI", "CE", "PASAPORTE", "OTROS"])
                    doi_num = st.text_input("Número de Documento")
                    email_cli = st.text_input("Email de contacto")
                    celular_cli = st.text_input("Número de Contacto")
                    direccion = st.text_area("Dirección")
                    
                    if st.form_submit_button("Guardar Cliente"):
                        existente = supabase.table("clients").select("*").or_(f"name.eq.{nombre},doi_number.eq.{doi_num}").execute()
                        if existente.data:
                            st.error("❌ Error: Ya existe un cliente con ese nombre o número de documento.")
                        else:
                            supabase.table("clients").insert({
                                "name": nombre, "doi_type": doi_type, "doi_number": doi_num, 
                                "address": direccion, "email": email_cli, "contact_number": celular_cli
                            }).execute()
                            st.success(f"✅ Cliente '{nombre}' creado con éxito.")
            
            st.subheader("Clientes Registrados")
            clientes_q = supabase.table("clients").select("*").order("name").execute()
            if clientes_q.data:
                c_df = pd.DataFrame(clientes_q.data)
                edited_clients = st.data_editor(
                    c_df[['id', 'name', 'doi_type', 'doi_number', 'email', 'contact_number', 'address']],
                    use_container_width=True, hide_index=True,
                    disabled=["id"]
                )
                if st.button("Guardar Cambios de Clientes"):
                    for i, row in edited_clients.iterrows():
                        orig = clientes_q.data[i]
                        # Solo actualizar si hubo cambios
                        diff = {k: v for k, v in row.items() if v != orig.get(k)}
                        if diff:
                            supabase.table("clients").update(diff).eq("id", row['id']).execute()
                    st.success("✅ Datos de clientes actualizados.")
                    st.rerun()

        elif choice == "Proyectos":
            st.header("📁 Gestión de Proyectos")
            clientes = supabase.table("clients").select("id, name").order("name").execute()
            if not clientes.data:
                st.warning("Debe crear un cliente primero.")
            else:
                client_map = {c['name']: c['id'] for c in clientes.data}
                if 'proj_key_suffix' not in st.session_state: st.session_state.proj_key_suffix = 0
                if 'proj_success_msg' in st.session_state:
                    st.success(st.session_state.proj_success_msg)
                    del st.session_state.proj_success_msg
                    
                with st.expander("➕ Crear Nuevo Proyecto"):
                    with st.form("form_proyecto"):
                        cliente_create = st.selectbox("Seleccionar Cliente", list(client_map.keys()), key=f"p_c_create_{st.session_state.proj_key_suffix}")
                        proj_name = st.text_input("Nombre del Proyecto", key=f"p_name_{st.session_state.proj_key_suffix}")
                        moneda = st.selectbox("Moneda del Proyecto", ["PEN", "USD"], key=f"p_curr_{st.session_state.proj_key_suffix}")
                        if st.form_submit_button("Crear Proyecto"):
                            existente = supabase.table("projects").select("*").eq("client_id", client_map[cliente_create]).eq("name", proj_name).execute()
                            if existente.data:
                                st.error(f"❌ El cliente '{cliente_create}' ya tiene un proyecto llamado '{proj_name}'.")
                            else:
                                supabase.table("projects").insert({
                                    "client_id": client_map[cliente_create], "name": proj_name, "currency": moneda
                                }).execute()
                                st.session_state.proj_success_msg = f"✅ Proyecto '{proj_name}' creado con éxito."
                                st.session_state.proj_key_suffix += 1
                                st.rerun()

                st.subheader("Proyectos Existentes")
                proyectos = supabase.table("projects").select("*, clients(name)").order("name").execute()
                if proyectos.data:
                    p_df = pd.json_normalize(proyectos.data)
                    p_df = p_df[['clients.name', 'name', 'currency']]
                    p_df.columns = ['Cliente', 'Proyecto', 'Moneda']
                    st.table(p_df)
                    
                    st.markdown("---")
                    st.subheader("✏️ Editar Moneda de Proyecto")
                    proj_list = {f"{p['clients']['name']} - {p['name']}": p['id'] for p in proyectos.data}
                    p_to_edit = st.selectbox("Seleccionar Proyecto para Editar", list(proj_list.keys()))
                    
                    # Buscar moneda actual
                    curr_p = [p for p in proyectos.data if p['id'] == proj_list[p_to_edit]][0]
                    new_curr = st.selectbox("Nueva Moneda", ["PEN", "USD"], index=0 if curr_p['currency'] == 'PEN' else 1)
                    
                    if st.button("Actualizar Moneda"):
                        supabase.table("projects").update({"currency": new_curr}).eq("id", proj_list[p_to_edit]).execute()
                        st.success(f"✅ Moneda de '{p_to_edit}' actualizada a {new_curr}.")
                        st.rerun()
                else:
                    st.info("No hay proyectos registrados.")

        elif choice == "Usuarios":
            st.header("👥 Gestión de Usuarios")
            roles = supabase.table("roles").select("id, name").order("name").execute()
            role_map = {r['name']: r['id'] for r in roles.data}
            
            with st.form("form_usuario"):
                u_email = st.text_input("Email (será su acceso)")
                u_pass = st.text_input("Contraseña", type="password")
                u_name = st.text_input("Nombre Completo")
                u_username = st.text_input("Nombre de Usuario (interno)")
                u_doi_type = st.selectbox("Tipo DOI", ["DNI", "RUC", "CE", "PASAPORTE"])
                u_doi_number = st.text_input("Número de DOI")
                u_role = st.selectbox("Rol Operativo (para tarifas)", list(role_map.keys()))
                u_is_admin = st.checkbox("¿Es Administrador?")
                st.info("💡 Por seguridad, los nuevos usuarios se crean DESACTIVADOS.")
                
                if st.form_submit_button("Crear Usuario"):
                    if not u_email or not u_pass:
                        st.error("❌ Email y contraseña son obligatorios.")
                    else:
                        try:
                            new_u = supabase.auth.admin.create_user({
                                "email": u_email.strip(), 
                                "password": u_pass, 
                                "email_confirm": True
                            })
                            supabase.table("profiles").insert({
                                "id": new_u.user.id, 
                                "username": u_username, 
                                "full_name": u_name, 
                                "role_id": role_map[u_role],
                                "doi_type": u_doi_type,
                                "doi_number": u_doi_number,
                                "is_active": False,
                                "is_admin": u_is_admin
                            }).execute()
                            st.success(f"✅ Usuario '{u_name}' creado con éxito.")
                            st.info("⚠️ Recuerde activarlo en la tabla de abajo para que pueda iniciar sesión.")
                        except Exception as e:
                            st.error(f"❌ Error de permisos: {e}")
                            st.warning("Asegúrese de que el 'SUPABASE_SERVICE_KEY' esté bien configurado en los Secretos de Streamlit.")
            
            st.subheader("Usuarios Registrados")
            users_resp = supabase.table("profiles").select("*, roles(name)").execute()
            if users_resp.data:
                u_df = pd.DataFrame([
                    {
                        "ID": u['id'],
                        "Nombre": u['full_name'], 
                        "Interno": u['username'],
                        "DOI": f"{u['doi_type']} {u['doi_number']}",
                        "Rol": u['roles']['name'],
                        "Activo": u['is_active'],
                        "Admin": u['is_admin']
                    }
                    for u in users_resp.data
                ])
                
                # Editor para activar/desactivar y dar admin
                st.write("Edite directamente los privilegios en la tabla y guarde:")
                edited_users = st.data_editor(u_df, use_container_width=True, hide_index=True, disabled=["ID", "Interno", "DOI"])
                
                if st.button("Guardar Cambios de Usuarios"):
                    changed_count = 0
                    for i, row in edited_users.iterrows():
                        orig = users_resp.data[i]
                        if row['Activo'] != orig['is_active'] or row['Admin'] != orig['is_admin']:
                            supabase.table("profiles").update({
                                "is_active": row['Activo'], "is_admin": row['Admin']
                            }).eq("id", row['ID']).execute()
                            changed_count += 1
                    if changed_count > 0:
                        st.session_state.user_success_msg = f"✅ {changed_count} usuarios actualizados."
                        st.rerun()

            # Gestión de Mensajes persistentes para usuarios
            if 'user_success_msg' in st.session_state:
                st.success(st.session_state.user_success_msg)
                del st.session_state.user_success_msg

        elif choice == "Roles y Tarifas":
            st.header("💰 Roles y Tarifas por Proyecto")
            proyectos = supabase.table("projects").select("id, name, clients(name)").execute()
            if not proyectos.data:
                st.warning("Debe crear proyectos primero.")
            else:
                proj_map = {f"{p['clients']['name']} - {p['name']}": p['id'] for p in proyectos.data}
                proj_sel = st.selectbox("Seleccionar Proyecto", list(proj_map.keys()))
                
                # Suffix para resetear inputs de tarifas al cambiar proyecto
                if 'rate_key_prefix' not in st.session_state: st.session_state.rate_key_prefix = 0
                if 'last_proj_sel' not in st.session_state: st.session_state.last_proj_sel = proj_sel
                
                if st.session_state.last_proj_sel != proj_sel:
                    st.session_state.rate_key_prefix += 1
                    st.session_state.last_proj_sel = proj_sel
                    st.rerun()
                
                roles = supabase.table("roles").select("*").execute()
                
                st.subheader(f"Tarifas para: {proj_sel}")
                tarifas_nuevas = {}
                for role in roles.data:
                    col_r1, col_r2 = st.columns([2, 1])
                    with col_r1:
                        st.write(f"Rol: **{role['name']}**")
                    with col_r2:
                        # Buscar tarifa actual - Usamos cache local por proyecto para evitar queries repetitivas si es posible, o simplemente select simple
                        current_rate = supabase.table("project_rates").select("rate").eq("project_id", proj_map[proj_sel]).eq("role_id", role['id']).execute()
                        initial_val = float(current_rate.data[0]['rate']) if current_rate.data else 0.0
                        new_rate = st.number_input(f"Tarifa ({role['name']})", value=initial_val, key=f"rate_{proj_map[proj_sel]}_{role['id']}")
                        tarifas_nuevas[role['id']] = (new_rate, initial_val, bool(current_rate.data))
                
                if st.button("Guardar Todas las Tarifas"):
                    for r_id, (val, old_val, exists) in tarifas_nuevas.items():
                        if val != old_val:
                            if exists:
                                supabase.table("project_rates").update({"rate": val}).eq("project_id", proj_map[proj_sel]).eq("role_id", r_id).execute()
                            else:
                                supabase.table("project_rates").insert({"project_id": proj_map[proj_sel], "role_id": r_id, "rate": val}).execute()
                    st.success(f"✅ Tarifas para '{proj_sel}' guardadas.")
                    st.rerun()

        elif choice == "Facturación y Reportes":
            st.header("📄 Facturación y Reportes")
            
            # Filtros de Reporte
            clientes_q = supabase.table("clients").select("id, name, doi_type, doi_number, address").order("name").execute()
            if not clientes_q.data:
                st.warning("Debe registrar clientes primero.")
            else:
                with st.sidebar:
                    st.markdown("### Configuración de Reporte")
                    cli_map = {c['name']: c for c in clientes_q.data}
                    cli_name_sel = st.selectbox("Seleccionar Cliente", list(cli_map.keys()))
                    cli_data = cli_map[cli_name_sel]
                    
                    date_range = st.date_input("Rango de Fechas", [datetime.today().replace(day=1), datetime.today()])
                    
                    st.markdown("---")
                    st.markdown("### Datos para la Carta")
                    tenor = st.text_area("Tenor de la Carta", value="Por la presente detallamos los servicios profesionales realizados en el periodo indicado:", height=100)
                    cuentas = st.text_area("Cuentas Bancarias", value="BCP Soles: XXX-XXXXXXX-X-XX\nBCP Dólares: YYY-YYYYYYY-Y-YY", height=80)
                    firma = st.text_input("Responsable Firma", value=st.session_state.profile['full_name'])
                
                if len(date_range) == 2:
                    start_d, end_d = date_range
                    report_q = supabase.table("time_entries").select("*, profiles(full_name, role_id, roles(name)), projects(name, currency, client_id)").eq("projects.client_id", cli_data['id']).execute()
                    
                    if report_q.data:
                        df_rep = pd.json_normalize(report_q.data)
                        if df_rep.empty:
                            st.info("No hay registros en el periodo seleccionado.")
                        else:
                            # Shift horario para reportes
                            df_rep['dt_ref'] = df_rep['start_time'].fillna(df_rep['created_at'])
                            df_rep['dt_start'] = (pd.to_datetime(df_rep['dt_ref'], utc=True, errors='coerce') - pd.Timedelta(hours=5)).dt.tz_localize(None)
                            df_rep['Fecha_dt'] = df_rep['dt_start'].dt.date
                            df_rep['Fecha_str'] = df_rep['dt_start'].dt.strftime('%d.%m-%Y')
                            
                            # Filtro solo para CARTA (Tab 1), el Dashboard verá todo
                            # Saneamiento de monedas (eliminar nan)
                            monedas_disp = [m for m in df_rep['projects.currency'].unique() if pd.notna(m) and str(m) != 'nan']
                            if not monedas_disp:
                                st.warning("No hay monedas válidas en los proyectos seleccionados.")
                                moneda_liq = None
                            else:
                                moneda_liq = st.sidebar.selectbox("Moneda a Liquidar (Carta)", monedas_disp)
                            
                            rates = supabase.table("project_rates").select("*").execute()
                            rates_df = pd.DataFrame(rates.data)
                            
                            def get_cost_rep(row):
                                if rates_df.empty: return 0.0
                                r = rates_df[(rates_df['project_id'] == row['project_id']) & (rates_df['role_id'] == row['profiles.role_id'])]
                                return float(r['rate'].iloc[0]) if not r.empty else 0.0
                            
                            df_rep['Horas_num'] = df_rep['total_minutes'] / 60
                            df_rep['Costo_H'] = df_rep.apply(get_cost_rep, axis=1)
                            df_rep['Total_Monto'] = df_rep['Horas_num'] * df_rep['Costo_H']
                            # Saneamiento de Fecha_str para evitar NaN visuales
                            df_rep['Fecha_str'] = df_rep['dt_start'].dt.strftime('%d.%m-%Y').fillna('---')
                            
                            tab1, tab2 = st.tabs(["📝 Carta de Liquidación", "📋 Anexo Detallado / Dashboard"])
                            
                            with tab1:
                                if moneda_liq:
                                    # Generar contenido de carta FORMAL (Página 1)
                                    df_carta = df_rep[df_rep['projects.currency'] == moneda_liq].copy()
                                    total_general_liq = df_carta['Total_Monto'].sum()
                                    
                                    # Limpieza de DOI y Dirección para evitar "nan"
                                    doi_str = str(cli_data.get('doi_number', '')).strip()
                                    if doi_str == 'nan': doi_str = '---'
                                    addr_str = str(cli_data.get('address', '')).strip()
                                    if addr_str == 'nan' or not addr_str: addr_str = 'Lima, Perú.'

                                    # Limpieza de DOI y Dirección para evitar "nan"
                                    doi_str = str(cli_data.get('doi_number', '')).strip()
                                    if doi_str == 'nan' or not doi_str: doi_str = '---'
                                    addr_str = str(cli_data.get('address', '')).strip()
                                    if addr_str == 'nan' or not addr_str: addr_str = 'Lima, Perú.'

                                    carta_formal_html = f"""
<div style="padding: 40px; background-color: white; color: black !important; font-family: 'Times New Roman', serif; line-height: 1.5; border: 1px solid #ddd;">
    <div style="text-align: center; margin-bottom: 30px; border-bottom: 2px solid #7c3aed; padding-bottom: 10px;">
        <h2 style="margin: 0; color: #333;">REPORTE DE SERVICIOS PROFESIONALES</h2>
    </div>
    
    <p style="text-align: right;">Lima, {datetime.today().strftime('%d de %B de %Y')}</p>
    
    <div style="margin-bottom: 30px;">
        <p><strong>Señores:</strong><br>
        {cli_name_sel.upper()}<br>
        RUC: {doi_str}<br>
        {addr_str}</p>
    </div>
    
    <div style="margin-bottom: 20px;">
        <p><strong>Ref: Liquidación de Honorarios</strong><br>
        Periodo: {start_d.strftime('%d.%m-%Y')} al {end_d.strftime('%d.%m-%Y')}</p>
    </div>

    <p style="text-align: justify;">{tenor}</p>
    
    <div style="background-color: #f9f9f9; padding: 20px; text-align: center; margin: 30px 0; border: 1px solid #ccc;">
        <p style="margin: 0; font-size: 0.9em;">MONTO TOTAL A LIQUIDAR</p>
        <h3 style="margin: 5px 0; color: #000;">{moneda_liq} {total_general_liq:,.2f}</h3>
    </div>
    
    <div style="margin-bottom: 40px;">
        <p><strong>Instrucciones de Pago:</strong></p>
        <div style="font-family: monospace; white-space: pre-wrap; background: #fafafa; padding: 10px; border-left: 3px solid #7c3aed;">{cuentas}</div>
    </div>
    
    <div style="margin-top: 50px; width: 250px; border-top: 1px solid #000; text-align: center; padding-top: 5px;">
        <strong>{firma}</strong><br>Responsable
    </div>
</div>
"""
                                    st.markdown(carta_formal_html, unsafe_allow_html=True)
                                    st.info("💡 El detalle de horas aparece en la pestaña 'Anexo Detallado'.")
                                else:
                                    st.warning("Seleccione una moneda para generar el reporte.")

                            with tab2:
                                if moneda_liq:
                                    st.subheader(f"Anexo: Detalle Técnico ({moneda_liq})")
                                    df_anexo = df_rep[df_rep['projects.currency'] == moneda_liq].copy()
                                    
                                    # Tabla de anexo interactiva con formato
                                    anexo_display = df_anexo[['Fecha_str', 'profiles.full_name', 'projects.name', 'description', 'total_minutes', 'Total_Monto']].copy()
                                    anexo_display['Tiempo'] = anexo_display['total_minutes'].apply(lambda x: f"{int(x)//60:02d}:{int(x)%60:02d}")
                                    anexo_display = anexo_display[['Fecha_str', 'profiles.full_name', 'projects.name', 'description', 'Tiempo', 'Total_Monto']]
                                    anexo_display.columns = ['Fecha', 'Consultor', 'Proyecto', 'Actividad', 'Tiempo', 'Valor']
                                    
                                    st.dataframe(
                                        anexo_display,
                                        column_config={"Valor": st.column_config.NumberColumn(format=f"{moneda_liq} %,.2f")},
                                        use_container_width=True, hide_index=True
                                    )
                                    
                                    # Botón Excel para este reporte con manejo de errores
                                    try:
                                        exc_io = io.BytesIO()
                                        with pd.ExcelWriter(exc_io, engine='openpyxl') as writer:
                                            anexo_display.to_excel(writer, index=False, sheet_name='Detalle')
                                        st.download_button(
                                            label="Descargar Anexo Excel 📥",
                                            data=exc_io.getvalue(),
                                            file_name=f"Anexo_{cli_name_sel}_{moneda_liq}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                                    except ModuleNotFoundError:
                                        st.warning("⚠️ Módulo Excel en preparación. Intente en unos instantes.")
                                    except Exception as e:
                                        st.error(f"Error al generar Anexo Excel: {e}")

                                st.markdown("---")
                                st.subheader("Dashboard Consolidado (Todas las monedas)")
                                dash_data = df_rep.groupby(['profiles.full_name', 'projects.currency']).agg({
                                    'Horas_num': 'sum',
                                    'Total_Monto': 'sum'
                                }).reset_index()
                                # Saneamiento de moneda en dashboard
                                dash_data['projects.currency'] = dash_data['projects.currency'].fillna('---')
                                dash_data['Tiempo'] = dash_data['Horas_num'].apply(lambda h: f"{int(round(h*60))//60:02d}:{int(round(h*60))%60:02d}")
                                dash_data = dash_data.rename(columns={
                                    'profiles.full_name': 'Usuario',
                                    'projects.currency': 'Moneda',
                                    'Total_Monto': 'Inversión'
                                })
                                st.dataframe(
                                    dash_data[['Usuario', 'Moneda', 'Tiempo', 'Inversión']],
                                    column_config={"Inversión": st.column_config.NumberColumn(format="%,.2f")},
                                    use_container_width=True, hide_index=True
                                )
                    else:
                        st.info("No se encontraron registros para este cliente.")
                else:
                    st.info("Seleccione un rango de fechas en la barra lateral.")

    else:
        # Para roles de usuario no administrador
        mostrar_registro_tiempos()

# --- REFRESH DINÁMICO (Al final para no bloquear UI) ---
if st.session_state.get('timer_running'):
    time.sleep(1)
    st.rerun()
