import streamlit as st
import pandas as pd
from supabase import create_client, Client
import os
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv

# Cargar variables del archivo .env
load_dotenv()

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
    # Intentar obtener de variables de entorno (Local)
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    service_key = os.getenv("SUPABASE_SERVICE_KEY")
    
    # Si no están en el sistema (Nube), intentar con st.secrets
    if not url:
        try:
            url = st.secrets["SUPABASE_URL"]
            key = st.secrets["SUPABASE_KEY"]
            service_key = st.secrets["SUPABASE_SERVICE_KEY"]
        except:
            st.error("❌ Error: No se configuraron las llaves de Supabase. Revisa la Guía de Despliegue.")
            st.stop()
            
    # Para el administrador usamos service_key para gestionar usuarios, si no, anon key
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
        st.warning(f"⚠️ **Atención**: No se han definido tarifas para el rol en este proyecto. Seleccione otro proyecto o pida al administrador que configure las tarifas.")
        can_register = False
    else:
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
                
                # Usar timezone aware datetimes para evitar "None" en base de datos al parsear
                t1 = datetime.combine(fecha_sel, t1_dt.time()).replace(tzinfo=timezone.utc)
                t2 = datetime.combine(fecha_sel, t2_dt.time()).replace(tzinfo=timezone.utc)
                
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
                
                # Combinar fecha seleccionada con las horas del cronómetro, asegurar timezone UTC
                start_dt = datetime.combine(fecha_sel, t_start.time()).replace(tzinfo=timezone.utc)
                end_dt = datetime.combine(fecha_sel, t_now.time()).replace(tzinfo=timezone.utc)
                
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
        
        # Formatear hh:mm de forma robusta
        df['Inicio'] = pd.to_datetime(df['start_time'], errors='coerce', utc=True).dt.strftime('%H:%M')
        df['Fin'] = pd.to_datetime(df['end_time'], errors='coerce', utc=True).dt.strftime('%H:%M')
        df['Tiempo (hh:mm)'] = df['total_minutes'].apply(lambda x: f"{x//60:02d}:{x%60:02d}")
        df['Fecha'] = pd.to_datetime(df['created_at'], errors='coerce', utc=True).dt.date
        
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
            
            display_cols = ['Fecha', 'profiles.full_name', 'projects.clients.name', 'projects.name', 'description', 'Inicio', 'Fin', 'Tiempo (hh:mm)', 'Costo Hora', 'is_billable', 'Costo Total', 'Costo Facturable', 'is_paid', 'invoice_number']
            col_names = ['Fecha', 'Usuario', 'Cliente', 'Proyecto', 'Detalle', 'Hora Inicio', 'Hora Final', 'Tiempo', 'Costo Hora', 'Facturable', 'Costo Total', 'Costo Facturable', '¿Cobrado?', 'Factura #']
            
            edited_df = st.data_editor(
                df[display_cols].rename(columns=dict(zip(display_cols, col_names))),
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
            display_cols = ['Fecha', 'projects.clients.name', 'projects.name', 'description', 'Inicio', 'Fin', 'Tiempo (hh:mm)']
            col_names = ['Fecha', 'Cliente', 'Proyecto', 'Detalle', 'Hora Inicio', 'Hora Final', 'Tiempo']
            st.table(df[display_cols].rename(columns=dict(zip(display_cols, col_names))))

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
                
                # Formatear mm:hh y Horas de forma robusta
                df['Hora Inicio'] = pd.to_datetime(df['start_time'], errors='coerce', utc=True).dt.strftime('%H:%M')
                df['Hora Final'] = pd.to_datetime(df['end_time'], errors='coerce', utc=True).dt.strftime('%H:%M')
                df['Tiempo (hh:mm)'] = df['total_minutes'].apply(lambda x: f"{x//60:02d}:{x%60:02d}")
                df['Fecha'] = pd.to_datetime(df['created_at'], errors='coerce', utc=True).dt.date
                
                def get_cost(row):
                    if not rates_df.empty:
                        r = rates_df[(rates_df['project_id'] == row['project_id']) & (rates_df['role_id'] == row['profiles.role_id'])]
                        return float(r['rate'].iloc[0]) if not r.empty else 0.0
                    return 0.0
 
                df['Costo Hora'] = df.apply(get_cost, axis=1)
                df['Valor Total'] = (df['total_minutes'] / 60) * df['Costo Hora']
                
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
                
                # Columnas finales (Admin ve todo y puede editar 'Facturable')
                display_cols = ['Fecha', 'Usuario', 'Rol', 'Cliente', 'Proyecto', 'Hora Inicio', 'Hora Final', 'Tiempo (hh:mm)', 'Costo Hora', 'Valor Total', 'Facturable']
                
                # Redondear y formatear para tabla
                filtered_df['Costo Hora'] = filtered_df['Costo Hora'].round(2)
                filtered_df['Valor Total'] = filtered_df['Valor Total'].round(2)
                
                edited_gen = st.data_editor(
                    filtered_df[display_cols], 
                    use_container_width=True, hide_index=True,
                    disabled=[c for c in display_cols if c != 'Facturable']
                )
                
                if st.button("Guardar cambios en Panel General"):
                    for i, row in edited_gen.iterrows():
                        orig = filtered_df.iloc[i]
                        if row['Facturable'] != orig['Facturable']:
                            supabase.table("time_entries").update({"is_billable": row['Facturable']}).eq("id", orig['id']).execute()
                    st.success("✅ Privilegios de facturación actualizados.")
                    st.rerun()
                
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
                    try:
                        # Asegurar que usamos service_role key si existe
                        new_u = supabase.auth.admin.create_user({
                            "email": u_email, "password": u_pass, "email_confirm": True
                        })
                        supabase.table("profiles").insert({
                            "id": new_u.user.id, 
                            "username": u_username, 
                            "full_name": u_name, 
                            "role_id": role_map[u_role],
                            "doi_type": u_doi_type,
                            "doi_number": u_doi_number,
                            "is_active": False, # Desactivado por defecto
                            "is_admin": u_is_admin
                        }).execute()
                        st.success(f"✅ Usuario '{u_name}' creado. Debe activarlo en la tabla inferior para que pueda entrar.")
                    except Exception as e:
                        st.error(f"Error: {e}")
                        st.info("Nota: La creación de usuarios requiere que la Service Key esté configurada correctamente en los secretos.")
            
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
            clientes_q = supabase.table("clients").select("id, name").order("name").execute()
            if not clientes_q.data:
                st.warning("Debe registrar clientes primero.")
            else:
                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    cli_map = {c['name']: c['id'] for c in clientes_q.data}
                    cli_sel = st.selectbox("Seleccionar Cliente", list(cli_map.keys()))
                with col_r2:
                    date_range = st.date_input("Rango de Fechas", [datetime.today().replace(day=1), datetime.today()])
                
                if len(date_range) == 2:
                    start_d, end_d = date_range
                    # Traer registros para ese cliente y rango
                    report_q = supabase.table("time_entries").select("*, profiles(full_name, role_id, roles(name)), projects(name, currency, client_id)").eq("projects.client_id", cli_map[cli_sel]).execute()
                    
                    if report_q.data:
                        # Filtrar por fecha localmente para mayor precisión con el selector de Streamlit
                        df_rep = pd.json_normalize(report_q.data)
                        df_rep['Fecha_dt'] = pd.to_datetime(df_rep['created_at']).dt.date
                        df_rep = df_rep[(df_rep['Fecha_dt'] >= start_d) & (df_rep['Fecha_dt'] <= end_d)]
                        
                        if df_rep.empty:
                            st.info("No hay registros en el periodo seleccionado para este cliente.")
                        else:
                            # Calcular costos
                            rates = supabase.table("project_rates").select("*").execute()
                            rates_df = pd.DataFrame(rates.data)
                            
                            def get_cost_rep(row):
                                if not rates_df.empty:
                                    r = rates_df[(rates_df['project_id'] == row['project_id']) & (rates_df['role_id'] == row['profiles.role_id'])]
                                    return float(r['rate'].iloc[0]) if not r.empty else 0.0
                                return 0.0
                            
                            df_rep['Costo_H'] = df_rep.apply(get_cost_rep, axis=1)
                            df_rep['Total'] = (df_rep['total_minutes'] / 60) * df_rep['Costo_H']
                            df_rep['Horas'] = df_rep['total_minutes'] / 60
                            
                            tab1, tab2 = st.tabs(["📝 Carta al Cliente", "👤 Resumen por Usuario"])
                            
                            with tab1:
                                st.subheader(f"Carta de Cobro: {cli_sel}")
                                st.write(f"**Periodo:** {start_d} al {end_d}")
                                
                                # Consolidado por Proyecto/Moneda
                                for (proj, curr), sub in df_rep.groupby(['projects.name', 'projects.currency']):
                                    st.markdown(f"### Proyecto: {proj}")
                                    st.write(f"**Moneda:** {curr}")
                                    
                                    # Tabla resumida de tareas
                                    resumen_tareas = sub[['Fecha_dt', 'description', 'Horas', 'Total']].copy()
                                    resumen_tareas.columns = ['Fecha', 'Descripción', 'Horas', 'Monto']
                                    resumen_tareas['Monto'] = resumen_tareas['Monto'].map(lambda x: f"{curr} {x:,.2f}")
                                    resumen_tareas['Horas'] = resumen_tareas['Horas'].map(lambda x: f"{x:,.2f}")
                                    st.table(resumen_tareas)
                                    
                                    total_p = sub['Total'].sum()
                                    st.markdown(f"**TOTAL {proj}: {curr} {total_p:,.2f}**")
                                    st.markdown("---")

                            with tab2:
                                st.subheader("Resumen de Horas por Usuario")
                                # Agrupar por nombre de usuario
                                user_summary = df_rep.groupby('profiles.full_name').agg({
                                    'Horas': 'sum',
                                    'Total': 'sum'
                                }).reset_index()
                                user_summary.columns = ['Usuario', 'Total Horas', 'Costo Total']
                                user_summary['Total Horas'] = user_summary['Total Horas'].round(2)
                                user_summary['Costo Total'] = user_summary['Costo Total'].round(2)
                                st.table(user_summary)
                    else:
                        st.info("No se encontraron registros para este cliente.")
                else:
                    st.info("Seleccione un rango de fechas (Inicio y Fin).")

    else:
        # Para roles de usuario no administrador
        mostrar_registro_tiempos()
