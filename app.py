import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob
import time
import fitz  # PyMuPDF
from datetime import datetime, timezone, timedelta
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL Y ESTILOS DE LA PÁGINA
# -----------------------------------------------------------------------------
ZONA_HORARIA_QUITO = timezone(timedelta(hours=-5))

def obtener_hora_quito():
    return datetime.now(ZONA_HORARIA_QUITO)

st.set_page_config(
    page_title="Dashboard WAF & Sustentación de Tesis - UCG",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS institucional (Diseño Azul F5 / Modo Oscuro)
st.markdown("""
    <style>
    .stApp { background-color: #0F172A; color: #F8FAFC; }
    .kpi-card { border-radius: 12px 12px 0 0; padding: 16px 12px; text-align: center; color: white; box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
    .kpi-tot { background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); }
    .kpi-alto { background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); }
    .kpi-med { background: linear-gradient(135deg, #D97706 0%, #B45309 100%); }
    .kpi-bajo { background: linear-gradient(135deg, #EAB308 0%, #CA8A04 100%); }
    .kpi-seguro { background: linear-gradient(135deg, #16A34A 0%, #15803D 100%); }
    .kpi-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; opacity: 0.95; }
    .kpi-number { font-size: 30px; font-weight: 800; margin-top: 2px; }
    .section-header { background-color: #1E293B; padding: 14px 20px; border-radius: 10px; border-left: 5px solid #0284C7; margin-bottom: 15px; }
    h1, h2, h3 { color: #F8FAFC !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    [data-testid="stSidebar"] { background-color: #0B1120; }
    [data-testid="stSidebar"] * { color: #F8FAFC !important; }
    .author-credit { font-size: 11px; color: #94A3B8; font-style: italic; margin-top: -10px; margin-bottom: 15px; }
    
    /* Estilos para caja de guión de voz */
    .speech-box { background-color: #1E293B; border-left: 4px solid #38BDF8; padding: 15px; border-radius: 8px; margin-top: 15px; }
    .speech-title { color: #38BDF8; font-weight: bold; font-size: 14px; margin-bottom: 5px; }
    
    /* Estilos para botones flotantes interactivos sobre la diapositiva */
    div.element-container:has(button[key="btn_slide_prev"]),
    div.element-container:has(button[key="btn_slide_next"]) {
        margin-top: -10px;
    }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_HISTORICO = "historico_matrices_waf.xlsx"

# Widget Reloj dinámico de Quito en Sidebar
reloj_html = """
<div style="background: rgba(255, 255, 255, 0.08); border: 1px solid rgba(255, 255, 255, 0.15); border-radius: 10px; padding: 10px 12px; text-align: center; font-family: 'Segoe UI', Roboto, sans-serif;">
    <div style="font-size: 12px; font-weight: 700; color: #38BDF8; text-transform: uppercase; letter-spacing: 0.5px;">📍 QUITO, ECUADOR ⛅</div>
    <div id="reloj-quito" style="font-size: 22px; font-weight: 800; color: #F8FAFC; margin: 4px 0;">⏰ 00:00:00</div>
    <div id="fecha-quito" style="font-size: 11px; color: #94A3B8;">📅 00/00/0000 | 18°C Parcialmente Nublado</div>
</div>
<script>
function actualizarRelojQuito() {
    const ahora = new Date();
    const utc = ahora.getTime() + (ahora.getTimezoneOffset() * 60000);
    const horaQuito = new Date(utc + (3600000 * -5));
    const h = String(horaQuito.getHours()).padStart(2, '0');
    const m = String(horaQuito.getMinutes()).padStart(2, '0');
    const s = String(horaQuito.getSeconds()).padStart(2, '0');
    const dia = String(horaQuito.getDate()).padStart(2, '0');
    const mes = String(horaQuito.getMonth() + 1).padStart(2, '0');
    const anio = horaQuito.getFullYear();
    document.getElementById('reloj-quito').innerHTML = '⏰ ' + h + ':' + m + ':' + s;
    document.getElementById('fecha-quito').innerHTML = '📅 ' + dia + '/' + mes + '/' + anio + ' | 18°C Parcialmente Nublado';
}
actualizarRelojQuito();
setInterval(actualizarRelojQuito, 1000);
</script>
"""

with st.sidebar:
    components.html(reloj_html, height=115)

# USUARIOS AUTORIZADOS
USUARIOS_AUTORIZADOS = {
    "david valverde": {"clave": "David Valverde", "rol": "Administrador", "nombre": "David Ricardo Valverde B"},
    "david.valverde": {"clave": "Valverde2026!", "rol": "Administrador", "nombre": "David Ricardo Valverde B"},
    "gabriela.armendariz": {"clave": "GArmendariz2026!", "rol": "Administrador", "nombre": "Gabriela Armendariz"},
    "admin": {"clave": "admin123", "rol": "Administrador", "nombre": "Administrador Principal"}
}

# CONTROL DE SESIÓN
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<h1>🔒 Control de Acceso WAF</h1>', unsafe_allow_html=True)
        st.markdown('<p class="author-credit">Maestría en Ciberseguridad - Universidad Casa Grande</p>', unsafe_allow_html=True)
        st.caption("Ingrese sus credenciales autorizadas para acceder.")
        
        with st.form("form_login"):
            usr_input = st.text_input("👤 Usuario:")
            pwd_input = st.text_input("🔑 Contraseña:", type="password")
            if st.form_submit_button("Iniciar Sesión", use_container_width=True):
                usr_clean = usr_input.strip().lower()
                if usr_clean in USUARIOS_AUTORIZADOS and USUARIOS_AUTORIZADOS[usr_clean]["clave"] == pwd_input:
                    st.session_state['autenticado'] = True
                    st.session_state['nombre_actual'] = USUARIOS_AUTORIZADOS[usr_clean]["nombre"]
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas.")
    st.stop()

# ESTADO PARA BOTONES KPI INTERACTIVOS
if 'filtro_kpi_riesgo' not in st.session_state:
    st.session_state['filtro_kpi_riesgo'] = "Todos"

def set_filtro_kpi(categoria):
    st.session_state['filtro_kpi_riesgo'] = categoria

# LIMPIEZA Y PROCESAMIENTO DE DATOS
def procesar_df(df, fecha_carga_str=None):
    if df.empty:
        return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]

    col_pais = next((c for c in df.columns if 'PAIS' in c.upper() or 'PAÍS' in c.upper()), None)
    col_riesgo = next((c for c in df.columns if 'RIESGO' in c.upper()), None)
    col_ataque = next((c for c in df.columns if 'ATAQUE' in c.upper()), None)
    col_app = next((c for c in df.columns if 'APLICACI' in c.upper() or 'WEB' in c.upper()), None)

    rename_dict = {}
    if col_pais: rename_dict[col_pais] = 'PAÍS'
    if col_riesgo: rename_dict[col_riesgo] = 'RIESGO ABUSEIPDB (%)'
    if col_ataque: rename_dict[col_ataque] = 'TIPO DE ATAQUE'
    if col_app: rename_dict[col_app] = 'NOMBRE APLICACIÓN WEB'
    
    df = df.rename(columns=rename_dict)

    if 'PAÍS' not in df.columns: df['PAÍS'] = 'Desconocido'
    if 'TIPO DE ATAQUE' not in df.columns: df['TIPO DE ATAQUE'] = 'No especificado'
    if 'NOMBRE APLICACIÓN WEB' not in df.columns: df['NOMBRE APLICACIÓN WEB'] = 'General'

    if fecha_carga_str and 'FechaCargaMatriz' not in df.columns:
        df['FechaCargaMatriz'] = fecha_carga_str

    if 'RIESGO ABUSEIPDB (%)' in df.columns:
        df['Riesgo_Num'] = df['RIESGO ABUSEIPDB (%)'].astype(str).str.replace('%', '').str.strip()
        df['Riesgo_Num'] = pd.to_numeric(df['Riesgo_Num'], errors='coerce').fillna(0)
    else:
        df['Riesgo_Num'] = 0

    def clasificar_riesgo(val):
        if val == 0: 
            return 'Seguro (0%)'
        elif val <= 25: 
            return 'Bajo (1-25%)'
        elif val <= 60: 
            return 'Medio (26-60%)'
        else: 
            return 'Crítico (61-100%)'

    df['Nivel_Riesgo'] = df['Riesgo_Num'].apply(clasificar_riesgo)
    return df

# CARGA DE BASE HISTÓRICA CORREGIDA PERMANENTE
def cargar_historico_base():
    if os.path.exists(ARCHIVO_HISTORICO):
        try:
            df_hist = pd.read_excel(ARCHIVO_HISTORICO)
            if not df_hist.empty:
                return df_hist
        except Exception:
            pass
    
    archivos = glob.glob("resultado_ips*.xlsx")
    if archivos:
        try:
            df_init = pd.read_excel(archivos[0], sheet_name='Reporte IPs')
        except Exception:
            df_init = pd.read_excel(archivos[0], sheet_name=0)
        
        df_proc = procesar_df(df_init, fecha_carga_str=obtener_hora_quito().strftime("%Y-%m-%d"))
        df_proc.to_excel(ARCHIVO_HISTORICO, index=False, engine="openpyxl")
        return df_proc
        
    return pd.DataFrame()

df_base = cargar_historico_base()

# BARRA LATERAL
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.get('nombre_actual')}")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state['autenticado'] = False
    st.rerun()

st.sidebar.markdown("---")

tab_dash, tab_tesis, tab_admin = st.tabs([
    "📊 Dashboard WAF & IPs", 
    "🎓 Defensa de Tesis (UCG)", 
    "⚙️ Carga Diaria & Histórico"
])

# =============================================================================
# PESTAÑA 1: DASHBOARD CON BOTONES KPI INTERACTIVOS
# =============================================================================
with tab_dash:
    if df_base.empty:
        st.warning("⚠️ No hay matrices cargadas. Ve a la pestaña '⚙️ Carga Diaria & Histórico' para subir la primera matriz.")
    else:
        st.sidebar.title("📌 Consulta Histórica por Fecha")
        fechas_disponibles = sorted([str(x) for x in df_base["FechaCargaMatriz"].unique() if pd.notnull(x)], reverse=True) if "FechaCargaMatriz" in df_base.columns else []
        
        opcion_fecha = st.sidebar.selectbox("📂 Seleccionar Matriz por Fecha:", ["🌐 Ver Todo el Histórico Acumulado"] + fechas_disponibles)
        
        if opcion_fecha == "🌐 Ver Todo el Histórico Acumulado":
            df_fecha = df_base.copy()
            fecha_lbl = "Consolidado Completo"
        else:
            df_fecha = df_base[df_base["FechaCargaMatriz"].astype(str) == opcion_fecha]
            fecha_lbl = opcion_fecha

        st.sidebar.markdown("---")
        paises = ["Todos"] + sorted([str(x) for x in df_fecha['PAÍS'].dropna().unique() if str(x).strip() != ""])
        pais_sel = st.sidebar.selectbox("Filtrar por País:", paises)

        apps = ["Todas"] + sorted([str(x) for x in df_fecha['NOMBRE APLICACIÓN WEB'].dropna().unique() if str(x).strip() != ""])
        app_sel = st.sidebar.selectbox("Filtrar por Aplicación Web:", apps)

        df_pre_filtrado = df_fecha.copy()
        if pais_sel != "Todos":
            df_pre_filtrado = df_pre_filtrado[df_pre_filtrado['PAÍS'] == pais_sel]
        if app_sel != "Todas":
            df_pre_filtrado = df_pre_filtrado[df_pre_filtrado['NOMBRE APLICACIÓN WEB'] == app_sel]

        tot_ips = len(df_pre_filtrado)
        criticas = len(df_pre_filtrado[df_pre_filtrado['Riesgo_Num'] > 60])
        medios = len(df_pre_filtrado[(df_pre_filtrado['Riesgo_Num'] >= 26) & (df_pre_filtrado['Riesgo_Num'] <= 60)])
        bajos = len(df_pre_filtrado[(df_pre_filtrado['Riesgo_Num'] >= 1) & (df_pre_filtrado['Riesgo_Num'] <= 25)])
        seguros_puros = len(df_pre_filtrado[df_pre_filtrado['Riesgo_Num'] == 0])

        if st.session_state['filtro_kpi_riesgo'] == "Crítico":
            df_filtrado = df_pre_filtrado[df_pre_filtrado['Riesgo_Num'] > 60]
            txt_filtro_kpi = "🔴 Filtrado por: RIESGO CRÍTICO (61-100%)"
        elif st.session_state['filtro_kpi_riesgo'] == "Medio":
            df_filtrado = df_pre_filtrado[(df_pre_filtrado['Riesgo_Num'] >= 26) & (df_pre_filtrado['Riesgo_Num'] <= 60)]
            txt_filtro_kpi = "🟠 Filtrado por: RIESGO MEDIO (26-60%)"
        elif st.session_state['filtro_kpi_riesgo'] == "Bajo":
            df_filtrado = df_pre_filtrado[(df_pre_filtrado['Riesgo_Num'] >= 1) & (df_pre_filtrado['Riesgo_Num'] <= 25)]
            txt_filtro_kpi = "🟡 Filtrado por: RIESGO BAJO (1-25%)"
        elif st.session_state['filtro_kpi_riesgo'] == "Seguro":
            df_filtrado = df_pre_filtrado[df_pre_filtrado['Riesgo_Num'] == 0]
            txt_filtro_kpi = "🟢 Filtrado por: SEGURO PURO (0%)"
        else:
            df_filtrado = df_pre_filtrado.copy()
            txt_filtro_kpi = "🔵 Mostrando: TODOS LOS REGISTROS"

        st.title("🛡️ Dashboard WAF - Monitoreo de IPs y Ataques")
        st.caption(f"📌 **Consulta Activa:** {fecha_lbl} | **Total Evaluados:** {tot_ips:,} | **{txt_filtro_kpi}** ({len(df_filtrado):,} registros mostrados)")

        k1, k2, k3, k4, k5 = st.columns(5)
        
        with k1:
            st.markdown(f'<div class="kpi-card kpi-tot"><div class="kpi-title">TOTAL REGISTROS</div><div class="kpi-number">{tot_ips:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Todos", key="btn_tot", on_click=set_filtro_kpi, args=("Todos",), use_container_width=True)

        with k2:
            st.markdown(f'<div class="kpi-card kpi-alto"><div class="kpi-title">CRÍTICO (61-100%)</div><div class="kpi-number">{criticas:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Críticos", key="btn_crit", on_click=set_filtro_kpi, args=("Crítico",), use_container_width=True)

        with k3:
            st.markdown(f'<div class="kpi-card kpi-med"><div class="kpi-title">MEDIO (26-60%)</div><div class="kpi-number">{medios:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Medios", key="btn_med", on_click=set_filtro_kpi, args=("Medio",), use_container_width=True)

        with k4:
            st.markdown(f'<div class="kpi-card kpi-bajo"><div class="kpi-title">BAJO (1-25%)</div><div class="kpi-number">{bajos:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Bajos", key="btn_bajo", on_click=set_filtro_kpi, args=("Bajo",), use_container_width=True)

        with k5:
            st.markdown(f'<div class="kpi-card kpi-seguro"><div class="kpi-title">SEGURO (0%)</div><div class="kpi-number">{seguros_puros:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Seguros (0%)", key="btn_seg", on_click=set_filtro_kpi, args=("Seguro",), use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        pct_mitigacion = round(((criticas + medios) / tot_ips) * 100, 1) if tot_ips > 0 else 0.0
        pct_visibilidad = round((df_filtrado['TIPO DE ATAQUE'].notna().sum() / len(df_filtrado)) * 100, 1) if len(df_filtrado) > 0 else 0.0
        apps_protegidas = df_filtrado['NOMBRE APLICACIÓN WEB'].nunique()

        st.markdown('<div class="section-header"><h3>📈 Indicadores de Eficiencia WAF Calculados en Tiempo Real</h3></div>', unsafe_allow_html=True)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Tasa de Bloqueo / Mitigación", f"{pct_mitigacion}%", help="Porcentaje de IPs con riesgo malicioso (>25%) mitigadas activamente")
        m2.metric("Visibilidad Inspección Capa 7", f"{pct_visibilidad}%", help="Porcentaje de peticiones clasificadas con tipo de ataque OWASP")
        m3.metric("Aplicaciones Web Protegidas", f"{apps_protegidas}", help="Número de aplicaciones institucionales monitoreadas en la matriz")
        m4.metric("Disponibilidad del Servicio", "99.9%", help="SLA de continuidad de servicios judiciales sin afectación")

        st.markdown("<br>", unsafe_allow_html=True)

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.subheader(f"Top Tipos de Ataques ({len(df_filtrado):,} eventos)")
            df_at = df_filtrado['TIPO DE ATAQUE'].value_counts().head(8).reset_index()
            df_at.columns = ['Tipo de Ataque', 'Eventos']
            fig_at = px.bar(df_at, x='Eventos', y='Tipo de Ataque', orientation='h', color='Eventos', color_continuous_scale='Reds', text='Eventos')
            fig_at.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig_at, use_container_width=True)

        with col_g2:
            st.subheader(f"Origen del Tráfico por País ({len(df_filtrado):,} IPs)")
            df_p = df_filtrado['PAÍS'].value_counts().head(8).reset_index()
            df_p.columns = ['País', 'Total IPs']
            fig_p = px.pie(df_p, names='País', values='Total IPs', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_p.update_layout(height=350, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig_p, use_container_width=True)

        st.markdown("---")
        st.subheader("🔍 Matriz Detallada de IPs Auditadas")
        cols_mostrar = [c for c in ['FechaCargaMatriz', 'IP', 'PAÍS', 'RIESGO ABUSEIPDB (%)', 'Nivel_Riesgo', 'TIPO DE ATAQUE', 'NOMBRE APLICACIÓN WEB', 'Conexión Equipo Interno', 'Dest_port'] if c in df_filtrado.columns]
        st.dataframe(df_filtrado[cols_mostrar], use_container_width=True, hide_index=True)

# =============================================================================
# PESTAÑA 2: SUSTENTACIÓN Y VISUALIZADOR DE PDF RÁPIDO + GUIÓN Y GRÁFICOS REALES
# =============================================================================
with tab_tesis:
    st.title("🎓 Defensa de Tesis de Maestría - Universidad Casa Grande")
    st.markdown("**Tema:** Implementación de un Firewall de Aplicaciones Web (WAF) para el Fortalecimiento de la Seguridad en el Acceso a Aplicaciones Judiciales")
    st.markdown("**Autor:** David Ricardo Valverde B | **Fecha:** 20 de mayo de 2026")
    st.markdown("---")

    t_pdf, t_monitoreo, t_s2, t_s3 = st.tabs([
        "📄 Presentación en Diapositivas", 
        "📊 Monitoreo Real (566 Registros & Pasteles)", 
        "⚖️ Justificación de Falsos Positivos (0%)", 
        "🗣️ Guión de Exposición Lámina por Lámina"
    ])

    # SUB-PESTAÑA 1: VISOR INTERACTIVO
    with t_pdf:
        st.subheader("📌 Visor Interactivo de Diapositivas de la Tesis")
        
        archivos_pdf = glob.glob("*.pdf")
        
        if archivos_pdf:
            pdf_path = archivos_pdf[0]
            
            doc_pdf = fitz.open(pdf_path)
            total_paginas = len(doc_pdf)

            if 'num_slide' not in st.session_state:
                st.session_state['num_slide'] = 1

            def prev_slide():
                if st.session_state['num_slide'] > 1:
                    st.session_state['num_slide'] -= 1

            def next_slide():
                if st.session_state['num_slide'] < total_paginas:
                    st.session_state['num_slide'] += 1

            col_ctrl1, col_ctrl2 = st.columns([1, 3])
            
            with col_ctrl1:
                st.markdown("#### 🔍 Control de Navegación")
                
                num_p = st.number_input(
                    "Ir a la Diapositiva N°:", 
                    min_value=1, 
                    max_value=total_paginas, 
                    value=st.session_state['num_slide'], 
                    step=1
                )
                if num_p != st.session_state['num_slide']:
                    st.session_state['num_slide'] = num_p
                    st.rerun()

                st.info(f"Mostrando lámina **{st.session_state['num_slide']}** de **{total_paginas}**.")
                
                with open(pdf_path, "rb") as f:
                    pdf_bytes_data = f.read()
                
                st.download_button(
                    label="📥 Descargar Documento PDF Oficial",
                    data=pdf_bytes_data,
                    file_name=pdf_path,
                    mime="application/pdf",
                    use_container_width=True
                )

            with col_ctrl2:
                pagina = doc_pdf.load_page(st.session_state['num_slide'] - 1)
                pix = pagina.get_pixmap(dpi=140)
                img_bytes = pix.tobytes("png")

                st.image(img_bytes, use_container_width=True)

                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    st.button("⬅️ Diapositiva Anterior", on_click=prev_slide, use_container_width=True, key="btn_slide_prev")
                with c_btn2:
                    st.button("Siguiente Diapositiva ➡️", on_click=next_slide, use_container_width=True, key="btn_slide_next")

        else:
            st.warning("⚠️ No se encontró ningún archivo `.pdf` en el repositorio.")

    # SUB-PESTAÑA 2: GRÁFICOS DE PASTEL DEL CONSOLIDADO GLOBAL (566 REGISTROS)
    with t_monitoreo:
        st.markdown('<div class="section-header"><h3>📈 Evidencia de Monitoreo Global (566 Registros Auditados)</h3></div>', unsafe_allow_html=True)
        st.caption("Análisis consolidado de la muestra total procesada en la consola F5 BIG-IQ y validada con reputación AbuseIPDB.")

        c_pie1, c_pie2 = st.columns(2)

        with c_pie1:
            st.subheader("Clasificación de Riesgo AbuseIPDB (566 Registros)")
            labels_riesgo = [
                'Seguro - 0% Riesgo (56.7%)', 
                'Crítico - 61-100% Riesgo (29.0%)', 
                'Bajo - 1-25% Riesgo (9.5%)', 
                'Medio - 26-60% Riesgo (4.8%)'
            ]
            values_riesgo = [321, 164, 54, 27]
            
            fig_p1 = go.Figure(data=[go.Pie(
                labels=labels_riesgo, 
                values=values_riesgo, 
                hole=.4,
                marker_colors=['#16A34A', '#DC2626', '#EAB308', '#D97706']
            )])
            fig_p1.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font_color='white', margin=dict(t=20, b=20, l=10, r=10))
            st.plotly_chart(fig_p1, use_container_width=True)

        with c_pie2:
            st.subheader("Tipología de Ataques en Amenazas Críticas (164 Eventos)")
            labels_ataques = [
                'Forceful Browsing (45.7%)', 
                'Buffer Overflow (18.3%)', 
                'Trojan / Backdoor (15.2%)', 
                'HTTP Parser / SSRF (7.3%)', 
                'Otros (SSRF, Predictable Resource, etc.) (13.5%)'
            ]
            values_ataques = [75, 30, 25, 12, 22]
            
            fig_p2 = go.Figure(data=[go.Pie(
                labels=labels_ataques, 
                values=values_ataques, 
                hole=.4,
                marker_colors=['#FF4C4C', '#FF7B54', '#FFB26B', '#FFD93D', '#4D96FF']
            )])
            fig_p2.update_layout(height=380, paper_bgcolor='rgba(0,0,0,0)', font_color='white', margin=dict(t=20, b=20, l=10, r=10))
            st.plotly_chart(fig_p2, use_container_width=True)

        st.markdown("""
        <div class="speech-box">
            <div class="speech-title">🎙️ Speech Sugerido para la Muestra de 566 Registros:</div>
            <i>"Auditamos una muestra de <b>566 registros de tráfico</b> en la consola F5 BIG-IQ. El <b>56.7% (321 registros)</b> correspondió a tráfico totalmente seguro (0% riesgo), garantizando que las consultas judiciales legítimas fluyeran sin interrupción. El <b>29.0% (164 registros)</b> representó amenazas de riesgo crítico (61-100%), concentradas principalmente en escaneos automatizados (Forceful Browsing con 45.7%) y desbordamiento de búfer. El 14.3% restante (81 registros) se mantuvo bajo clasificación de riesgo bajo y medio. Esto demuestra que el WAF fue 100% efectivo al mitigar los ataques críticos y transparente al permitir las transacciones seguras."</i>
        </div>
        """, unsafe_allow_html=True)

    # SUB-PESTAÑA 3: FALSOS POSITIVOS (0% AFECTACIÓN EN PRODUCCIÓN)
    with t_s2:
        st.markdown('<div class="section-header"><h3>⚖️ Justificación Técnica de Falsos Positivos & Indicadores Clave</h3></div>', unsafe_allow_html=True)
        
        st.subheader("1. Respuesta Rigurosa sobre la Tasa de Falsos Positivos")
        st.success("""
        * **¿Cual es la tasa final de Falsos Positivos?** Tras la fase de aprendizaje en *Modo Transparente* y la posterior creación de reglas de excepción (*Policy Tuning*), **la tasa de afectación por Falsos Positivos fue exactamente del 0% (100% de efectividad en la continuidad del servicio)**[cite: 5].
        * **¿Existieron Falsos Positivos?** **SÍ, pero únicamente en la fase previa en Modo Transparente**[cite: 5]. El motor del WAF asociaba erróneamente consultas complejas de la Función Judicial con firmas de ataque[cite: 5]. Mediante el *Policy Tuning*, esas falsas alarmas se eliminaron antes del despliegue en producción[cite: 5].
        * **¿Los 321 registros seguros son Falsos Positivos?** **NO**[cite: 5]. Corresponden a tráfico limpio de usuarios legítimos que el WAF inspeccionó en Capa 7 y dejó pasar a los servidores sin generar alertas ni bloqueos[cite: 5].
        """)
        
        st.subheader("2. Origen Metodológico de las Métricas (95%, 88%, 99.9%, 80%)")
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.markdown("""
            * **95% Mitigación:** Eficiencia en el bloqueo perimetral de vectores de ataque e IPs con reputación maliciosa (>25%) antes de llegar a los servidores de origen[cite: 5].
            * **88% Visibilidad:** Porcentaje de incremento en inspección Capa 7 para tráfico HTTP/HTTPS cifrado previamente no analizado[cite: 5].
            """)
        with col_res2:
            st.markdown("""
            * **99.9% Disponibilidad:** SLA de operatividad mantenido en las aplicaciones judiciales durante la implementación sin interrupción de servicios[cite: 5].
            * **80% Resp. Incidentes:** Porcentaje de reducción en el tiempo de análisis del equipo SOC para identificar la IP, puerto y tipo de ataque gracias a la centralización en BIG-IQ[cite: 5].
            """)

    # SUB-PESTAÑA 4: GUIÓN LÁMINA POR LÁMINA
    with t_s3:
        st.markdown("### 🗣️ Guión de Exposición Oral Lámina por Lámina")
        
        guion_pasos = [
            ("Lámina 1: Portada", "Estimados miembros del tribunal, buenas tardes. Presento mi proyecto de titulación de Maestría titulado: Implementación de un Firewall de Aplicaciones Web (WAF) para el fortalecimiento de la seguridad en el acceso a las aplicaciones judiciales."),
            ("Lámina 2: Contexto y Justificación", "En el ámbito judicial, la transformación digital ha convertido a las aplicaciones web en servicios críticos que manejan datos sensibles. Era necesario transicionar hacia una protección profunda en Capa 7 para mitigar riesgos cibernéticos sin degradar el rendimiento."),
            ("Lámina 3: Planteamiento del Problema", "Diagnosticamos tres vulnerabilidades clave: acceso directo sin filtrado perimetral profundo, ausencia de inspección Capa 7 bajo estándar OWASP y un riesgo crítico de interrupción de servicios judiciales ante ataques cibernéticos."),
            ("Lámina 4: Objetivos del Proyecto", "El objetivo general fue implementar y gestionar un WAF para fortalecer el acceso a las aplicaciones judiciales, desglosado en evaluar la infraestructura F5, configurar políticas de seguridad, ejecutar pruebas de inyección y optimizar el tuning dinámico."),
            ("Lámina 5: Marco Teórico", "Fundamentamos el proyecto en tres pilares internacionales: ISO/IEC 27001 para la gobernanza del SGSI, NIST CSF 2.0 para la resiliencia operativa y OWASP Top 10 como catálogo para neutralizar las principales vulnerabilidades web."),
            ("Lámina 6: Arquitectura de la Solución", "El WAF actúa como un punto de inspección bidireccional en Capa 7. El tráfico HTTP/HTTPS entrante se descifra, se analiza contra la política para bloquear solicitudes maliciosas y solo el tráfico limpio llega a los servidores de origen."),
            ("Lámina 7: Tecnologías Implementadas", "Desplegamos un ecosistema F5 de alta disponibilidad sobre hardware rSeries r5900, con instancias virtuales BIG-IP Tenants por cada aplicación crítica, administración centralizada mediante BIG-IQ y el módulo API Security."),
            ("Lámina 8: Metodología de Implementación", "Seguimos un proceso estructurado en 5 fases: Análisis del entorno, Diseño de topología, Configuración de políticas base, Pruebas de validación OWASP y Optimización/Tuning continuo."),
            ("Lámina 9: Gestión Operativa (Transparente vs. Bloqueo)", "Para no interrumpir la atención ciudadana, iniciamos en Modo Transparente. El motor de Policy Building aprendió del tráfico real, afinamos excepciones para alcanzar 0% de falsos positivos y finalmente activamos el Modo Bloqueo Activo."),
            ("Lámina 10: Pruebas de Seguridad", "Sometimos la solución a pruebas activas de SQL Injection, Cross-Site Scripting (XSS), Remote Code Execution y Escaneos Automatizados, logrando el bloqueo inmediato y registro en tiempo real en el 100% de las pruebas."),
            ("Lámina 11: Resultados de la Implementación", "Logramos 95% de mitigación de riesgos, 99.9% de disponibilidad en servicios judiciales, 88% de visibilidad del tráfico Capa 7 y 80% de reducción en el tiempo de respuesta a incidentes del equipo SOC."),
            ("Lámina 12: Conclusiones y Recomendaciones", "Concluimos que la implementación del WAF fortalece de manera medible la postura de ciberseguridad. Recomendamos mantener la actualización continua de firmas y el monitoreo especializado de APIs."),
            ("Lámina 13: Cierre y Preguntas", "Muchas gracias por su atención. Quedo a disposición de los miembros del tribunal para atender sus preguntas u observaciones.")
        ]

        for titulo_l, texto_l in guion_pasos:
            st.markdown(f"**📌 {titulo_l}**")
            st.info(f'"{texto_l}"')

# =============================================================================
# PESTAÑA 3: MÓDULO ADMINISTRADOR & CARGA DIARIA DE MATRICES
# =============================================================================
with tab_admin:
    st.title("⚙️ Gestión Diaria de Matrices & Almacenamiento Histórico")
    st.caption("Cargue diariamente el archivo Excel exportado del WAF para alimentar el histórico del sistema.")

    col_c1, col_c2 = st.columns([2, 1])

    with col_c1:
        st.markdown("### 📥 Subir Matriz Diaria")
        fecha_carga = st.date_input("🗓️ Seleccionar Fecha de Carga:", obtener_hora_quito())
        archivo_nuevo = st.file_uploader("📂 Seleccionar archivo Excel (.xlsx)", type=["xlsx"])

        if archivo_nuevo and st.button("💾 Guardar y Registrar en Histórico", use_container_width=True):
            try:
                try:
                    df_nuevo = pd.read_excel(archivo_nuevo, sheet_name='Reporte IPs')
                except Exception:
                    df_nuevo = pd.read_excel(archivo_nuevo, sheet_name=0)

                f_str = fecha_carga.strftime("%Y-%m-%d")
                df_procesado = procesar_df(df_nuevo, fecha_carga_str=f_str)

                if os.path.exists(ARCHIVO_HISTORICO):
                    df_actual = pd.read_excel(ARCHIVO_HISTORICO)
                    df_sin_fecha = df_actual[df_actual["FechaCargaMatriz"].astype(str) != f_str]
                    df_consolidado = pd.concat([df_sin_fecha, df_procesado], ignore_index=True)
                else:
                    df_consolidado = df_procesado

                df_consolidado.to_excel(ARCHIVO_HISTORICO, index=False, engine="openpyxl")
                st.cache_data.clear()
                st.balloons()
                st.success(f"✅ ¡Matriz correspondiente al {f_str} guardada con éxito! Registros acumulados: {len(df_consolidado):,}")
                time.sleep(1)
                st.rerun()

            except Exception as e:
                st.error(f"❌ Error al procesar el archivo: {e}")

    with col_c2:
        st.markdown("### 📊 Estado del Histórico")
        st.metric("Total Registros Acumulados", f"{len(df_base):,}")
        fechas_hist = df_base["FechaCargaMatriz"].unique() if not df_base.empty and "FechaCargaMatriz" in df_base.columns else []
        st.metric("Fechas Registradas", f"{len(fechas_hist)}")

    st.markdown("---")
    st.markdown("### 🗑️ Administración de Fechas")
    if len(fechas_hist) > 0:
        fecha_del = st.selectbox("Seleccionar fecha para borrar registros:", sorted([str(x) for x in fechas_hist], reverse=True))
        if st.button(f"🚨 Eliminar Definitivamente Registros del {fecha_del}"):
            df_limpio = df_base[df_base["FechaCargaMatriz"].astype(str) != fecha_del]
            df_limpio.to_excel(ARCHIVO_HISTORICO, index=False, engine="openpyxl")
            st.cache_data.clear()
            st.success(f"✅ Registros del {fecha_del} eliminados del histórico.")
            time.sleep(1)
            st.rerun()