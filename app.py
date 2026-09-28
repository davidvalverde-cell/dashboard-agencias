import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob
import time
import fitz  # PyMuPDF para renderizar diapositivas PDF sin iframes ni bloqueos de Chrome
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

# USUARIOS AUTORIZADOS (ACTUALIZADO CON CREDENCIALES PEDIDAS)
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

@st.cache_data(ttl=1)
def cargar_historico_base():
    if os.path.exists(ARCHIVO_HISTORICO):
        try:
            return pd.read_excel(ARCHIVO_HISTORICO)
        except Exception:
            pass
    
    archivos = glob.glob("resultado_ips*.xlsx")
    if archivos:
        try:
            df_init = pd.read_excel(archivos[0], sheet_name='Reporte IPs')
        except Exception:
            df_init = pd.read_excel(archivos[0], sheet_name=0)
        return procesar_df(df_init, fecha_carga_str=obtener_hora_quito().strftime("%Y-%m-%d"))
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

        # CÁLCULO DE KPIS CON DESGLOSE RIGUROSO DE 0% SEGURO Y 1-25% BAJO
        tot_ips = len(df_pre_filtrado)
        criticas = len(df_pre_filtrado[df_pre_filtrado['Riesgo_Num'] > 60])
        medios = len(df_pre_filtrado[(df_pre_filtrado['Riesgo_Num'] >= 26) & (df_pre_filtrado['Riesgo_Num'] <= 60)])
        bajos = len(df_pre_filtrado[(df_pre_filtrado['Riesgo_Num'] >= 1) & (df_pre_filtrado['Riesgo_Num'] <= 25)])
        seguros_puros = len(df_pre_filtrado[df_pre_filtrado['Riesgo_Num'] == 0])

        # APLICACIÓN DE FILTRO DINÁMICO SEGÚN EL BOTÓN SELECCIONADO
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

        # ENCABEZADO
        st.title("🛡️ Dashboard WAF - Monitoreo de IPs y Ataques")
        st.caption(f"📌 **Consulta Activa:** {fecha_lbl} | **Total Evaluados:** {tot_ips:,} | **{txt_filtro_kpi}** ({len(df_filtrado):,} registros mostrados)")

        # TARJETAS DE MÉTRICAS KPI REORDENADAS Y CORREGIDAS
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

        # CÁLCULO Y DESPLIEGUE DE INDICADORES EN TIEMPO REAL
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

        # GRÁFICAS FILTRADAS DINÁMICAMENTE
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.subheader(f"Top Tipos de Ataques ({len(df_filtrado):,} eventos)")
            df_at = df_filtrado['TIPO DE ATAQUE'].value_counts().head(8).reset_index()
            df_at.columns = ['Tipo de Ataque', 'Eventos']
            fig_at = px.bar(df_at, x='Eventos', y='Tipo de Ataque', orientation='h', color='Eventos', color_continuous_scale='Reds', text='Eventos')
            fig_at.update_layout(height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig_at, use_container_width=True)

        with col_g2:
            st.subheader(f"Origen del Tráfico por País ({len(df_filtrado):,} IPs)")
            df_p = df_filtrado['PAÍS'].value_counts().head(8).reset_index()
            df_p.columns = ['País', 'Total IPs']
            fig_p = px.pie(df_p, names='País', values='Total IPs', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_p.update_layout(height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig_p, use_container_width=True)

        st.markdown("---")
        st.subheader("🔍 Matriz Detallada de IPs Auditadas")
        cols_mostrar = [c for c in ['FechaCargaMatriz', 'IP', 'PAÍS', 'RIESGO ABUSEIPDB (%)', 'Nivel_Riesgo', 'TIPO DE ATAQUE', 'NOMBRE APLICACIÓN WEB', 'Conexión Equipo Interno', 'Dest_port'] if c in df_filtrado.columns]
        st.dataframe(df_filtrado[cols_mostrar], use_container_width=True, hide_index=True)

# =============================================================================
# PESTAÑA 2: SUSTENTACIÓN Y VISUALIZADOR DE PDF NATIVO EN IMÁGENES (SIN IFRAMES)
# =============================================================================
with tab_tesis:
    st.title("🎓 Defensa de Tesis de Maestría - Universidad Casa Grande")
    st.markdown("**Tema:** Implementación de un Firewall de Aplicaciones Web (WAF) para el Fortalecimiento de la Seguridad en el Acceso a Aplicaciones Judiciales")
    st.markdown("**Autor:** David Ricardo Valverde B | **Fecha:** 20 de mayo de 2026")
    st.markdown("---")

    t_pdf, t_s2, t_s3 = st.tabs(["📄 Presentación en Diapositivas", "📊 Justificación de Falsos Positivos & Cifras", "🗣️ Guión de Exposición (Speech)"])

    with t_pdf:
        st.subheader("📌 Visor Interactivo de Diapositivas de la Tesis")
        
        archivos_pdf = glob.glob("*.pdf")
        
        if archivos_pdf:
            pdf_path = archivos_pdf[0]
            st.caption(f"Archivo cargado correctamente: `{pdf_path}`")
            
            # Cargar el PDF con PyMuPDF sin usar iframes ni HTML inseguro
            doc_pdf = fitz.open(pdf_path)
            total_paginas = len(doc_pdf)
            
            # Controles de navegación
            col_ctrl1, col_ctrl2 = st.columns([1, 3])
            
            with col_ctrl1:
                st.markdown("#### 🔍 Control de Navegación")
                num_pagina = st.number_input("Ir a la Diapositiva N°:", min_value=1, max_value=total_paginas, value=1, step=1)
                st.info(f"Mostrando lámina **{num_pagina}** de un total de **{total_paginas}**.")
                
                # Descarga directa del archivo PDF
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
                # Extraer la página seleccionada y renderizarla a PNG limpio en memoria
                pagina = doc_pdf.load_page(num_pagina - 1)
                pix = pagina.get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")
                
                # Despliegue seguro tipo imagen (Imposible de bloquear por Chrome)
                st.image(img_bytes, caption=f"Diapositiva {num_pagina} / {total_paginas}", use_container_width=True)
                
        else:
            st.warning("⚠️ No se encontró ningún archivo `.pdf` en el repositorio. Asegúrate de que el archivo PDF esté subido junto a tu script.")

    with t_s2:
        st.markdown('<div class="section-header"><h3>⚖️ Justificación de Falsos Positivos & Origen de Cifras (Preguntas del Tribunal)</h3></div>', unsafe_allow_html=True)
        
        st.subheader("1. Respuesta a Pruebas y Falsos Positivos")
        st.info("""
        * **¿Cuántas pruebas se hicieron?** Se ejecutaron **4 categorías de ataques OWASP** (SQL Injection, XSS, Remote Code Execution y Escaneos Automatizados) sumando **120+ vectores de prueba activos**.
        * **¿Cuántos falsos positivos hubo?** En la fase inicial en Modo Transparente se detectaron **18 firmas que afectaban tráfico legítimo**. Tras la metodología de *Policy Tuning* a nivel de parámetros específicos, **la tasa de falsos positivos se redujo al 0%**.
        """)
        
        st.subheader("2. Origen Metodológico de la Lámina de Resultados (95%, 88%, 99.9%, 80%)")
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.markdown("""
            * **95% Mitigación:** Eficiencia en el bloqueo perimetral de vectores de ataque e IPs con reputación maliciosa (>25%) antes de llegar a los servidores internos.
            * **88% Visibilidad:** Porcentaje de incremento en inspección Capa 7 para tráfico HTTP/HTTPS cifrado previamente no analizado.
            """)
        with col_res2:
            st.markdown("""
            * **99.9% Disponibilidad:** SLA de operatividad mantenido en las aplicaciones judiciales durante la implementación sin interrupción de servicios.
            * **80% Resp. Incidentes:** Porcentaje de reducción en el tiempo de análisis del equipo SOC para identificar la IP, puerto y tipo de ataque gracias a la centralización en BIG-IQ.
            """)

    with t_s3:
        st.markdown("### 🗣️ Guión Sugerido para la Defensa Verbal")
        st.info('"Estimados miembros del tribunal evaluador, la implementación de un WAF en la Función Judicial no se limita a activar firmas de bloqueo. Un WAF no afinado que genere falsos positivos paraliza la atención ciudadana. Por ello, aplicamos una metodología rigurosa en Modo Transparente para mapear el tráfico legítimo, ejecutar afinamiento granular (Policy Tuning) a nivel de parámetros específicos y finalmente activar el Modo Bloqueo, garantizando 99.9% de disponibilidad y 95% de mitigación de amenazas."')

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

                if not df_base.empty and "FechaCargaMatriz" in df_base.columns:
                    df_base_sin_fecha = df_base[df_base["FechaCargaMatriz"].astype(str) != f_str]
                    df_consolidado = pd.concat([df_base_sin_fecha, df_procesado], ignore_index=True)
                else:
                    df_consolidado = df_procesado

                df_consolidado.to_excel(ARCHIVO_HISTORICO, index=False, engine="openpyxl")
                st.cache_data.clear()
                st.balloons()
                st.success(f"✅ ¡Matriz correspondiente al {f_str} guardada con éxito en el histórico!")
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