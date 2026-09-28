import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import time
from datetime import datetime, timezone, timedelta
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Configuración de la zona horaria para Quito, Ecuador (UTC-5)
ZONA_HORARIA_QUITO = timezone(timedelta(hours=-5))

def obtener_hora_quito():
    """Obtiene la fecha y hora actual exacta en Quito, Ecuador (UTC-5)."""
    return datetime.now(ZONA_HORARIA_QUITO)

# Configuración de la página
st.set_page_config(
    page_title="Dashboard WAF - Monitoreo & Auditoría de IPs",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS institucional
st.markdown("""
    <style>
    .stApp { background-color: #F1F5F9; color: #0F172A; }
    .kpi-card { border-radius: 12px; padding: 16px 12px; text-align: center; color: white; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08); }
    .kpi-tot { background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); }
    .kpi-alto { background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); }
    .kpi-med { background: linear-gradient(135deg, #D97706 0%, #B45309 100%); }
    .kpi-bajo { background: linear-gradient(135deg, #16A34A 0%, #15803D 100%); }
    .kpi-title { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; opacity: 0.95; }
    .kpi-number { font-size: 32px; font-weight: 800; margin-top: 2px; }
    .section-header { background-color: #FFFFFF; padding: 12px 20px; border-radius: 10px; border-left: 5px solid #0284C7; box-shadow: 0 2px 4px rgba(0,0,0,0.04); margin-bottom: 15px; }
    h1, h2, h3 { color: #0F172A !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    [data-testid="stSidebar"] { background-color: #0F172A; }
    [data-testid="stSidebar"] * { color: #F8FAFC !important; }
    .author-credit { font-size: 11px; color: #64748B; font-style: italic; margin-top: -10px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_EXCEL = "resultado_ips(31-08-2026).xlsx"

# -----------------------------------------------------------------------------
# WIDGET DE CLIMA Y RELOJ DINÁMICO EN VIVO EN SIDEBAR
# -----------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------
# DICCIONARIO DE USUARIOS AUTORIZADOS
# -----------------------------------------------------------------------------
USUARIOS_AUTORIZADOS = {
    "david.valverde": {"clave": "Valverde2026!", "rol": "Administrador", "nombre": "David Ricardo Valverde Benítez"},
    "gabriela.armendariz": {"clave": "GArmendariz2026!", "rol": "Administrador", "nombre": "Gabriela Armendariz"},
    "admin": {"clave": "admin123", "rol": "Administrador", "nombre": "Administrador Principal"}
}

# -----------------------------------------------------------------------------
# CONTROL DE SESIÓN
# -----------------------------------------------------------------------------
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<h1>🔒 Control de Acceso WAF</h1>', unsafe_allow_html=True)
        st.markdown('<p class="author-credit">Maestría en Ciberseguridad - UCG</p>', unsafe_allow_html=True)
        st.caption("Ingrese sus credenciales para consultar el Monitoreo de IPs.")
        
        with st.form("form_login"):
            usr_input = st.text_input("👤 Usuario:")
            pwd_input = st.text_input("🔑 Contraseña:", type="password")
            btn_login = st.form_submit_button("Iniciar Sesión", use_container_width=True)
            
            if btn_login:
                usr_clean = usr_input.strip().lower()
                if usr_clean in USUARIOS_AUTORIZADOS and USUARIOS_AUTORIZADOS[usr_clean]["clave"] == pwd_input:
                    st.session_state['autenticado'] = True
                    st.session_state['nombre_actual'] = USUARIOS_AUTORIZADOS[usr_clean]["nombre"]
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
    st.stop()

# -----------------------------------------------------------------------------
# CARGA ROBUSTA Y ANTI-ERRORES DEL EXCEL
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1)
def cargar_datos_ips():
    path = ARCHIVO_EXCEL if os.path.exists(ARCHIVO_EXCEL) else None
    if not path:
        return pd.DataFrame()
    
    try:
        # Intentar leer la hoja específica 'Reporte IPs'
        df = pd.read_excel(path, sheet_name='Reporte IPs')
    except Exception:
        # Si falla por el nombre de la hoja, lee la primera hoja por índice
        df = pd.read_excel(path, sheet_name=0)
        
    if df.empty:
        return pd.DataFrame()

    # Limpiar espacios en blanco al inicio/final de los nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]

    # Mapeo flexible e inteligente para detectar columnas sin importar tildes o mayúsculas
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

    # Garantizar que existan las columnas mínimas esperadas
    if 'PAÍS' not in df.columns:
        df['PAÍS'] = 'Desconocido'
    if 'TIPO DE ATAQUE' not in df.columns:
        df['TIPO DE ATAQUE'] = 'No especificado'
    if 'NOMBRE APLICACIÓN WEB' not in df.columns:
        df['NOMBRE APLICACIÓN WEB'] = 'General'

    # Procesar porcentaje de riesgo a un número entero/flotante
    if 'RIESGO ABUSEIPDB (%)' in df.columns:
        df['Riesgo_Num'] = df['RIESGO ABUSEIPDB (%)'].astype(str).str.replace('%', '').str.strip()
        df['Riesgo_Num'] = pd.to_numeric(df['Riesgo_Num'], errors='coerce').fillna(0)
    else:
        df['Riesgo_Num'] = 0

    def clasificar_riesgo(val):
        if val == 0: return 'Seguro (0%)'
        elif val <= 25: return 'Bajo (1-25%)'
        elif val <= 60: return 'Medio (26-60%)'
        else: return 'Crítico (61-100%)'

    df['Nivel_Riesgo'] = df['Riesgo_Num'].apply(clasificar_riesgo)
    return df

df_ips = cargar_datos_ips()

# -----------------------------------------------------------------------------
# GENERACIÓN DE INFORME EN WORD (.DOCX)
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def generar_reporte_word_waf(df_data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    COLOR_AZUL = RGBColor(2, 132, 199)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("INFORME AUDITORÍA Y MONITOREO WAF - IPS BLOQUEADAS\nAPLICACIONES JUDICIALES")
    r_title.font.name = 'Segoe UI'
    r_title.font.size = Pt(15)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_AZUL

    doc.add_heading("1. Resumen de Tráfico e IPs Escaneadas", level=1)
    tot = len(df_data)
    criticas = (df_data['Riesgo_Num'] > 60).sum() if not df_data.empty else 0
    
    doc.add_paragraph(f"Se presenta el análisis detallado del monitoreo de peticiones e IPs registradas en la solución WAF F5. Se auditó un total de {tot:,} eventos/IPs de origen, detectando {criticas} eventos con nivel de riesgo crítico (>60% en AbuseIPDB).")
    
    doc.add_heading("2. Detalle por Categoría de Ataque", level=1)
    if 'TIPO DE ATAQUE' in df_data.columns and not df_data.empty:
        resumen_ataque = df_data['TIPO DE ATAQUE'].value_counts().reset_index()
        resumen_ataque.columns = ["Tipo de Ataque", "Cantidad"]
        
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, h in enumerate(["Tipo de Ataque", "Total Registros"]):
            cell = table.rows[0].cells[i]
            set_cell_background(cell, "0284C7")
            p = cell.paragraphs[0]
            r = p.add_run(h)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            
        for _, r_row in resumen_ataque.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(r_row["Tipo de Ataque"])
            row_cells[1].text = str(r_row["Cantidad"])

    stream = io.BytesIO()
    doc.save(stream)
    return stream.getvalue()

# -----------------------------------------------------------------------------
# BARRA LATERAL (SIDEBAR) & FILTROS
# -----------------------------------------------------------------------------
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.get('nombre_actual')}")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state['autenticado'] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.title("📌 Filtros de Monitoreo")

if df_ips.empty or 'PAÍS' not in df_ips.columns:
    st.error("⚠️ No se pudo cargar el archivo Excel. Verifica que 'resultado_ips(31-08-2026).xlsx' esté subido en el repositorio GitHub.")
    st.stop()

# Filtros por desplegable
paises = ["Todos"] + sorted([str(x) for x in df_ips['PAÍS'].dropna().unique() if str(x).strip() != ""])
pais_sel = st.sidebar.selectbox("Filtrar por País de Origen:", paises)

apps = ["Todas"] + sorted([str(x) for x in df_ips['NOMBRE APLICACIÓN WEB'].dropna().unique() if str(x).strip() != ""])
app_sel = st.sidebar.selectbox("Filtrar por Aplicación Web:", apps)

# Aplicar filtros
df_filtrado = df_ips.copy()
if pais_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado['PAÍS'] == pais_sel]
if app_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado['NOMBRE APLICACIÓN WEB'] == app_sel]

# Botón de Descarga del Reporte Word
st.sidebar.markdown("---")
st.sidebar.subheader("📄 Generar Reporte")
doc_bytes = generar_reporte_word_waf(df_filtrado)
st.sidebar.download_button(
    "📄 Descargar Informe WAF (Word)",
    data=doc_bytes,
    file_name=f"Informe_Monitoreo_WAF_{obtener_hora_quito().strftime('%Y%m%d')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True
)

# -----------------------------------------------------------------------------
# PANTALLA PRINCIPAL (DASHBOARD)
# -----------------------------------------------------------------------------
st.title("🛡️ Dashboard WAF - Monitoreo de IPs y Ataques")
st.caption("Visualización interactiva de eventos de seguridad registrados por la plataforma F5 WAF")

# MÉTIRICAS / KPIS PRINCIPALES
tot_ips = len(df_filtrado)
criticas = len(df_filtrado[df_filtrado['Riesgo_Num'] > 60])
medios = len(df_filtrado[(df_filtrado['Riesgo_Num'] >= 26) & (df_filtrado['Riesgo_Num'] <= 60)])
seguros = len(df_filtrado[df_filtrado['Riesgo_Num'] <= 25])

k1, k2, k3, k4 = st.columns(4)
k1.markdown(f'<div class="kpi-card kpi-tot"><div class="kpi-title">Total Registros / IPs</div><div class="kpi-number">{tot_ips:,}</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="kpi-card kpi-alto"><div class="kpi-title">Riesgo Crítico (61-100%)</div><div class="kpi-number">{criticas:,}</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="kpi-card kpi-med"><div class="kpi-title">Riesgo Medio (26-60%)</div><div class="kpi-number">{medios:,}</div></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="kpi-card kpi-bajo"><div class="kpi-title">Riesgo Bajo/Seguro (0-25%)</div><div class="kpi-number">{seguros:,}</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# GRÁFICOS INTERACTIVOS
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("Top Tipos de Ataques Detectados")
    df_ataques = df_filtrado['TIPO DE ATAQUE'].value_counts().head(8).reset_index()
    df_ataques.columns = ['Tipo de Ataque', 'Eventos']
    fig_at = px.bar(
        df_ataques, x='Eventos', y='Tipo de Ataque', orientation='h',
        color='Eventos', color_continuous_scale='Reds', text='Eventos'
    )
    fig_at.update_layout(height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig_at, use_container_width=True)

with col_g2:
    st.subheader("Origen del Tráfico por País")
    df_pais = df_filtrado['PAÍS'].value_counts().head(8).reset_index()
    df_pais.columns = ['País', 'Total IPs']
    fig_p = px.pie(
        df_pais, names='País', values='Total IPs', hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig_p.update_layout(height=320, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig_p, use_container_width=True)

st.markdown("---")

# MATRIZ COMPLETA DE DATOS
st.subheader("🔍 Matriz Completa de IPs Auditadas")
cols_mostrar = [c for c in ['IP', 'PAÍS', 'RIESGO ABUSEIPDB (%)', 'Nivel_Riesgo', 'TIPO DE ATAQUE', 'NOMBRE APLICACIÓN WEB', 'Conexión Equipo Interno', 'Dest_port'] if c in df_filtrado.columns]
st.dataframe(df_filtrado[cols_mostrar], use_container_width=True, hide_index=True)