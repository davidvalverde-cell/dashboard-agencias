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

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE ZONA HORARIA Y PÁGINA
# -----------------------------------------------------------------------------
ZONA_HORARIA_QUITO = timezone(timedelta(hours=-5))

def obtener_hora_quito():
    """Obtiene la fecha y hora actual exacta en Quito, Ecuador (UTC-5)."""
    return datetime.now(ZONA_HORARIA_QUITO)

st.set_page_config(
    page_title="Dashboard de Seguridad SAST & SCA - SIGPRE CJ",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Corporativos y Tarjetas KPI por Severidad
st.markdown("""
    <style>
    .stApp { background-color: #F1F5F9; color: #0F172A; }
    .kpi-card { border-radius: 12px; padding: 16px 12px; text-align: center; color: white; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08); }
    .kpi-total { background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); }
    .kpi-alta { background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); }
    .kpi-media { background: linear-gradient(135deg, #F59E0B 0%, #B45309 100%); }
    .kpi-baja { background: linear-gradient(135deg, #0D9488 0%, #115E59 100%); }
    .kpi-gate { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #DC2626; }
    .kpi-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; opacity: 0.95; }
    .kpi-number { font-size: 32px; font-weight: 800; margin-top: 2px; }
    .section-header { background-color: #FFFFFF; padding: 12px 20px; border-radius: 10px; border-left: 5px solid #0284C7; box-shadow: 0 2px 4px rgba(0,0,0,0.04); margin-bottom: 15px; }
    h1, h2, h3 { color: #0F172A !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    [data-testid="stSidebar"] { background-color: #0F172A; }
    [data-testid="stSidebar"] * { color: #F8FAFC !important; }

    .author-credit {
        font-size: 11px;
        color: #64748B;
        font-style: italic;
        margin-top: -10px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_EXCEL = "Matriz_Ejecutiva_SAST_CODEC_v1.xlsx"
ARCHIVO_LOGS = "registro_ingresos.csv"

# -----------------------------------------------------------------------------
# WIDGET DE CLIMA Y RELOJ DINÁMICO EN VIVO EN SIDEBAR
# -----------------------------------------------------------------------------
reloj_html = """
<div style="
    background: rgba(255, 255, 255, 0.08);
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: 10px;
    padding: 10px 12px;
    text-align: center;
    font-family: 'Segoe UI', Roboto, sans-serif;
">
    <div style="font-size: 12px; font-weight: 700; color: #38BDF8; text-transform: uppercase; letter-spacing: 0.5px;">
        📍 QUITO, ECUADOR ⛅
    </div>
    <div id="reloj-quito" style="font-size: 22px; font-weight: 800; color: #F8FAFC; margin: 4px 0;">
        ⏰ 00:00:00
    </div>
    <div id="fecha-quito" style="font-size: 11px; color: #94A3B8;">
        📅 00/00/0000 | 18°C Parcialmente Nublado
    </div>
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
    "gabriela.armendariz": {"clave": "GArmendariz2026!", "rol": "Administrador", "nombre": "Gabriela Armendariz"},
    "usuario.generico": {"clave": "Lectura2026!", "rol": "Lectura", "nombre": "Usuario Consulta (Solo Lectura)"},
    "admin": {"clave": "admin123", "rol": "Administrador", "nombre": "Administrador Principal"}
}

# -----------------------------------------------------------------------------
# FUNCIONES AUXILIARES DE GUARDADO Y AUDITORÍA
# -----------------------------------------------------------------------------
def guardar_excel_seguro(df: pd.DataFrame, filepath: str = ARCHIVO_EXCEL) -> bool:
    """Guarda la matriz actualizada en disco."""
    temp_filepath = filepath.replace(".xlsx", "_temp.xlsx") if filepath.endswith(".xlsx") else f"{filepath}_temp.xlsx"
    try:
        df.to_excel(temp_filepath, index=False, engine="openpyxl")
        if os.path.exists(temp_filepath):
            os.replace(temp_filepath, filepath)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error al guardar el archivo en disco: {e}")
        if os.path.exists(temp_filepath):
            try: os.remove(temp_filepath)
            except Exception: pass
        return False

def registrar_ingreso(usuario, rol):
    """Registra inicios de sesión utilizando la hora exacta de Quito, Ecuador."""
    fecha_hora_quito = obtener_hora_quito().strftime("%Y-%m-%d %H:%M:%S")
    nuevo_log = pd.DataFrame([{"Usuario": usuario, "Rol": rol, "FechaHora": fecha_hora_quito}])
    
    try:
        header_needed = not os.path.exists(ARCHIVO_LOGS) or os.path.getsize(ARCHIVO_LOGS) == 0
        nuevo_log.to_csv(ARCHIVO_LOGS, mode='a', header=header_needed, index=False)
    except Exception as e:
        st.error(f"Error al registrar log de usuario: {e}")

# -----------------------------------------------------------------------------
# CONTROL DE SESIÓN
# -----------------------------------------------------------------------------
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False
if 'usuario_actual' not in st.session_state:
    st.session_state['usuario_actual'] = ""

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown('<h1>🔒 Control de Acceso - SAST / SCA</h1>', unsafe_allow_html=True)
        st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
        st.caption("Ingrese sus credenciales autorizadas para consultar el Dashboard de Seguridad de Código Fuente.")
        
        with st.form("form_login"):
            usr_input = st.text_input("👤 Usuario:")
            pwd_input = st.text_input("🔑 Contraseña:", type="password")
            btn_login = st.form_submit_button("Iniciar Sesión", use_container_width=True)
            
            if btn_login:
                usr_clean = usr_input.strip().lower()
                if usr_clean in USUARIOS_AUTORIZADOS and USUARIOS_AUTORIZADOS[usr_clean]["clave"] == pwd_input:
                    st.session_state['autenticado'] = True
                    st.session_state['usuario_actual'] = usr_clean
                    st.session_state['rol_actual'] = USUARIOS_AUTORIZADOS[usr_clean]["rol"]
                    st.session_state['nombre_actual'] = USUARIOS_AUTORIZADOS[usr_clean]["nombre"]
                    
                    registrar_ingreso(usr_clean, USUARIOS_AUTORIZADOS[usr_clean]["rol"])
                    st.success(f"Bienvenido/a {USUARIOS_AUTORIZADOS[usr_clean]['nombre']}")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos. Verifique sus credenciales.")
    st.stop()

# -----------------------------------------------------------------------------
# LECTURA DE MATRIZ EXCEL DE SEGURIDAD SAST
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1)
def cargar_datos_sast():
    path = ARCHIVO_EXCEL if os.path.exists(ARCHIVO_EXCEL) else "analisis/Matriz_Ejecutiva_SAST_CODEC_v1.xlsx"
    if not os.path.exists(path):
        return {}, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
        df_resumen = pd.read_excel(xls, sheet_name="Resumen Ejecutivo", header=2)
        df_hallazgos = pd.read_excel(xls, sheet_name="Matriz Hallazgos", header=2)
        df_semgrep = pd.read_excel(xls, sheet_name="Triaje Semgrep", header=2)
        df_plan = pd.read_excel(xls, sheet_name="Plan Remediación", header=2)
        df_tech = pd.read_excel(xls, sheet_name="Detalle Técnico", header=2)
        
        # Limpieza de nombres de columnas
        for df in [df_hallazgos, df_semgrep, df_plan, df_tech]:
            if not df.empty:
                df.columns = [str(c).strip() for c in df.columns]

        if "FechaCargaMatriz" not in df_hallazgos.columns:
            df_hallazgos["FechaCargaMatriz"] = "2026-09-11"

        return df_hallazgos, df_semgrep, df_plan, df_tech
    except Exception as e:
        st.error(f"Error al leer la matriz de seguridad Excel: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

df_hallazgos, df_semgrep, df_plan, df_tech = cargar_datos_sast()

if 'sev_activa' not in st.session_state: st.session_state['sev_activa'] = "Todas"
if 'modulo_activo' not in st.session_state: st.session_state['modulo_activo'] = "Todos"
if 'filtro_kpi' not in st.session_state: st.session_state['filtro_kpi'] = "Todas"

def set_filtro_kpi(valor):
    st.session_state['filtro_kpi'] = valor

# -----------------------------------------------------------------------------
# GENERACIÓN DE INFORME WORD EN FORMATO EJECUTIVO DE SEGURIDAD
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generar_reporte_word_detallado(df_data, df_tech_data, sev_filtro, mod_filtro, fecha_matriz):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    COLOR_NAVY = RGBColor(15, 23, 42)
    COLOR_AZUL = RGBColor(2, 132, 199)
    HEX_HEADER = "0F172A"
    HEX_ALT = "F8FAFC"

    # Título Principal
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("INFORME DE EVALUACIÓN DE SEGURIDAD ESTÁTICA Y AUDITORÍA DE CÓDIGO FUENTE (SAST / SCA)\nPROYECTO SIGPRE INTERINSTITUCIONAL")
    r_title.font.name = 'Segoe UI'
    r_title.font.size = Pt(15)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sub = p_sub.add_run(f"Fecha de emisión: {obtener_hora_quito().strftime('%d/%m/%Y %H:%M')}\nElaborado por: Gabriela Armendariz / Subdirección de Seguridad Informática\nFiltro Aplicado: Severidad [{sev_filtro}] | Módulo [{mod_filtro}]")
    r_sub.font.size = Pt(9)
    r_sub.font.italic = True

    # 1. Resumen Ejecutivo
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Resumen Ejecutivo y Cuadro de Mando")
    r_h1.font.name = 'Segoe UI'
    r_h1.font.color.rgb = COLOR_NAVY

    tot_registros = len(df_data) if not df_data.empty else 0
    tot_altas = (df_data["Severidad final"] == "ALTA").sum() if not df_data.empty else 0
    tot_medias = (df_data["Severidad final"] == "MEDIA").sum() if not df_data.empty else 0
    tot_bajas = (df_data["Severidad final"] == "BAJA").sum() if not df_data.empty else 0

    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body1.add_run(
        f"El presente informe consolida los resultados de la auditoría de seguridad realizada sobre el código fuente "
        f"de la aplicación 'SIGPRE Interinstitucional' (submódulo 'sigpre-carga-inicial-sigpre' | Laravel 10.50.3). "
        f"En la evaluación seleccionada (Filtro Severidad: '{sev_filtro}', Módulo: '{mod_filtro}') "
        f"se registran {tot_registros} hallazgos consolidados: {tot_altas} de severidad ALTA (P1), "
        f"{tot_medias} de severidad MEDIA (P2) y {tot_bajas} de severidad BAJA (P3). "
        f"El estado del Quality Gate General es FAILED CONTROLADO (No apto para Producción sin remediación)."
    )

    table_kpi = doc.add_table(rows=2, cols=4)
    table_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    headers_kpi = ["Total Hallazgos", "Severidad ALTA", "Severidad MEDIA", "Severidad BAJA"]
    values_kpi = [f"{tot_registros}", f"{tot_altas}", f"{tot_medias}", f"{tot_bajas}"]

    for i, h in enumerate(headers_kpi):
        cell = table_kpi.cell(0, i)
        set_cell_background(cell, HEX_HEADER)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for i, v in enumerate(values_kpi):
        cell = table_kpi.cell(1, i)
        set_cell_background(cell, HEX_ALT)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(v)
        r.font.bold = True
        r.font.size = Pt(12)
        r.font.color.rgb = COLOR_AZUL

    doc.add_paragraph()

    # 2. Matriz de Vulnerabilidades
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Guía Detallada de Vulnerabilidades y Remediación")
    r_h2.font.name = 'Segoe UI'
    r_h2.font.color.rgb = COLOR_NAVY

    if not df_data.empty:
        table_hallazgos = doc.add_table(rows=1, cols=4)
        table_hallazgos.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_cells = table_hallazgos.rows[0].cells
        cols_t = ["Código / ID", "Hallazgo", "Severidad", "Archivo / Ubicación"]
        for i, title in enumerate(cols_t):
            set_cell_background(hdr_cells[i], HEX_HEADER)
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in df_data.iterrows():
            row_cells = table_hallazgos.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            for col_idx, val in enumerate([row.get("Código", "N/A"), row.get("Hallazgo", "N/A"), row.get("Severidad final", "N/A"), str(row.get("Archivo / endpoint", "N/A"))[:45]]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx in [1, 3] else WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(str(val))

    doc.add_paragraph()

    # 3. Detalle Técnico
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Detalle Técnico y Componentes de Arquitectura")
    r_h3.font.name = 'Segoe UI'
    r_h3.font.color.rgb = COLOR_NAVY

    if not df_tech_data.empty:
        table_tech = doc.add_table(rows=1, cols=3)
        table_tech.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_type = table_tech.rows[0].cells
        hdr_titles = ["Categoría", "Tecnología / Librería", "Versión Identificada"]
        for i, title in enumerate(hdr_titles):
            set_cell_background(hdr_type[i], HEX_HEADER)
            p = hdr_type[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in df_tech_data.iterrows():
            row_cells = table_tech.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            for col_idx, val in enumerate([row.get("Categoría", "N/A"), row.get("Tecnología / Librería", "N/A"), row.get("Versión Identificada", "N/A")]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx < 2 else WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(str(val))

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# -----------------------------------------------------------------------------
# BARRA LATERAL
# -----------------------------------------------------------------------------
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.get('nombre_actual', 'Invitado')}")
st.sidebar.caption(f"Rol: {st.session_state.get('rol_actual', 'General')}")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state['autenticado'] = False
    st.rerun()

st.sidebar.markdown('<p class="author-credit" style="color:#94A3B8;">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
st.sidebar.markdown("---")

# PESTAÑAS PRINCIPALES
if st.session_state.get('rol_actual') == "Administrador":
    tab_dash, tab_admin = st.tabs(["📊 Dashboard SAST & Seguridad", "⚙️ Módulo Administrador & Logs"])
else:
    tab_dash, = st.tabs(["📊 Dashboard SAST & Seguridad"])

# =============================================================================
# PESTAÑA 1: DASHBOARD DE SEGURIDAD SAST
# =============================================================================
with tab_dash:
    if df_hallazgos.empty:
        st.warning("⚠️ No se ha encontrado la matriz de seguridad cargada. Por favor verifique el archivo Excel.")
    else:
        st.sidebar.title("📌 Filtros de Auditoría")

        fechas_cargas = sorted([str(x) for x in df_hallazgos["FechaCargaMatriz"].unique() if pd.notnull(x)], reverse=True)
        opciones_matriz = ["Último Escaneo (2026-09-11)"] + fechas_cargas
        matriz_sel = st.sidebar.selectbox("📂 Seleccionar Auditoría:", opciones_matriz)

        df_matriz_activa = df_hallazgos.copy()

        # Filtro por Severidad
        severidades = ["Todas", "ALTA", "MEDIA", "BAJA"]
        idx_sev = severidades.index(st.session_state['sev_activa']) if st.session_state['sev_activa'] in severidades else 0
        sev_sel = st.sidebar.selectbox("Filtrar por Severidad:", severidades, index=idx_sev)

        if sev_sel != st.session_state['sev_activa']:
            st.session_state['sev_activa'] = sev_sel
            st.session_state['modulo_activo'] = "Todos"
            st.session_state['filtro_kpi'] = "Todas"
            st.rerun()

        df_sev = df_matriz_activa[df_matriz_activa["Severidad final"].astype(str) == st.session_state['sev_activa']] if st.session_state['sev_activa'] != "Todas" else df_matriz_activa.copy()

        # Filtro por Módulo
        modulos_col = "Módulo / componente" if "Módulo / componente" in df_sev.columns else df_sev.columns[5]
        modulos = ["Todos"] + sorted([str(x) for x in df_sev[modulos_col].dropna().unique() if str(x).strip() != ""])
        idx_mod = modulos.index(st.session_state['modulo_activo']) if st.session_state['modulo_activo'] in modulos else 0
        mod_sel = st.sidebar.selectbox("Filtrar por Módulo/Componente:", modulos, index=idx_mod)

        if mod_sel != st.session_state['modulo_activo']:
            st.session_state['modulo_activo'] = mod_sel
            st.session_state['filtro_kpi'] = "Todas"
            st.rerun()

        if st.sidebar.button("🔄 Restablecer Filtros"):
            st.session_state['sev_activa'] = "Todas"
            st.session_state['modulo_activo'] = "Todos"
            st.session_state['filtro_kpi'] = "Todas"
            st.rerun()

        st.sidebar.markdown("---")
        st.sidebar.subheader("📄 Reporte Oficial SAST")

        df_filtrado = df_matriz_activa.copy()
        if st.session_state['sev_activa'] != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Severidad final"].astype(str) == st.session_state['sev_activa']]
        if st.session_state['modulo_activo'] != "Todos":
            df_filtrado = df_filtrado[df_filtrado[modulos_col].astype(str) == st.session_state['modulo_activo']]

        # BOTÓN DE DESCARGA WORD OFICIAL
        doc_bytes = generar_reporte_word_detallado(
            df_filtrado, df_tech,
            st.session_state['sev_activa'], st.session_state['modulo_activo'], "2026-09-11"
        )
        st.sidebar.download_button(
            label="📄 Descargar Informe Detallado (Word)",
            data=doc_bytes,
            file_name=f"Informe_Tecnico_SAST_SIGPRE_{st.session_state['sev_activa']}_{obtener_hora_quito().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

        # MÉTRICAS Y TARJETAS KPI
        tot_hallazgos = len(df_filtrado)
        tot_altas = (df_filtrado["Severidad final"] == "ALTA").sum()
        tot_medias = (df_filtrado["Severidad final"] == "MEDIA").sum()
        tot_bajas = (df_filtrado["Severidad final"] == "BAJA").sum()

        st.title("🛡️ Dashboard de Seguridad SAST & SCA")
        st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
        st.caption(f"📌 **Proyecto:** SIGPRE Interinstitucional - Consejo de la Judicatura | **Quality Gate:** FAILED CONTROLADO (Riesgo Global: ALTO)")

        # TARJETAS KPI INTERACTIVAS
        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            st.markdown(f'<div class="kpi-card kpi-total"><div class="kpi-title">Total Hallazgos</div><div class="kpi-number">{tot_hallazgos}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Ver Todos", key="btn_tot", on_click=set_filtro_kpi, args=("Todas",), use_container_width=True)

        with k2:
            st.markdown(f'<div class="kpi-card kpi-alta"><div class="kpi-title">Severidad ALTA</div><div class="kpi-number">{tot_altas}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Filtrar ALTAS", key="btn_alt", on_click=set_filtro_kpi, args=("ALTA",), use_container_width=True)

        with k3:
            st.markdown(f'<div class="kpi-card kpi-media"><div class="kpi-title">Severidad MEDIA</div><div class="kpi-number">{tot_medias}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Filtrar MEDIAS", key="btn_med", on_click=set_filtro_kpi, args=("MEDIA",), use_container_width=True)

        with k4:
            st.markdown(f'<div class="kpi-card kpi-baja"><div class="kpi-title">Severidad BAJA</div><div class="kpi-number">{tot_bajas}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Filtrar BAJAS", key="btn_baj", on_click=set_filtro_kpi, args=("BAJA",), use_container_width=True)

        with k5:
            st.markdown(f'<div class="kpi-card kpi-gate"><div class="kpi-title">Quality Gate</div><div class="kpi-number" style="font-size:16px; margin-top:8px;">FAILED 🚫</div></div>', unsafe_allow_html=True)
            st.button("🔍 Estado Gate", key="btn_gat", on_click=set_filtro_kpi, args=("Todas",), use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # APLICACIÓN DE FILTRO DE BOTÓN DE KPI
        if st.session_state['filtro_kpi'] != "Todas":
            df_kpi_vista = df_filtrado[df_filtrado["Severidad final"] == st.session_state['filtro_kpi']]
            titulo_kpi = f"📌 Hallazgos de Severidad {st.session_state['filtro_kpi']}"
        else:
            df_kpi_vista = df_filtrado.copy()
            titulo_kpi = "📌 Matriz Consolidada de Hallazgos SAST / SCA"

        # GRÁFICOS VISUALES CON PLOTLY
        st.markdown('<div class="section-header"><h3>📈 Gráficos de Distribución de Seguridad</h3></div>', unsafe_allow_html=True)
        g_col1, g_col2 = st.columns(2)

        with g_col1:
            st.subheader("Distribución por Nivel de Severidad")
            df_sev_chart = df_filtrado["Severidad final"].value_counts().reset_index()
            df_sev_chart.columns = ["Severidad", "Cantidad"]
            
            color_map = {"ALTA": "#DC2626", "MEDIA": "#F59E0B", "BAJA": "#0D9488"}
            fig_pie = px.pie(
                df_sev_chart, names="Severidad", values="Cantidad",
                color="Severidad", color_discrete_map=color_map,
                hole=0.4
            )
            fig_pie.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=290)
            st.plotly_chart(fig_pie, use_container_width=True)

        with g_col2:
            st.subheader("Hallazgos por OWASP Top 10 Categoría")
            if "OWASP Top 10 2021" in df_filtrado.columns:
                df_owasp = df_filtrado["OWASP Top 10 2021"].value_counts().reset_index()
                df_owasp.columns = ["OWASP Category", "Cantidad"]
                fig_owasp = px.bar(
                    df_owasp, x="Cantidad", y="OWASP Category", orientation='h',
                    color="Cantidad", color_continuous_scale="Reds", text="Cantidad"
                )
                fig_owasp.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=290)
                st.plotly_chart(fig_owasp, use_container_width=True)

        st.markdown("---")

        # PESTAÑAS SECUNDARIAS DE NAVEGACIÓN
        tab_sub1, tab_sub2, tab_sub3, tab_sub4, tab_sub5 = st.tabs([
            "📌 Matriz de Hallazgos", "🔍 Triaje Semgrep", "🛠️ Plan de Remediación", "💻 Detalle Técnico", "🛡️ Quality Gate & WAF"
        ])

        with tab_sub1:
            st.markdown(f"### {titulo_kpi} ({len(df_kpi_vista)} registros)")
            cols_mostrar = [c for c in ["Código", "Hallazgo", "Severidad final", modulos_col, "Archivo / endpoint", "OWASP Top 10 2021", "CWE", "Prioridad", "Plazo", "Estado final"] if c in df_kpi_vista.columns]
            st.dataframe(df_kpi_vista[cols_mostrar], use_container_width=True, hide_index=True)

        with tab_sub2:
            st.markdown("### 🔍 Triaje Detallado de Resultados Semgrep SAST")
            if not df_semgrep.empty:
                st.dataframe(df_semgrep, use_container_width=True, hide_index=True)
            else:
                st.info("No hay registros adicionales de Semgrep.")

        with tab_sub3:
            st.markdown("### 🛠️ Plan de Remediación Priorizado en Código")
            if not df_plan.empty:
                st.dataframe(df_plan, use_container_width=True, hide_index=True)
            else:
                st.info("No se encontró plan de remediación en el archivo.")

        with tab_sub4:
            st.markdown("### 💻 Detalle Técnico y Arquitectura del Sistema")
            if not df_tech.empty:
                st.dataframe(df_tech, use_container_width=True, hide_index=True)
            else:
                st.info("No hay detalle técnico disponible.")

        with tab_sub5:
            st.markdown("### 🛡️ Reglas Quality Gate CI/CD & Controles Compensatorios WAF F5")
            c_f5_1, c_f5_2 = st.columns(2)
            with c_f5_1:
                st.subheader("Reglas de Bloqueo Pipeline CI/CD")
                st.json({
                    "Quality Gate": "FAILED CONTROLADO",
                    "Regla Críticas": "0 Permitidas (0 Confirmadas)",
                    "Regla Altas": "0 Permitidas (12 Confirmadas - BLOQUEAR BUILD)",
                    "Regla Secretos": "0 Permitidos (8 Confirmados - BLOQUEAR BUILD)"
                })
            with c_f5_2:
                st.subheader("Controles Compensatorios F5 WAF")
                st.info("✔ Bloquear rutas /tareas/* y /test_mensaje/* en F5 Reverse Proxy.\n\n✔ Aplicar firmas de inspección SQLi agresivas para parámetros de búsqueda.\n\n✔ Inyectar cabeceras Content-Security-Policy (CSP) estrictas.")

# =============================================================================
# PESTAÑA 2: MÓDULO ADMINISTRADOR
# =============================================================================
if st.session_state.get('rol_actual') == "Administrador":
    with tab_admin:
        st.title("⚙️ Módulo Administrador & Auditoría de Accesos")
        st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
        st.success("🔓 Acceso de Administración Concedido")
        
        tab_sub_carga, tab_sub_logs = st.tabs(["📂 Carga y Borrado de Matrices SAST", "📋 Log de Usuarios Registrados"])
        
        with tab_sub_carga:
            col_up1, col_up2 = st.columns([2, 1])
            
            with col_up1:
                st.markdown("### 📥 Cargar Nueva Matriz de Auditoría SAST")
                fecha_subida = st.date_input("🗓️ Fecha del Escaneo / Auditoría:", obtener_hora_quito())
                archivo_nuevo = st.file_uploader("📂 Seleccionar Matriz Excel (.xlsx)", type=["xlsx"])

                if archivo_nuevo and st.button("💾 Guardar y Registrar Matriz de Auditoría", use_container_width=True):
                    try:
                        df_nuevo = pd.read_excel(archivo_nuevo, sheet_name="Matriz Hallazgos", header=2)
                        df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]

                        f_str = fecha_subida.strftime("%Y-%m-%d")
                        df_nuevo["FechaCargaMatriz"] = f_str

                        if guardar_excel_seguro(df_nuevo, ARCHIVO_EXCEL):
                            st.balloons()
                            st.success(f"✅ ¡Matriz del {f_str} guardada con éxito!")
                            time.sleep(0.5)
                            st.rerun()

                    except Exception as e:
                        st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")

            with col_up2:
                st.markdown("### 📊 Estado del Archivo Acumulado")
                st.metric("Total Hallazgos Registrados", f"{len(df_hallazgos)}")

            st.markdown("---")

        with tab_sub_logs:
            st.markdown("### 📋 Historial de Ingresos de Usuarios")
            st.caption("A continuación se muestra el registro acumulado con fecha y hora de Quito de cada usuario que ha iniciado sesión.")
            
            if os.path.exists(ARCHIVO_LOGS) and os.path.getsize(ARCHIVO_LOGS) > 0:
                try:
                    df_logs_read = pd.read_csv(ARCHIVO_LOGS)
                    st.dataframe(df_logs_read.sort_index(ascending=False), use_container_width=True, hide_index=True)
                except Exception as e:
                    st.error(f"Error al leer el archivo de logs: {e}")
            else:
                st.info("Aún no se registran inicios de sesión.")