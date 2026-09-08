import os
import io
import time
import socket
from datetime import datetime, timezone, timedelta
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE ZONA HORARIA (QUITO, ECUADOR)
# -----------------------------------------------------------------------------
ZONA_HORARIA_QUITO = timezone(timedelta(hours=-5))

def obtener_hora_quito():
    """Obtiene la fecha y hora actual exacta en Quito, Ecuador (UTC-5)."""
    return datetime.now(ZONA_HORARIA_QUITO)

# Configuración de la página Streamlit
st.set_page_config(
    page_title="Dashboard de Gestión & Operaciones de Agencias",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
    <style>
    .stApp { background-color: #F1F5F9; color: #0F172A; }
    .kpi-card { border-radius: 12px; padding: 16px 12px; text-align: center; color: white; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08); }
    .kpi-asig { background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); }
    .kpi-vis { background: linear-gradient(135deg, #16A34A 0%, #15803D 100%); }
    .kpi-novis { background: linear-gradient(135deg, #DC2626 0%, #991B1B 100%); }
    .kpi-cob { background: linear-gradient(135deg, #0D9488 0%, #115E59 100%); }
    .kpi-title { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px; opacity: 0.95; }
    .kpi-number { font-size: 34px; font-weight: 800; margin-top: 2px; }
    .section-header { background-color: #FFFFFF; padding: 12px 20px; border-radius: 10px; border-left: 5px solid #0284C7; box-shadow: 0 2px 4px rgba(0,0,0,0.04); margin-bottom: 15px; }
    h1, h2, h3 { color: #0284C7 !important; font-weight: 800 !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    [data-testid="stSidebar"] { background-color: #0F172A; }
    [data-testid="stSidebar"] * { color: #F8FAFC !important; }
    .author-credit { font-size: 11px; color: #64748B; font-style: italic; margin-top: -10px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_EXCEL = "base_visitas.xlsx"
ARCHIVO_LOGS = "registro_ingresos.csv"

# -----------------------------------------------------------------------------
# FUNCIONES DE CAPTURA DE RED Y LOGS AVANZADOS
# -----------------------------------------------------------------------------
def obtener_datos_cliente():
    """Obtiene la dirección IP y el hostname del cliente o servidor."""
    ip_cliente = "127.0.0.1"
    try:
        headers = st.context.headers
        if headers:
            if "X-Forwarded-For" in headers:
                ip_cliente = headers["X-Forwarded-For"].split(",")[0].strip()
            elif "Remote-Addr" in headers:
                ip_cliente = headers["Remote-Addr"].strip()
    except Exception:
        pass

    try:
        hostname = socket.gethostname()
    except Exception:
        hostname = "Desconocido"

    return ip_cliente, hostname

def registrar_ingreso(usuario, rol):
    """Registra inicios de sesión progresivamente con Hora, Usuario, Rol, IP y Hostname."""
    fecha_hora_quito = obtener_hora_quito().strftime("%Y-%m-%d %H:%M:%S")
    ip_cliente, hostname = obtener_datos_cliente()
    
    nuevo_log = pd.DataFrame([{
        "FechaHora": fecha_hora_quito,
        "Usuario": usuario,
        "Rol": rol,
        "IP_Cliente": ip_cliente,
        "Hostname": hostname
    }])
    
    try:
        header_needed = not os.path.exists(ARCHIVO_LOGS) or os.path.getsize(ARCHIVO_LOGS) == 0
        nuevo_log.to_csv(ARCHIVO_LOGS, mode='a', header=header_needed, index=False)
    except Exception as e:
        st.error(f"Error al registrar log de usuario: {e}")

# -----------------------------------------------------------------------------
# FUNCIONES DE ESTILOS PARA TABLAS EN STREAMLIT
# -----------------------------------------------------------------------------
def colorear_visitada(val):
    """Aplica color de fondo a la celda según si la visita fue Sí o No."""
    if str(val).strip() == "Sí":
        return 'background-color: #DCFCE7; color: #15803D; font-weight: bold; text-align: center;'
    elif str(val).strip() == "No":
        return 'background-color: #FEE2E2; color: #B91C1C; font-weight: bold; text-align: center;'
    return ''

def estilo_tabla_ejecutiva(df, col_status="Visitada"):
    """Formatea la tabla con encabezados azules en negrita y resalta la columna Visitada."""
    styler = df.style.set_table_styles([
        {
            'selector': 'th',
            'props': [
                ('background-color', '#0284C7'),
                ('color', '#FFFFFF'),
                ('font-weight', 'bold'),
                ('text-align', 'center'),
                ('font-size', '14px')
            ]
        },
        {
            'selector': 'td',
            'props': [('font-size', '13px')]
        }
    ])
    if col_status in df.columns:
        styler = styler.map(colorear_visitada, subset=[col_status])
    return styler

# -----------------------------------------------------------------------------
# WIDGET DE RELOJ DINÁMICO EN SIDEBAR
# -----------------------------------------------------------------------------
reloj_html = """
<div style="background: rgba(255, 255, 255, 0.08); border: 1px solid rgba(255, 255, 255, 0.15); border-radius: 10px; padding: 10px 12px; text-align: center; font-family: 'Segoe UI', Roboto, sans-serif;">
    <div style="font-size: 12px; font-weight: 700; color: #38BDF8; text-transform: uppercase; letter-spacing: 0.5px;">📍 QUITO, ECUADOR ⛅</div>
    <div id="reloj-quito" style="font-size: 22px; font-weight: 800; color: #F8FAFC; margin: 4px 0;">⏰ 00:00:00</div>
    <div id="fecha-quito" style="font-size: 11px; color: #94A3B8;">📅 --/--/----</div>
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
    document.getElementById('fecha-quito').innerHTML = '📅 ' + dia + '/' + mes + '/' + anio;
}
actualizarRelojQuito();
setInterval(actualizarRelojQuito, 1000);
</script>
"""

with st.sidebar:
    components.html(reloj_html, height=115)

# USUARIOS AUTORIZADOS
USUARIOS_AUTORIZADOS = {
    "gabriela.armendariz": {"clave": "GArmendariz2026!", "rol": "Administrador", "nombre": "Gabriela Armendariz"},
    "usuario.generico": {"clave": "Lectura2026!", "rol": "Lectura", "nombre": "Usuario Consulta (Solo Lectura)"},
    "admin": {"clave": "admin123", "rol": "Administrador", "nombre": "Administrador Principal"}
}

def guardar_excel_seguro(df: pd.DataFrame, filepath: str = ARCHIVO_EXCEL) -> bool:
    temp_filepath = f"{os.path.splitext(filepath)[0]}_temp.xlsx"
    try:
        df.to_excel(temp_filepath, index=False, engine="openpyxl")
        if os.path.exists(temp_filepath):
            os.replace(temp_filepath, filepath)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error al guardar el archivo en disco: {e}")
        if os.path.exists(temp_filepath):
            try:
                os.remove(temp_filepath)
            except Exception:
                pass
        return False

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

# -----------------------------------------------------------------------------
# GENERACIÓN DE INFORME DETALLADO EN WORD
# -----------------------------------------------------------------------------
def generar_reporte_word_detallado(df_data, col_agencia, col_region, col_coord, col_fecha, cols_eval, region_filtro, coord_filtro, fecha_matriz):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    COLOR_AZUL = RGBColor(2, 132, 199)
    COLOR_TITULO = RGBColor(15, 23, 42)
    HEX_HEADER = "0284C7"
    HEX_ALT = "F8FAFC"
    HEX_SI = "DCFCE7"
    HEX_NO = "FEE2E2"

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("INFORME EJECUTIVO Y AUDITORÍA DE OPERACIONES DE AGENCIAS\nCONTROL DE GESTIÓN Y OPERACIÓN DE REDES")
    r_title.font.name = 'Segoe UI'
    r_title.font.size = Pt(15)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_AZUL

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sub = p_sub.add_run(f"Fecha de emisión: {obtener_hora_quito().strftime('%d/%m/%Y %H:%M')}\nElaborado por: Gabriela Armendariz\nFiltros aplicados: Región: {region_filtro} | Zona/Coordinador: {coord_filtro} | Matriz Activa: {fecha_matriz}")
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True

    doc.add_paragraph()

    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Resumen Ejecutivo de Cobertura Operativa")
    r_h1.font.name = 'Segoe UI'
    r_h1.font.color.rgb = COLOR_AZUL
    r_h1.font.bold = True

    tot_registros = len(df_data) if not df_data.empty else 0
    tot_vis = (df_data["Visitada"] == "Sí").sum() if not df_data.empty else 0
    tot_novis = tot_registros - tot_vis
    pct_cob = (tot_vis / tot_registros * 100) if tot_registros > 0 else 0.0
    pct_pen = (tot_novis / tot_registros * 100) if tot_registros > 0 else 0.0

    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body1.add_run(
        f"El presente informe consolida la auditoría operativa para el segmento seleccionado "
        f"(Región: '{region_filtro}', Zona: '{coord_filtro}'). En la matriz de trabajo ('{fecha_matriz}'), "
        f"se registra un universo total de {tot_registros:,} operaciones transaccionales y asignaciones.\n\n"
        f"De este total auditado, {tot_vis:,} corresponden a supervisiones efectivas realizadas en campo ({pct_cob:.1f}% de cobertura), "
        f"mientras que {tot_novis:,} inspecciones se encuentran pendientes ({pct_pen:.1f}%)."
    )

    table_kpi = doc.add_table(rows=1, cols=4)
    table_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells = table_kpi.rows[0].cells
    headers_kpi = ["Indicador KPI", "Valor Registrado", "% del Total", "Estado / Diagnóstico"]
    for i, title in enumerate(headers_kpi):
        set_cell_background(hdr_cells[i], HEX_HEADER)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    diag_cob = "🟢 Cobertura Excelente" if pct_cob >= 80 else ("🟡 Cobertura Aceptable" if pct_cob >= 65 else "🔴 Alerta Rezagados")

    datos_kpi = [
        ["Registros Totales", f"{tot_registros:,}", "100.0%", "Universo total de control programado"],
        ["Visitas Realizadas", f"{tot_vis:,}", f"{pct_cob:.1f}%", "Supervisiones ejecutadas en campo"],
        ["Visitas Pendientes", f"{tot_novis:,}", f"{pct_pen:.1f}%", "Brecha operativa por completar"],
        ["% Cobertura Real", f"{pct_cob:.1f}%", "—", diag_cob]
    ]

    for idx, row in enumerate(datos_kpi):
        row_cells = table_kpi.add_row().cells
        bg = HEX_ALT if idx % 2 == 0 else "FFFFFF"
        for col_idx, val in enumerate(row):
            set_cell_background(row_cells[col_idx], bg)
            p = row_cells[col_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx in [0, 3] else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            if col_idx == 1:
                r.font.bold = True

    doc.add_paragraph()

    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Desglose Operativo por Coordinador y Zona de Supervisión")
    r_h2.font.name = 'Segoe UI'
    r_h2.font.color.rgb = COLOR_AZUL
    r_h2.font.bold = True

    p_body2 = doc.add_paragraph()
    p_body2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body2.add_run("Desglose detallado por coordinador dentro de la selección actual:")

    if not df_data.empty and col_coord in df_data.columns:
        resumen_coord = df_data.groupby(col_coord).agg(
            RegistrosTotales=(col_agencia, "count"),
            VisitasRealizadas=("Visitada", lambda x: (x == "Sí").sum()),
            VisitasPendientes=("Visitada", lambda x: (x == "No").sum())
        ).reset_index()

        resumen_coord["Pct"] = (resumen_coord["VisitasRealizadas"] / resumen_coord["RegistrosTotales"] * 100)
        resumen_coord = resumen_coord.sort_values(by="Pct", ascending=False)

        table_coord = doc.add_table(rows=1, cols=5)
        table_coord.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_c = table_coord.rows[0].cells
        cols_t = ["Coordinador / Zona", "Registros Totales", "Visitadas", "Pendientes", "% Cumplimiento"]
        for i, title in enumerate(cols_t):
            set_cell_background(hdr_c[i], HEX_HEADER)
            p = hdr_c[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in resumen_coord.iterrows():
            row_cells = table_coord.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            pct_val = row["Pct"]
            
            vals = [str(row[col_coord]), f"{row['RegistrosTotales']:,}", f"{row['VisitasRealizadas']:,}", f"{row['VisitasPendientes']:,}", f"{pct_val:.1f}%"]
            for col_idx, val in enumerate(vals):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(val)

    doc.add_paragraph()

    h3_m = doc.add_heading(level=1)
    r_h3m = h3_m.add_run("3. Matriz de Detalle con Estado de Visita (Sí / No)")
    r_h3m.font.name = 'Segoe UI'
    r_h3m.font.color.rgb = COLOR_AZUL
    r_h3m.font.bold = True

    p_det = doc.add_paragraph()
    p_det.add_run("A continuación se muestra el extracto de registros evaluados en la muestra activa:")

    cols_reg_show = [c for c in [col_region, col_coord, col_agencia, "Visitada", col_fecha] if c in df_data.columns]
    df_sample = df_data[cols_reg_show].head(15)

    table_mat = doc.add_table(rows=1, cols=len(cols_reg_show))
    table_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_m = table_mat.rows[0].cells
    for i, c_name in enumerate(cols_reg_show):
        set_cell_background(hdr_m[i], HEX_HEADER)
        p = hdr_m[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(c_name)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, row in df_sample.iterrows():
        row_cells = table_mat.add_row().cells
        for col_idx, c_name in enumerate(cols_reg_show):
            val_str = str(row[c_name]) if pd.notnull(row[c_name]) else "None"
            bg = HEX_ALT
            
            if c_name == "Visitada":
                if val_str == "Sí":
                    bg = HEX_SI
                elif val_str == "No":
                    bg = HEX_NO
            elif row_idx % 2 == 1:
                bg = "FFFFFF"

            set_cell_background(row_cells[col_idx], bg)
            p = row_cells[col_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_name in ["Visitada", col_fecha] else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val_str)
            if c_name == "Visitada":
                r.font.bold = True
                r.font.color.rgb = RGBColor(21, 128, 61) if val_str == "Sí" else RGBColor(185, 28, 28)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# -----------------------------------------------------------------------------
# CONTROL DE SESIÓN Y LOGIN
# -----------------------------------------------------------------------------
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False
if 'usuario_actual' not in st.session_state:
    st.session_state['usuario_actual'] = ""

if not st.session_state['autenticado']:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown('<h1>🔒 Control de Acceso</h1>', unsafe_allow_html=True)
        st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
        st.caption("Ingrese sus credenciales autorizadas para consultar el Dashboard de Agencias.")
        
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
# LECTURA DE DATOS
# -----------------------------------------------------------------------------
@st.cache_data(ttl=120)
def cargar_datos():
    path = ARCHIVO_EXCEL if os.path.exists(ARCHIVO_EXCEL) else ("Base_Visitas.xlsx" if os.path.exists("Base_Visitas.xlsx") else None)
    if not path or os.path.getsize(path) == 0:
        return pd.DataFrame(), "NombreAgencia", "Región", "Zona", "FechaVisita", []
    
    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception:
        return pd.DataFrame(), "NombreAgencia", "Región", "Zona", "FechaVisita", []

    if not df.empty:
        df.columns = [str(c).strip() for c in df.columns]

        if "FechaCargaMatriz" not in df.columns:
            df["FechaCargaMatriz"] = obtener_hora_quito().strftime("%Y-%m-%d")

        col_region = "Región" if "Región" in df.columns else ("Region" if "Region" in df.columns else df.columns[0])
        col_coord = "Zona" if "Zona" in df.columns else ("Coordinador" if "Coordinador" in df.columns else df.columns[1])
        col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else ("Agencia" if "Agencia" in df.columns else df.columns[3])
        col_fecha = "FechaVisita" if "FechaVisita" in df.columns else ("Fecha de Visita" if "Fecha de Visita" in df.columns else "Fecha")

        if "Visitada" not in df.columns:
            if col_fecha in df.columns:
                df["Fecha_DT"] = pd.to_datetime(df[col_fecha], errors='coerce')
                df["Visitada"] = df["Fecha_DT"].apply(lambda x: "Sí" if pd.notnull(x) else "No")
            else:
                df["Visitada"] = "No"

        cols_eval = [c for c in ["Calidad", "ProcesosTransaccionales", "ManejoEfectivoControl", "TalentoCultura"] if c in df.columns]
        
        if cols_eval:
            df["PromedioGral"] = df[cols_eval].apply(pd.to_numeric, errors='coerce').mean(axis=1)
        else:
            df["PromedioGral"] = 0.0

        return df, col_agencia, col_region, col_coord, col_fecha, cols_eval
    return df, "NombreAgencia", "Región", "Zona", "FechaVisita", []

df_base, col_agencia, col_region, col_coord, col_fecha, cols_eval = cargar_datos()

if 'region_activa' not in st.session_state: st.session_state['region_activa'] = "Todas"
if 'coordinador_activo' not in st.session_state: st.session_state['coordinador_activo'] = "Todos"
if 'filtro_kpi' not in st.session_state: st.session_state['filtro_kpi'] = "Todas"

def set_filtro_kpi(valor):
    st.session_state['filtro_kpi'] = valor

# -----------------------------------------------------------------------------
# SIDEBAR NAVEGACIÓN
# -----------------------------------------------------------------------------
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.get('nombre_actual', 'Invitado')}")
st.sidebar.caption(f"Rol: {st.session_state.get('rol_actual', 'General')}")
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state['autenticado'] = False
    st.rerun()

st.sidebar.markdown('<p class="author-credit" style="color:#94A3B8;">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
st.sidebar.markdown("---")

if st.session_state.get('rol_actual') == "Administrador":
    tab_dash, tab_admin = st.tabs(["📊 Dashboard de Gestión & Operaciones", "⚙️ Módulo Administrador & Logs"])
else:
    tab_dash, = st.tabs(["📊 Dashboard de Gestión & Operaciones"])

# =============================================================================
# PESTAÑA 1: DASHBOARD
# =============================================================================
with tab_dash:
    if df_base.empty:
        st.warning("⚠️ No se ha encontrado ninguna matriz cargada. Comuníquese con un Administrador para cargar la información.")
    else:
        st.sidebar.title("📌 Control de Gestión")

        fechas_cargas = sorted([str(x) for x in df_base["FechaCargaMatriz"].unique() if pd.notnull(x)], reverse=True)
        opciones_matriz = ["Última Carga (Actual)", "🌐 Ver Todas las Matrices Acumuladas"] + fechas_cargas
        matriz_sel = st.sidebar.selectbox("📂 Seleccionar Matriz por Fecha de Carga:", opciones_matriz)

        if matriz_sel == "Última Carga (Actual)":
            df_matriz_activa = df_base[df_base["FechaCargaMatriz"] == fechas_cargas[0]] if fechas_cargas else df_base.copy()
            fecha_display = fechas_cargas[0] if fechas_cargas else "Hoy"
        elif matriz_sel == "🌐 Ver Todas las Matrices Acumuladas":
            df_matriz_activa = df_base.copy()
            fecha_display = "Consolidado Histórico Completo"
        else:
            df_matriz_activa = df_base[df_base["FechaCargaMatriz"] == matriz_sel]
            fecha_display = matriz_sel

        regiones = ["Todas"] + sorted([str(x) for x in df_matriz_activa[col_region].dropna().unique() if str(x).strip() != ""])
        idx_reg = regiones.index(st.session_state['region_activa']) if st.session_state['region_activa'] in regiones else 0
        region_sel = st.sidebar.selectbox("Filtrar por Región:", regiones, index=idx_reg)

        if region_sel != st.session_state['region_activa']:
            st.session_state['region_activa'] = region_sel
            st.session_state['coordinador_activo'] = "Todos"
            st.session_state['filtro_kpi'] = "Todas"
            st.rerun()

        df_region = df_matriz_activa[df_matriz_activa[col_region].astype(str) == st.session_state['region_activa']] if st.session_state['region_activa'] != "Todas" else df_matriz_activa.copy()

        coordinadores = ["Todos"] + sorted([str(x) for x in df_region[col_coord].dropna().unique() if str(x).strip() != ""])
        idx_coord = coordinadores.index(st.session_state['coordinador_activo']) if st.session_state['coordinador_activo'] in coordinadores else 0
        coord_sel = st.sidebar.selectbox("Filtrar por Coordinador/Zona:", coordinadores, index=idx_coord)

        if coord_sel != st.session_state['coordinador_activo']:
            st.session_state['coordinador_activo'] = coord_sel
            st.session_state['filtro_kpi'] = "Todas"
            st.rerun()

        if st.sidebar.button("🔄 Restablecer Filtros"):
            st.session_state['region_activa'] = "Todas"
            st.session_state['coordinador_activo'] = "Todos"
            st.session_state['filtro_kpi'] = "Todas"
            st.rerun()

        st.sidebar.markdown("---")
        st.sidebar.subheader("📄 Reporte Oficial")

        df_filtrado = df_matriz_activa.copy()
        if st.session_state['region_activa'] != "Todas":
            df_filtrado = df_filtrado[df_filtrado[col_region].astype(str) == st.session_state['region_activa']]
        if st.session_state['coordinador_activo'] != "Todos":
            df_filtrado = df_filtrado[df_filtrado[col_coord].astype(str) == st.session_state['coordinador_activo']]

        # REPORTE WORD
        doc_bytes = generar_reporte_word_detallado(
            df_filtrado, col_agencia, col_region, col_coord, col_fecha, cols_eval,
            st.session_state['region_activa'], st.session_state['coordinador_activo'], fecha_display
        )
        st.sidebar.download_button(
            label="📄 Descargar Informe Detallado (Word)",
            data=doc_bytes,
            file_name=f"Informe_Detallado_{st.session_state['region_activa']}_{obtener_hora_quito().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

        tot_registros = len(df_filtrado)
        tot_vis_registros = (df_filtrado["Visitada"] == "Sí").sum()
        tot_novis_registros = (df_filtrado["Visitada"] == "No").sum()
        pct_cob_real = (tot_vis_registros / tot_registros * 100) if tot_registros > 0 else 0.0

        st.title("📊 Control de Gestión de Agencias")
        st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
        st.caption(f"📌 **Matriz Activa:** {fecha_display} | **Total Filas Evaluadas:** {tot_registros:,} (Filtros: Región='{st.session_state['region_activa']}', Zona='{st.session_state['coordinador_activo']}')")

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f'<div class="kpi-card kpi-asig"><div class="kpi-title">Registros Totales</div><div class="kpi-number">{tot_registros:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Ver Todos", key="btn_asig", on_click=set_filtro_kpi, args=("Todas",), use_container_width=True)

        with k2:
            st.markdown(f'<div class="kpi-card kpi-vis"><div class="kpi-title">Visitas Realizadas</div><div class="kpi-number">{tot_vis_registros:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Filtrar Realizadas", key="btn_vis", on_click=set_filtro_kpi, args=("Visitadas",), use_container_width=True)

        with k3:
            st.markdown(f'<div class="kpi-card kpi-novis"><div class="kpi-title">Visitas Pendientes</div><div class="kpi-number">{tot_novis_registros:,}</div></div>', unsafe_allow_html=True)
            st.button("🔍 Filtrar Pendientes", key="btn_novis", on_click=set_filtro_kpi, args=("No Visitadas",), use_container_width=True)

        with k4:
            st.markdown(f'<div class="kpi-card kpi-cob"><div class="kpi-title">% Cobertura Real</div><div class="kpi-number">{pct_cob_real:.1f}%</div></div>', unsafe_allow_html=True)
            st.button("🔍 Ver Estado General", key="btn_cob", on_click=set_filtro_kpi, args=("Todas",), use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if st.session_state['filtro_kpi'] == "Visitadas":
            df_kpi_vista = df_filtrado[df_filtrado["Visitada"] == "Sí"]
            titulo_kpi = "📌 Detalle de Registros con Visitas Realizadas"
        elif st.session_state['filtro_kpi'] == "No Visitadas":
            df_kpi_vista = df_filtrado[df_filtrado["Visitada"] == "No"]
            titulo_kpi = "📌 Detalle de Registros con Visitas Pendientes"
        else:
            df_kpi_vista = df_filtrado.copy()
            titulo_kpi = "📌 Detalle General de Registros de la Matriz"

        st.markdown(f"### {titulo_kpi} ({len(df_kpi_vista):,} registros)")
        cols_mostrar = [c for c in [col_region, col_coord, col_agencia, "Visitada", col_fecha, "PromedioGral"] if c in df_kpi_vista.columns]
        
        df_styled = estilo_tabla_ejecutiva(df_kpi_vista[cols_mostrar])
        st.dataframe(df_styled, use_container_width=True, hide_index=True)

        st.markdown("---")

        col_izq, col_der = st.columns([1.1, 1])

        with col_izq:
            st.subheader("Resumen por Coordinador / Zona")
            resumen_coord = df_filtrado.groupby(col_coord).agg(
                RegistrosTotales=(col_agencia, "count"),
                VisitasRealizadas=("Visitada", lambda x: (x == "Sí").sum()),
                VisitasPendientes=("Visitada", lambda x: (x == "No").sum())
            ).reset_index()

            st.dataframe(estilo_tabla_ejecutiva(resumen_coord, col_status=""), use_container_width=True, hide_index=True)

        with col_der:
            v_col, nv_col = st.columns(2)
            with v_col:
                st.subheader("Visitas Realizadas")
                df_v = df_filtrado[df_filtrado["Visitada"] == "Sí"][[col_region, col_agencia]].drop_duplicates()
                st.dataframe(estilo_tabla_ejecutiva(df_v, col_status=""), use_container_width=True, hide_index=True, height=220)
            with nv_col:
                st.subheader("Visitas Pendientes")
                df_nv = df_filtrado[df_filtrado["Visitada"] == "No"][[col_region, col_agencia]].drop_duplicates()
                st.dataframe(estilo_tabla_ejecutiva(df_nv, col_status=""), use_container_width=True, hide_index=True, height=220)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-header"><h3>📈 Dashboard de Operaciones y Transacciones</h3></div>', unsafe_allow_html=True)

        if "TipoTran" in df_filtrado.columns:
            c_tr1, c_tr2, c_tr3 = st.columns(3)
            total_trans = len(df_filtrado)
            terminales_unicas = df_filtrado["Terminal"].nunique() if "Terminal" in df_filtrado.columns else 0
            tipos_unicos = df_filtrado["TipoTran"].nunique()

            c_tr1.metric("Total Transacciones Registradas", f"{total_trans:,}")
            c_tr2.metric("Terminales / Cajas Activas", f"{terminales_unicas}")
            c_tr3.metric("Tipos de Transacciones", f"{tipos_unicos}")

            g_col1, g_col2 = st.columns(2)

            with g_col1:
                st.subheader("Volumen por Tipo de Transacción")
                df_tipo = df_filtrado["TipoTran"].value_counts().reset_index()
                df_tipo.columns = ["TipoTran", "Cantidad"]
                fig_tipo = px.bar(
                    df_tipo, x="Cantidad", y="TipoTran", orientation='h',
                    color="Cantidad", color_continuous_scale="Blues", text="Cantidad"
                )
                fig_tipo.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=280)
                st.plotly_chart(fig_tipo, use_container_width=True)

            with g_col2:
                st.subheader("Distribución por Hora de Operación")
                if "Hora" in df_filtrado.columns:
                    df_hora = df_filtrado["Hora"].value_counts().reset_index()
                    df_hora.columns = ["Hora", "Transacciones"]
                    df_hora = df_hora.sort_values("Hora")
                    fig_hora = px.line(df_hora, x="Hora", y="Transacciones", markers=True, line_shape="spline")
                    fig_hora.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=280)
                    st.plotly_chart(fig_hora, use_container_width=True)

            st.subheader(f"🔍 Matriz Transaccional Completa ({total_trans:,} registros evaluados)")
            cols_trans_mostrar = [c for c in [col_region, col_coord, "Codigo", "Código", col_agencia, "Anio", "Mes", "Dia", "Semana", "Hora", "Terminal", "NombreMaq", "TipoTran"] if c in df_filtrado.columns]
            st.dataframe(estilo_tabla_ejecutiva(df_filtrado[cols_trans_mostrar], col_status=""), use_container_width=True, hide_index=True)
        else:
            st.info("Cargue una matriz con registros transaccionales para ver esta sección en detalle.")

# =============================================================================
# PESTAÑA 2: MÓDULO ADMINISTRADOR
# =============================================================================
if st.session_state.get('rol_actual') == "Administrador":
    with tab_admin:
        st.title("⚙️ Módulo Administrador & Auditoría de Accesos")
        st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
        st.success("🔓 Acceso de Administración Concedido")
        
        tab_sub_carga, tab_sub_logs = st.tabs(["📂 Carga y Borrado de Matrices", "📋 Log de Usuarios Registrados"])
        
        with tab_sub_carga:
            col_up1, col_up2 = st.columns([2, 1])
            
            with col_up1:
                st.markdown("### 📥 Cargar Nueva Matriz Diaria")
                
                fecha_actual_quito = obtener_hora_quito().date()
                fecha_subida = st.date_input(
                    "🗓️ Fecha de Carga de la Matriz:",
                    value=fecha_actual_quito,
                    min_value=fecha_actual_quito
                )
                
                archivo_nuevo = st.file_uploader("📂 Seleccionar Matriz Excel (.xlsx)", type=["xlsx"])

                if archivo_nuevo and st.button("💾 Guardar y Registrar Matriz de la Fecha", use_container_width=True):
                    if fecha_subida < fecha_actual_quito:
                        st.error("❌ No se permite registrar matrices con fechas pasadas. Seleccione la fecha de hoy o posterior.")
                    else:
                        try:
                            df_nuevo = pd.read_excel(archivo_nuevo)
                            df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]

                            f_str = fecha_subida.strftime("%Y-%m-%d")
                            df_nuevo["FechaCargaMatriz"] = f_str

                            if not df_base.empty and "FechaCargaMatriz" in df_base.columns:
                                df_base_sin_fecha = df_base[df_base["FechaCargaMatriz"] != f_str]
                                df_consolidado = pd.concat([df_base_sin_fecha, df_nuevo], ignore_index=True)
                            else:
                                df_consolidado = df_nuevo

                            if guardar_excel_seguro(df_consolidado, ARCHIVO_EXCEL):
                                st.balloons()
                                st.success(f"✅ ¡Matriz del {f_str} guardada con éxito!")
                                time.sleep(0.5)
                                st.rerun()

                        except Exception as e:
                            st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")

            with col_up2:
                st.markdown("### 📊 Estado del Histórico Acumulado")
                st.metric("Total Registros Acumulados", f"{len(df_base):,}")
                fechas_existentes = sorted([str(x) for x in df_base["FechaCargaMatriz"].unique() if pd.notnull(x)], reverse=True) if not df_base.empty and "FechaCargaMatriz" in df_base.columns else []
                st.metric("Fechas Registradas", f"{len(fechas_existentes)}")

            st.markdown("---")

            st.markdown("### 🗑️ Eliminar Información de una Fecha Específica")
            if fechas_existentes:
                fecha_a_eliminar = st.selectbox("Elija la fecha de carga que desea borrar:", fechas_existentes)
                if st.button(f"🚨 Eliminar Definitivamente Registros del {fecha_a_eliminar}"):
                    df_limpio = df_base[df_base["FechaCargaMatriz"] != fecha_a_eliminar]
                    if guardar_excel_seguro(df_limpio, ARCHIVO_EXCEL):
                        st.success(f"✅ Se han eliminado todos los registros del día {fecha_a_eliminar}.")
                        time.sleep(0.5)
                        st.rerun()

        with tab_sub_logs:
            st.markdown("### 📋 Historial de Ingresos de Usuarios")
            st.caption("A continuación se muestra el registro acumulado con fecha y hora de cada usuario que ha iniciado sesión.")
            
            if os.path.exists(ARCHIVO_LOGS) and os.path.getsize(ARCHIVO_LOGS) > 0:
                try:
                    df_logs_read = pd.read_csv(ARCHIVO_LOGS)
                    st.dataframe(estilo_tabla_ejecutiva(df_logs_read.sort_index(ascending=False), col_status=""), use_container_width=True, hide_index=True)
                except Exception as e:
                    st.error(f"Error al leer el archivo de logs: {e}")
            else:
                st.info("Aún no se registran inicios de sesión.")