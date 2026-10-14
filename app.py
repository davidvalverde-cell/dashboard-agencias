import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
import time
import asyncio
import aiohttp
import pycountry
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.chart import BarChart, Reference
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
    page_title="Dashboard de Gestión & Operaciones de Agencias",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS
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

ARCHIVO_EXCEL = "base_visitas.xlsx"
ARCHIVO_LOGS = "registro_ingresos.csv"

# --- CONFIGURACIÓN Y CONSTANTES PARA EL MÓDULO DE ESCANEO DE IPS ---
IPINFO_API_TOKEN = "1610cc1e3d24a1"
VIRUSTOTAL_API_KEY = "2087ece6d1897627c3f204176af5cd0e413a7e4e111cff06b992308a90f5db50"
ABUSEIPDB_API_KEY = "ffafa47c3907b66766d7a3af9446530de589e6ea7d525780691b867962aef24c64b68020b358bf33"

RETRY_ATTEMPTS = 3
RETRY_DELAY = 2

CATEGORIAS_ABUSEIPDB = {
    3: "Phishing", 4: "Botnet", 5: "Malware", 6: "Spam", 7: "DDoS", 8: "Fuerza Bruta",
    9: "Escaneo de Puertos", 10: "Explotación Web", 11: "Explotación SSH",
    14: "VPN Abusiva", 15: "Explotación FTP", 18: "Comprometida"
}

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
    """Guarda en un archivo temporal .xlsx con el motor openpyxl para evitar errores de extensión."""
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
            try:
                os.remove(temp_filepath)
            except Exception:
                pass
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
# FUNCIONES DEL MÓDULO DE ESCANEO DE IPS
# -----------------------------------------------------------------------------
def obtener_nombre_pais_completo(codigo):
    try:
        return pycountry.countries.get(alpha_2=codigo.upper()).name.upper()
    except Exception:
        return "DESCONOCIDO"

async def obtener_info_geografica(session, ip):
    url = f"https://ipinfo.io/{ip}?token={IPINFO_API_TOKEN}"
    for _ in range(RETRY_ATTEMPTS):
        try:
            async with session.get(url, timeout=15) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return {
                        "pais": obtener_nombre_pais_completo(data.get("country", "")),
                        "region": data.get("region", "Desconocido"),
                        "ciudad": data.get("city", "Desconocido")
                    }
        except Exception:
            await asyncio.sleep(RETRY_DELAY)
    return {"pais": "DESCONOCIDO", "region": "", "ciudad": ""}

async def obtener_riesgo_virustotal(session, ip):
    url = f"https://www.virustotal.com/api/v3/ip_addresses/{ip}"
    headers = {"x-apikey": VIRUSTOTAL_API_KEY}
    for _ in range(RETRY_ATTEMPTS):
        try:
            async with session.get(url, headers=headers, timeout=30) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    resultados = data.get('data', {}).get('attributes', {}).get('last_analysis_results', {})
                    detectados = [engine for engine, r in resultados.items() if r.get("category") in ("malicious", "suspicious")]
                    return f"Detectado por: {', '.join(detectados)}" if detectados else "Sin amenazas"
        except Exception:
            await asyncio.sleep(RETRY_DELAY)
    return "Error al consultar VT"

async def obtener_info_abuseipdb(session, ip):
    url = "https://api.abuseipdb.com/api/v2/check"
    headers = {'Key': ABUSEIPDB_API_KEY, 'Accept': 'application/json'}
    params = {'ipAddress': ip, 'maxAgeInDays': 90}
    for _ in range(RETRY_ATTEMPTS):
        try:
            async with session.get(url, headers=headers, params=params, timeout=30) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    info = data.get('data', {})
                    score = info.get('abuseConfidenceScore', 0)
                    reports = info.get('reports', [])
                    timestamps = set()
                    categorias = set()
                    comportamientos = set()
                    for r in reports:
                        if r.get("reportedAt"):
                            timestamps.add(r["reportedAt"])
                        for cat in r.get("categories", []):
                            categorias.add(cat)
                        comentario = r.get("comment", "").lower()
                        if 'brute force' in comentario: comportamientos.add("Fuerza Bruta")
                        if 'port scan' in comentario: comportamientos.add("Escaneo de Puertos")
                        if 'ddos' in comentario: comportamientos.add("DDoS")
                        if 'spam' in comentario: comportamientos.add("Spam")
                        if 'botnet' in comentario: comportamientos.add("Botnet")
                        if 'phishing' in comentario: comportamientos.add("Phishing")
                    nombres_categorias = [CATEGORIAS_ABUSEIPDB.get(c, f"ID{c}") for c in categorias]
                    return score, ", ".join(sorted(comportamientos)) or "No especificado", \
                           "; ".join(sorted(timestamps)) or "No disponible", \
                           ", ".join(sorted(nombres_categorias)) or "No especificado"
        except Exception:
            await asyncio.sleep(RETRY_DELAY)
    return 0, "No especificado", "No disponible", "No especificado"

async def analizar_ip(session, ip, progress_bar, status_text, current, total):
    geo = await obtener_info_geografica(session, ip)
    vt = await obtener_riesgo_virustotal(session, ip)
    abuse_score, comportamiento, timestamps, categorias = await obtener_info_abuseipdb(session, ip)
    
    status_text.text(f"Procesando IP {current}/{total}: {ip} (País: {geo['pais']}, Riesgo: {abuse_score}%)")
    progress_bar.progress(current / total)
    
    return {
        "ip": ip,
        "pais": geo["pais"],
        "region": geo["region"],
        "ciudad": geo["ciudad"],
        "riesgo_abuse": abuse_score,
        "comportamiento": comportamiento,
        "timestamps": timestamps,
        "categorias": categorias,
        "virustotal": vt
    }

async def analizar_ips_async(ips_list, progress_bar, status_text):
    async with aiohttp.ClientSession() as session:
        resultados = []
        total = len(ips_list)
        for i, ip in enumerate(ips_list, 1):
            res = await analizar_ip(session, ip, progress_bar, status_text, i, total)
            if res:
                resultados.append(res)
        return resultados

def generar_excel_ips_bytes(resultados):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte IPs"

    ws.append([
        "IP", "PAÍS", "REGIÓN", "CIUDAD",
        "RIESGO ABUSEIPDB (%)", "COMPORTAMIENTO MALICIOSO",
        "TIMESTAMPS (UTC)", "CATEGORÍAS ABUSEIPDB",
        "AMENAZAS VIRUSTOTAL"
    ])

    verde = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    amarillo = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    naranja = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    rojo = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    azul = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")

    conteo = {"0%": 0, "1-25%": 0, "26-60%": 0, "61-100%": 0}

    for r in resultados:
        ws.append([
            r["ip"], r["pais"], r["region"], r["ciudad"],
            f"{r['riesgo_abuse']}%", r["comportamiento"],
            r["timestamps"], r["categorias"],
            r["virustotal"]
        ])
        fila = ws.max_row

        if r["pais"] == "ECUADOR":
            ws[f"B{fila}"].fill = azul

        riesgo = r["riesgo_abuse"]
        celda = ws[f"E{fila}"]
        if riesgo == 0:
            celda.fill = verde
            conteo["0%"] += 1
        elif 1 <= riesgo <= 25:
            celda.fill = amarillo
            conteo["1-25%"] += 1
        elif 26 <= riesgo <= 60:
            celda.fill = naranja
            conteo["26-60%"] += 1
        elif riesgo > 60:
            celda.fill = rojo
            conteo["61-100%"] += 1

    resumen = wb.create_sheet("Resumen")
    resumen.append(["TOTAL DE IPs ESCANEADAS", len(resultados)])
    resumen.append([])
    resumen.append(["RANGO DE RIESGO", "CANTIDAD DE IPs"])
    for rango, cantidad in conteo.items():
        resumen.append([rango, cantidad])

    chart = BarChart()
    chart.title = "Distribución de IPs por Nivel de Riesgo"
    chart.x_axis.title = "Rango de Riesgo"
    chart.y_axis.title = "Cantidad"
    datos = Reference(resumen, min_col=2, min_row=5, max_row=8)
    categorias_ref = Reference(resumen, min_col=1, min_row=5, max_row=8)
    chart.add_data(datos, titles_from_data=False)
    chart.set_categories(categorias_ref)
    resumen.add_chart(chart, "D5")

    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()

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
# LECTURA ROBUSTA DE DATOS E HISTÓRICO
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1)
def cargar_datos():
    path = ARCHIVO_EXCEL if os.path.exists(ARCHIVO_EXCEL) else ("Base_Visitas.xlsx" if os.path.exists("Base_Visitas.xlsx") else None)
    if not path or os.path.getsize(path) == 0:
        df = pd.DataFrame()
    else:
        for _ in range(3):
            try:
                df = pd.read_excel(path, engine="openpyxl")
                break
            except Exception:
                time.sleep(0.3)
        else:
            df = pd.read_excel(path, engine="openpyxl")
    
    if not df.empty:
        df.columns = [str(c).strip() for c in df.columns]

        if "FechaCargaMatriz" not in df.columns:
            df["FechaCargaMatriz"] = obtener_hora_quito().strftime("%Y-%m-%d")

        col_region = "Región" if "Región" in df.columns else ("Region" if "Region" in df.columns else df.columns[0])
        col_coord = "Zona" if "Zona" in df.columns else ("Coordinador" if "Coordinador" in df.columns else df.columns[1])
        col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else ("Agencia" if "Agencia" in df.columns else df.columns[3])
        col_fecha = "FechaVisita" if "FechaVisita" in df.columns else ("Fecha de Visita" if "Fecha de Visita" in df.columns else "Fecha")

        if col_fecha in df.columns:
            df["Fecha_DT"] = pd.to_datetime(df[col_fecha], errors='coerce')
            df["Visitada"] = df["Fecha_DT"].apply(lambda x: "Sí" if pd.notnull(x) else "No")
        else:
            df["Fecha_DT"] = pd.NaT
            df["Visitada"] = "No"

        cols_eval = [c for c in ["Calidad", "ProcesosTransaccionales", "Procesos Transaccionales", 
                                "ManejoEfectivoControl", "Manejo Efectivo Control", 
                                "TalentoCultura", "Talento y Cultura"] if c in df.columns]
        
        if cols_eval:
            df["PromedioGral"] = df[cols_eval].apply(pd.to_numeric, errors='coerce').mean(axis=1)
        else:
            df["PromedioGral"] = 0.0

        return df, col_agencia, col_region, col_coord, col_fecha, cols_eval
    else:
        return df, "NombreAgencia", "Región", "Zona", "FechaVisita", []

df_base, col_agencia, col_region, col_coord, col_fecha, cols_eval = cargar_datos()

if 'region_activa' not in st.session_state: st.session_state['region_activa'] = "Todas"
if 'coordinador_activo' not in st.session_state: st.session_state['coordinador_activo'] = "Todos"
if 'filtro_kpi' not in st.session_state: st.session_state['filtro_kpi'] = "Todas"

def set_filtro_kpi(valor):
    st.session_state['filtro_kpi'] = valor

# -----------------------------------------------------------------------------
# GENERACIÓN DE INFORME WORD DETALLADO Y COMPLETO
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generar_reporte_word_detallado(df_data, col_agencia, col_region, col_coord, col_fecha, cols_eval, region_filtro, coord_filtro, fecha_matriz):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    COLOR_AZUL = RGBColor(2, 132, 199)
    COLOR_TITULO = RGBColor(15, 23, 42)
    HEX_HEADER = "0284C7"
    HEX_ALT = "F8FAFC"

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("INFORME EJECUTIVO Y AUDITORÍA DE OPERACIONES\nCONTROL DE GESTIÓN Y OPERACIÓN DE AGENCIAS")
    r_title.font.name = 'Segoe UI'
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_AZUL

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sub = p_sub.add_run(f"Fecha de emisión: {obtener_hora_quito().strftime('%d/%m/%Y %H:%M')}\nElaborado por Gabriela Armendariz\nAlcance Filtro: Región [{region_filtro}] | Zona [{coord_filtro}]")
    r_sub.font.size = Pt(9)
    r_sub.font.italic = True

    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Resumen Ejecutivo de Operaciones")
    r_h1.font.name = 'Segoe UI'
    r_h1.font.color.rgb = COLOR_TITULO

    tot_registros = len(df_data) if not df_data.empty else 0
    tot_vis = (df_data["Visitada"] == "Sí").sum() if not df_data.empty else 0
    tot_novis = tot_registros - tot_vis
    pct_cob = (tot_vis / tot_registros * 100) if tot_registros > 0 else 0.0

    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body1.add_run(
        f"El presente informe consolida la auditoría operativa y de supervisión para el segmento seleccionado "
        f"(Región: '{region_filtro}', Coordinador/Zona: '{coord_filtro}'). En la matriz de trabajo ('{fecha_matriz}') "
        f"se registra un universo total de {tot_registros:,} operaciones transaccionales/registros, de los cuales "
        f"{tot_vis:,} corresponden a supervisiones efectivas realizadas, alcanzando una cobertura operativa del {pct_cob:.1f}%. "
        f"Asimismo, se reportan {tot_novis:,} visitas o tareas pendientes."
    )

    table_kpi = doc.add_table(rows=2, cols=4)
    table_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    headers_kpi = ["Registros Totales", "Visitas Realizadas", "Visitas Pendientes", "% Cobertura Real"]
    values_kpi = [f"{tot_registros:,}", f"{tot_vis:,}", f"{tot_novis:,}", f"{pct_cob:.1f}%"]

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

    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Desglose Operativo por Coordinador / Zona")
    r_h2.font.name = 'Segoe UI'
    r_h2.font.color.rgb = COLOR_TITULO

    if not df_data.empty:
        resumen_coord = df_data.groupby(col_coord).agg(
            RegistrosTotales=(col_agencia, "count"),
            VisitasRealizadas=("Visitada", lambda x: (x == "Sí").sum()),
            VisitasPendientes=("Visitada", lambda x: (x == "No").sum())
        ).reset_index()

        table_coord = doc.add_table(rows=1, cols=4)
        table_coord.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_cells = table_coord.rows[0].cells
        cols_t = ["Coordinador / Zona", "Registros Totales", "Visitadas", "Pendientes"]
        for i, title in enumerate(cols_t):
            set_cell_background(hdr_cells[i], HEX_HEADER)
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in resumen_coord.iterrows():
            row_cells = table_coord.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            for col_idx, val in enumerate([row[col_coord], f"{row['RegistrosTotales']:,}", f"{row['VisitasRealizadas']:,}", f"{row['VisitasPendientes']:,}"]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(str(val))

    doc.add_paragraph()

    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Análisis de Operaciones y Tipos de Transacciones")
    r_h3.font.name = 'Segoe UI'
    r_h3.font.color.rgb = COLOR_TITULO

    if "TipoTran" in df_data.columns and not df_data.empty:
        term_unicas = df_data["Terminal"].nunique() if "Terminal" in df_data.columns else 0
        tipos_unicos = df_data["TipoTran"].nunique()
        
        p_trans = doc.add_paragraph()
        p_trans.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_trans.add_run(
            f"Se identifican un total de {term_unicas} terminales/cajas activas operando un catálogo de "
            f"{tipos_unicos} tipos de transacciones distintas. A continuación se detalla la distribución del volumen por operación:"
        )

        df_tipo = df_data["TipoTran"].value_counts().reset_index()
        df_tipo.columns = ["TipoTran", "Cantidad"]

        table_tipo = doc.add_table(rows=1, cols=3)
        table_tipo.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_type = table_tipo.rows[0].cells
        hdr_titles = ["Tipo de Transacción", "Cantidad Registrada", "% del Total"]
        for i, title in enumerate(hdr_titles):
            set_cell_background(hdr_type[i], HEX_HEADER)
            p = hdr_type[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in df_tipo.iterrows():
            row_cells = table_tipo.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            pct_val = (row["Cantidad"] / tot_registros * 100) if tot_registros > 0 else 0.0
            
            for col_idx, val in enumerate([row["TipoTran"], f"{row['Cantidad']:,}", f"{pct_val:.1f}%"]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(str(val))

    doc.add_paragraph()

    h4 = doc.add_heading(level=1)
    r_h4 = h4.add_run("4. Evaluaciones Cualitativas Promedio (Escala 1.0 - 3.0)")
    r_h4.font.name = 'Segoe UI'
    r_h4.font.color.rgb = COLOR_TITULO

    if cols_eval and not df_data.empty:
        df_eval_filtered = df_data[df_data["Visitada"] == "Sí"]
        table_eval = doc.add_table(rows=1, cols=2)
        table_eval.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        set_cell_background(table_eval.rows[0].cells[0], HEX_HEADER)
        set_cell_background(table_eval.rows[0].cells[1], HEX_HEADER)
        
        p0 = table_eval.rows[0].cells[0].paragraphs[0]
        r0 = p0.add_run("Dimensión de Evaluación")
        r0.font.bold = True
        r0.font.color.rgb = RGBColor(255, 255, 255)
        
        p1 = table_eval.rows[0].cells[1].paragraphs[0]
        r1 = p1.add_run("Puntaje Promedio")
        r1.font.bold = True
        r1.font.color.rgb = RGBColor(255, 255, 255)
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for idx, c in enumerate(cols_eval):
            val_mean = df_eval_filtered[c].apply(pd.to_numeric, errors='coerce').mean()
            row_cells = table_eval.add_row().cells
            bg = HEX_ALT if idx % 2 == 0 else "FFFFFF"
            set_cell_background(row_cells[0], bg)
            set_cell_background(row_cells[1], bg)
            
            p_lbl = row_cells[0].paragraphs[0]
            p_lbl.add_run(str(c))
            
            p_val = row_cells[1].paragraphs[0]
            p_val.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_val.add_run(f"{val_mean:.2f}" if pd.notnull(val_mean) else "N/A").font.bold = True

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

# PESTAÑAS PRINCIPALES (INCLUYENDO LA NUEVA PESTAÑA DE ESCANEO DE IPS)
if st.session_state.get('rol_actual') == "Administrador":
    tab_dash, tab_ips, tab_admin = st.tabs([
        "📊 Dashboard de Gestión & Operaciones", 
        "🌐 Escaneo & Reputación de IPs", 
        "⚙️ Módulo Administrador & Logs"
    ])
else:
    tab_dash, tab_ips = st.tabs([
        "📊 Dashboard de Gestión & Operaciones", 
        "🌐 Escaneo & Reputación de IPs"
    ])

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
        st.dataframe(df_kpi_vista[cols_mostrar], use_container_width=True, hide_index=True)

        st.markdown("---")

        col_izq, col_der = st.columns([1.1, 1])

        with col_izq:
            st.subheader("Resumen por Coordinador / Zona")
            resumen_coord = df_filtrado.groupby(col_coord).agg(
                RegistrosTotales=(col_agencia, "count"),
                VisitasRealizadas=("Visitada", lambda x: (x == "Sí").sum()),
                VisitasPendientes=("Visitada", lambda x: (x == "No").sum())
            ).reset_index()

            st.dataframe(resumen_coord, use_container_width=True, hide_index=True)

        with col_der:
            v_col, nv_col = st.columns(2)
            with v_col:
                st.subheader("Visitas Realizadas")
                st.dataframe(df_filtrado[df_filtrado["Visitada"] == "Sí"][[col_region, col_agencia]].drop_duplicates(), use_container_width=True, hide_index=True, height=220)
            with nv_col:
                st.subheader("Visitas Pendientes")
                st.dataframe(df_filtrado[df_filtrado["Visitada"] == "No"][[col_region, col_agencia]].drop_duplicates(), use_container_width=True, hide_index=True, height=220)

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
            st.dataframe(df_filtrado[cols_trans_mostrar], use_container_width=True, hide_index=True)
        else:
            st.info("Cargue una matriz con registros transaccionales para ver esta sección en detalle.")

# =============================================================================
# PESTAÑA 2: ESCANEO Y REPUTACIÓN DE IPS
# =============================================================================
with tab_ips:
    st.title("🌐 Escaneo & Reputación de Direcciones IP")
    st.markdown('<p class="author-credit">Elaborado por Gabriela Armendariz</p>', unsafe_allow_html=True)
    st.caption("Consulte niveles de riesgo, geolocalización y amenazas en tiempo real vía IPInfo, VirusTotal y AbuseIPDB.")

    col_in1, col_in2 = st.columns([1, 1])

    with col_in1:
        st.subheader("📥 Cargar Lista de IPs (.txt)")
        archivo_ips = st.file_uploader("Seleccione un archivo de texto con una IP por línea:", type=["txt"])

    with col_in2:
        st.subheader("📝 O ingrese IPs manualmente")
        ips_manual = st.text_area("Ingrese direcciones IP separadas por línea:", placeholder="45.70.58.1\n157.100.134.44", height=130)

    ips_a_procesar = []
    if archivo_ips is not None:
        contenido = archivo_ips.read().decode("utf-8")
        ips_a_procesar = [line.strip() for line in contenido.splitlines() if line.strip()]
    elif ips_manual.strip():
        ips_a_procesar = [line.strip() for line in ips_manual.splitlines() if line.strip()]

    if ips_a_procesar:
        st.info(f"📋 Total de IPs identificadas para escaneo: **{len(ips_a_procesar)}**")
        
        if st.button("🚀 Iniciar Escaneo de Reputación", use_container_width=True):
            progress_bar = st.progress(0.0)
            status_text = st.empty()

            try:
                # Ejecutar ciclo asíncrono
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                resultados_ips = loop.run_until_complete(
                    analizar_ips_async(ips_a_procesar, progress_bar, status_text)
                )
                loop.close()

                status_text.success("✅ ¡Escaneo completado exitosamente!")
                st.session_state['resultados_ips'] = resultados_ips

            except Exception as e:
                st.error(f"Error durante el escaneo: {e}")

    if 'resultados_ips' in st.session_state and st.session_state['resultados_ips']:
        resultados = st.session_state['resultados_ips']
        df_res = pd.DataFrame(resultados)

        st.markdown("---")
        st.subheader("📊 Resumen General del Escaneo")

        c1, c2, c3, c4 = st.columns(4)
        total_scan = len(df_res)
        cero_risk = len(df_res[df_res['riesgo_abuse'] == 0])
        med_risk = len(df_res[(df_res['riesgo_abuse'] >= 1) & (df_res['riesgo_abuse'] <= 60)])
        high_risk = len(df_res[df_res['riesgo_abuse'] > 60])

        c1.metric("IPs Escaneadas", total_scan)
        c2.metric("Sin Riesgo (0%)", cero_risk)
        c3.metric("Riesgo Medio (1-60%)", med_risk)
        c4.metric("Riesgo Alto (>60%)", high_risk)

        g_col1, g_col2 = st.columns([1, 1])

        with g_col1:
            df_res['RangoRiesgo'] = pd.cut(
                df_res['riesgo_abuse'],
                bins=[-1, 0, 25, 60, 100],
                labels=['0%', '1-25%', '26-60%', '61-100%']
            )
            df_rango = df_res['RangoRiesgo'].value_counts().reset_index()
            df_rango.columns = ['Rango', 'Cantidad']
            fig_rango = px.bar(df_rango, x='Rango', y='Cantidad', color='Rango',
                               title="Distribución de IPs por Rango de Riesgo",
                               color_discrete_map={'0%': '#00B050', '1-25%': '#FFC000', '26-60%': '#ED7D31', '61-100%': '#C00000'})
            st.plotly_chart(fig_rango, use_container_width=True)

        with g_col2:
            df_pais = df_res['pais'].value_counts().reset_index()
            df_pais.columns = ['País', 'Cantidad']
            fig_pais = px.pie(df_pais, names='País', values='Cantidad', title="Distribución Geográfica por País", hole=0.4)
            st.plotly_chart(fig_pais, use_container_width=True)

        st.subheader("📄 Matriz Detallada de Resultados")
        st.dataframe(df_res, use_container_width=True, hide_index=True)

        excel_ips_bytes = generar_excel_ips_bytes(resultados)
        st.download_button(
            label="📥 Descargar Reporte en Excel (resultado_ips.xlsx)",
            data=excel_ips_bytes,
            file_name=f"resultado_ips_{obtener_hora_quito().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# =============================================================================
# PESTAÑA 3: MÓDULO ADMINISTRADOR
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
                fecha_subida = st.date_input("🗓️ Fecha de Carga de la Matriz:", obtener_hora_quito())
                archivo_nuevo = st.file_uploader("📂 Seleccionar Matriz Excel (.xlsx)", type=["xlsx"])

                if archivo_nuevo and st.button("💾 Guardar y Registrar Matriz de la Fecha", use_container_width=True):
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
                    st.dataframe(df_logs_read.sort_index(ascending=False), use_container_width=True, hide_index=True)
                except Exception as e:
                    st.error(f"Error al leer el archivo de logs: {e}")
            else:
                st.info("Aún no se registran inicios de sesión.")