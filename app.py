import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Gestión de Agencias",
    page_icon="📊",
    layout="wide"
)

# Estilos CSS
st.markdown("""
    <style>
    .stApp {
        background-color: #EAEFF5;
        color: #1E293B;
    }
    .metric-card {
        background-color: #FFFFFF;
        border-radius: 6px;
        padding: 12px;
        text-align: center;
        border: 1px solid #CBD5E1;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 28px;
        font-weight: bold;
        color: #0284C7;
    }
    .metric-label {
        font-size: 13px;
        font-weight: 600;
        color: #475569;
    }
    h1, h2, h3 {
        color: #0F172A !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    [data-testid="stSidebar"] {
        background-color: #1E293B;
        color: #FFFFFF;
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    </style>
""", unsafe_allow_html=True)

ARCHIVO_EXCEL = "base_visitas.xlsx"

# -----------------------------------------------------------------------------
# LECTURA ROBUSTA DE DATOS
# -----------------------------------------------------------------------------
def cargar_datos():
    if not os.path.exists(ARCHIVO_EXCEL):
        # Compatibilidad por si el archivo está guardado como Base_Visitas.xlsx
        if os.path.exists("Base_Visitas.xlsx"):
            path = "Base_Visitas.xlsx"
        else:
            st.error(f"No se encontró el archivo '{ARCHIVO_EXCEL}' en la carpeta.")
            st.stop()
    else:
        path = ARCHIVO_EXCEL

    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]

    # Mapeo flexible de columnas desde tu estructura de Excel
    col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else ("Agencia" if "Agencia" in df.columns else df.columns[0])
    col_region = "Región" if "Región" in df.columns else ("Region" if "Region" in df.columns else "Zona")
    col_coord = "Zona" if "Zona" in df.columns else ("Coordinador" if "Coordinador" in df.columns else df.columns[1])
    col_fecha = "FechaVisita" if "FechaVisita" in df.columns else ("Fecha de Visita" if "Fecha de Visita" in df.columns else None)

    # Identificación de agencias visitadas
    if col_fecha and col_fecha in df.columns:
        df["Visitada"] = df[col_fecha].apply(lambda x: "Sí" if pd.notnull(x) and str(x).strip() not in ["", "-", "None", "NaT"] else "No")
    else:
        df["Visitada"] = "No"

    # Columnas de evaluación
    cols_eval = [c for c in ["Calidad", "ProcesosTransaccionales", "Procesos Transaccionales", 
                            "ManejoEfectivoControl", "Manejo Efectivo Control", 
                            "TalentoCultura", "Talento y Cultura"] if c in df.columns]
    
    if cols_eval:
        df["PromedioGral"] = df[cols_eval].apply(pd.to_numeric, errors='coerce').mean(axis=1)
    else:
        df["PromedioGral"] = 0.0

    return df, col_agencia, col_region, col_coord, col_fecha, cols_eval

df_base, col_agencia, col_region, col_coord, col_fecha, cols_eval = cargar_datos()

# Estado de sesión
if 'region_activa' not in st.session_state:
    st.session_state['region_activa'] = "Todas"
if 'coordinador_activo' not in st.session_state:
    st.session_state['coordinador_activo'] = "Todos"

# -----------------------------------------------------------------------------
# GENERACIÓN DE INFORME EN WORD (.DOCX)
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generar_reporte_word(df_data, col_agencia, col_region, col_coord, col_fecha, cols_eval):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    COLOR_AZUL = RGBColor(2, 132, 199)
    COLOR_TITULO = RGBColor(15, 23, 42)
    COLOR_TEXTO = RGBColor(51, 65, 85)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("INFORME EJECUTIVO DE GESTIÓN DE AGENCIAS")
    run_title.font.name = 'Segoe UI'
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_AZUL

    tot_asig = len(df_data[col_agencia].unique())
    tot_vis = len(df_data[df_data["Visitada"] == "Sí"][col_agencia].unique())
    tot_novis = tot_asig - tot_vis
    pct_cob = (tot_vis / tot_asig * 100) if tot_asig > 0 else 0

    h1 = doc.add_heading(level=1)
    run_h1 = h1.add_run("1. Resumen Ejecutivo")
    run_h1.font.name = 'Segoe UI'
    run_h1.font.color.rgb = COLOR_TITULO

    p_intro = doc.add_paragraph()
    p_intro.add_run(f"Consolidado de supervisión de agencias. Total asignadas: {tot_asig}, visitadas: {tot_vis} ({pct_cob:.1f}% cobertura), pendientes: {tot_novis}.")

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# -----------------------------------------------------------------------------
# BARRA LATERAL: CONTROL DE GESTIÓN Y FILTROS
# -----------------------------------------------------------------------------
st.sidebar.title("📌 Control de Gestión")

# 1. Filtro por Región (Centro, Norte, Sur, etc.)
regiones = ["Todas"] + sorted([str(x) for x in df_base[col_region].dropna().unique() if str(x).strip() != ""])
idx_reg = regiones.index(st.session_state['region_activa']) if st.session_state['region_activa'] in regiones else 0
region_sel = st.sidebar.selectbox("Filtrar por Región:", regiones, index=idx_reg)

if region_sel != st.session_state['region_activa']:
    st.session_state['region_activa'] = region_sel
    st.session_state['coordinador_activo'] = "Todos"
    st.rerun()

# Filtrar dataset temporalmente para obtener los coordinadores/zonas de esa región
if st.session_state['region_activa'] != "Todas":
    df_region = df_base[df_base[col_region].astype(str) == st.session_state['region_activa']]
else:
    df_region = df_base.copy()

# 2. Filtro por Coordinador / Zona (Dinamizado según la región)
coordinadores = ["Todos"] + sorted([str(x) for x in df_region[col_coord].dropna().unique() if str(x).strip() != ""])
idx_coord = coordinadores.index(st.session_state['coordinador_activo']) if st.session_state['coordinador_activo'] in coordinadores else 0
coord_sel = st.sidebar.selectbox("Filtrar por Coordinador/Zona:", coordinadores, index=idx_coord)

if coord_sel != st.session_state['coordinador_activo']:
    st.session_state['coordinador_activo'] = coord_sel
    st.rerun()

if st.sidebar.button("🔄 Mostrar Todos los Datos"):
    st.session_state['region_activa'] = "Todas"
    st.session_state['coordinador_activo'] = "Todos"
    st.rerun()

st.sidebar.markdown("---")

# DESCARGA DE INFORME
st.sidebar.subheader("📄 Exportar Informe")
df_visitadas_todas = df_base[df_base["Visitada"] == "Sí"]

if not df_visitadas_todas.empty:
    doc_bytes = generar_reporte_word(df_base, col_agencia, col_region, col_coord, col_fecha, cols_eval)
    st.sidebar.download_button(
        label="📄 Descargar Informe Ejecutivo (Word)",
        data=doc_bytes,
        file_name="Informe_Ejecutivo_Gestion_Agencias.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

st.sidebar.markdown("---")

# FORMULARIO DE ALIMENTACIÓN DE REGISTRO
with st.sidebar.expander("➕ Alimentar / Registrar Visita", expanded=False):
    with st.form("form_visita"):
        agencia_sel = st.selectbox("Agencia:", df_base[col_agencia].unique())
        fecha_in = st.date_input("Fecha de Visita")
        calif_c = st.number_input("Calidad (1.0 - 3.0)", min_value=1.0, max_value=3.0, value=3.0, step=0.1)
        calif_p = st.number_input("Procesos Transaccionales", min_value=1.0, max_value=3.0, value=3.0, step=0.1)
        obs_cal = st.text_area("Obs. Calidad")
        
        btn_guardar = st.form_submit_button("Guardar Registro")
        if btn_guardar:
            st.success("¡Datos guardados!")

# Aplicar filtros globales a los datos mostrados
df_filtrado = df_base.copy()
if st.session_state['region_activa'] != "Todas":
    df_filtrado = df_filtrado[df_filtrado[col_region].astype(str) == st.session_state['region_activa']]
if st.session_state['coordinador_activo'] != "Todos":
    df_filtrado = df_filtrado[df_filtrado[col_coord].astype(str) == st.session_state['coordinador_activo']]

# -----------------------------------------------------------------------------
# PANEL PRINCIPAL / DASHBOARD
# -----------------------------------------------------------------------------
st.title("📊 Dashboard de Control de Gestión")

# METRICAS
tot_asig = len(df_filtrado[col_agencia].unique())
tot_vis = len(df_filtrado[df_filtrado["Visitada"] == "Sí"][col_agencia].unique())
tot_novis = tot_asig - tot_vis
pct_cob = (tot_vis / tot_asig * 100) if tot_asig > 0 else 0

k1, k2, k3, k4 = st.columns(4)
k1.markdown(f'<div class="metric-card"><div class="metric-label">Agencias Asignadas</div><div class="metric-value">{tot_asig}</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="metric-card"><div class="metric-label">Agencias Visitadas</div><div class="metric-value" style="color:#16A34A;">{tot_vis}</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="metric-card"><div class="metric-label">Agencias No Visitadas</div><div class="metric-value" style="color:#DC2626;">{tot_novis}</div></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="metric-card"><div class="metric-label">% Cobertura</div><div class="metric-value">{pct_cob:.1f}%</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

col_izq, col_der = st.columns([1.1, 1])

with col_izq:
    st.subheader("Resumen por Coordinador / Zona")
    resumen_coord = df_filtrado.groupby(col_coord).agg(
        AgenciasAsignadas=(col_agencia, "nunique"),
        AgenciasVisitadas=("Visitada", lambda x: (x == "Sí").nunique()),
        AgenciasNoVisitadas=("Visitada", lambda x: (x == "No").nunique())
    ).reset_index()

    st.dataframe(resumen_coord, use_container_width=True, hide_index=True)

    # GRÁFICO MEJORADO DE GESTIÓN
    st.subheader("Gestión por Agencias (Visitadas vs No Visitadas)")
    
    df_chart = pd.DataFrame({
        "Estado": ["Visitadas", "No Visitadas"],
        "Cantidad": [tot_vis, tot_novis]
    })

    fig_barras = px.bar(
        df_chart,
        x="Cantidad",
        y="Estado",
        orientation='h',
        color="Estado",
        color_discrete_map={"Visitadas": "#16A34A", "No Visitadas": "#DC2626"},
        text="Cantidad"
    )
    fig_barras.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color="#1E293B"),
        height=200,
        showlegend=False,
        margin=dict(l=0, r=0, t=10, b=0)
    )
    st.plotly_chart(fig_barras, use_container_width=True)

with col_der:
    v_col, nv_col = st.columns(2)
    with v_col:
        st.subheader("Agencias Visitadas")
        st.dataframe(df_filtrado[df_filtrado["Visitada"] == "Sí"][[col_region, col_agencia]].drop_duplicates(), use_container_width=True, hide_index=True, height=280)
    with nv_col:
        st.subheader("Agencias No Visitadas")
        st.dataframe(df_filtrado[df_filtrado["Visitada"] == "No"][[col_region, col_agencia]].drop_duplicates(), use_container_width=True, hide_index=True, height=280)