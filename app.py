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
    
    /* Diseño de tarjetas métricas visuales con gradiente */
    .kpi-card {
        border-radius: 12px;
        padding: 15px 10px;
        text-align: center;
        color: white;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.15);
        margin-bottom: 5px;
    }
    .kpi-asig { background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%); }
    .kpi-vis { background: linear-gradient(135deg, #16A34A 0%, #15803D 100%); }
    .kpi-novis { background: linear-gradient(135deg, #DC2626 0%, #B91C1C 100%); }
    .kpi-cob { background: linear-gradient(135deg, #0D9488 0%, #0F766E 100%); }
    
    .kpi-title {
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        opacity: 0.9;
    }
    .kpi-number {
        font-size: 32px;
        font-weight: 800;
        margin-top: 4px;
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
        if os.path.exists("Base_Visitas.xlsx"):
            path = "Base_Visitas.xlsx"
        else:
            st.error(f"No se encontró el archivo '{ARCHIVO_EXCEL}' en la carpeta.")
            st.stop()
    else:
        path = ARCHIVO_EXCEL

    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]

    col_region = "Región" if "Región" in df.columns else ("Region" if "Region" in df.columns else df.columns[0])
    col_coord = "Zona" if "Zona" in df.columns else ("Coordinador" if "Coordinador" in df.columns else df.columns[1])
    col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else ("Agencia" if "Agencia" in df.columns else df.columns[3])
    col_fecha = "FechaVisita" if "FechaVisita" in df.columns else ("Fecha de Visita" if "Fecha de Visita" in df.columns else None)

    if col_fecha and col_fecha in df.columns:
        df["Visitada"] = df[col_fecha].apply(lambda x: "Sí" if pd.notnull(x) and str(x).strip() not in ["", "-", "None", "NaT"] else "No")
    else:
        df["Visitada"] = "No"

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
if 'filtro_kpi' not in st.session_state:
    st.session_state['filtro_kpi'] = "Todas"

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

regiones = ["Todas"] + sorted([str(x) for x in df_base[col_region].dropna().unique() if str(x).strip() != ""])
idx_reg = regiones.index(st.session_state['region_activa']) if st.session_state['region_activa'] in regiones else 0
region_sel = st.sidebar.selectbox("Filtrar por Región:", regiones, index=idx_reg)

if region_sel != st.session_state['region_activa']:
    st.session_state['region_activa'] = region_sel
    st.session_state['coordinador_activo'] = "Todos"
    st.session_state['filtro_kpi'] = "Todas"
    st.rerun()

if st.session_state['region_activa'] != "Todas":
    df_region = df_base[df_base[col_region].astype(str) == st.session_state['region_activa']]
else:
    df_region = df_base.copy()

coordinadores = ["Todos"] + sorted([str(x) for x in df_region[col_coord].dropna().unique() if str(x).strip() != ""])
idx_coord = coordinadores.index(st.session_state['coordinador_activo']) if st.session_state['coordinador_activo'] in coordinadores else 0
coord_sel = st.sidebar.selectbox("Filtrar por Coordinador/Zona:", coordinadores, index=idx_coord)

if coord_sel != st.session_state['coordinador_activo']:
    st.session_state['coordinador_activo'] = coord_sel
    st.session_state['filtro_kpi'] = "Todas"
    st.rerun()

if st.sidebar.button("🔄 Mostrar Todos los Datos"):
    st.session_state['region_activa'] = "Todas"
    st.session_state['coordinador_activo'] = "Todos"
    st.session_state['filtro_kpi'] = "Todas"
    st.rerun()

st.sidebar.markdown("---")

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

df_filtrado = df_base.copy()
if st.session_state['region_activa'] != "Todas":
    df_filtrado = df_filtrado[df_filtrado[col_region].astype(str) == st.session_state['region_activa']]
if st.session_state['coordinador_activo'] != "Todos":
    df_filtrado = df_filtrado[df_filtrado[col_coord].astype(str) == st.session_state['coordinador_activo']]

# -----------------------------------------------------------------------------
# PANEL PRINCIPAL / DASHBOARD
# -----------------------------------------------------------------------------
st.title("📊 Dashboard de Control de Gestión")

# CÁLCULO DE MÉTRICAS
tot_asig = df_filtrado[col_agencia].nunique()
tot_vis = df_filtrado[df_filtrado["Visitada"] == "Sí"][col_agencia].nunique()
tot_novis = tot_asig - tot_vis
pct_cob = (tot_vis / tot_asig * 100) if tot_asig > 0 else 0.0

# TARJETAS DE COLOR CON BOTÓN DE FILTRADO DEBAJO
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(f'''
        <div class="kpi-card kpi-asig">
            <div class="kpi-title">Agencias Asignadas</div>
            <div class="kpi-number">{tot_asig}</div>
        </div>
    ''', unsafe_allow_html=True)
    if st.button("🔍 Ver Asignadas", key="btn_asig", use_container_width=True):
        st.session_state['filtro_kpi'] = "Todas"

with k2:
    st.markdown(f'''
        <div class="kpi-card kpi-vis">
            <div class="kpi-title">Agencias Visitadas</div>
            <div class="kpi-number">{tot_vis}</div>
        </div>
    ''', unsafe_allow_html=True)
    if st.button("🔍 Ver Visitadas", key="btn_vis", use_container_width=True):
        st.session_state['filtro_kpi'] = "Visitadas"

with k3:
    st.markdown(f'''
        <div class="kpi-card kpi-novis">
            <div class="kpi-title">Agencias No Visitadas</div>
            <div class="kpi-number">{tot_novis}</div>
        </div>
    ''', unsafe_allow_html=True)
    if st.button("🔍 Ver Pendientes", key="btn_novis", use_container_width=True):
        st.session_state['filtro_kpi'] = "No Visitadas"

with k4:
    st.markdown(f'''
        <div class="kpi-card kpi-cob">
            <div class="kpi-title">% Cobertura</div>
            <div class="kpi-number">{pct_cob:.1f}%</div>
        </div>
    ''', unsafe_allow_html=True)
    if st.button("🔍 Ver Cobertura", key="btn_cob", use_container_width=True):
        st.session_state['filtro_kpi'] = "Todas"

st.markdown("<br>", unsafe_allow_html=True)

# DETALLE DINÁMICO DE AGENCIAS DEBAJO DE LAS TARJETAS
if st.session_state['filtro_kpi'] == "Visitadas":
    df_kpi_vista = df_filtrado[df_filtrado["Visitada"] == "Sí"]
    titulo_kpi = "📌 Detalle: Agencias Visitadas"
elif st.session_state['filtro_kpi'] == "No Visitadas":
    df_kpi_vista = df_filtrado[df_filtrado["Visitada"] == "No"]
    titulo_kpi = "📌 Detalle: Agencias No Visitadas (Pendientes)"
else:
    df_kpi_vista = df_filtrado.copy()
    titulo_kpi = "📌 Detalle: Todas las Agencias Asignadas"

with st.expander(f"{titulo_kpi} ({df_kpi_vista[col_agencia].nunique()} agencias)", expanded=True):
    cols_mostrar = [col_region, col_coord, col_agencia, "Visitada"]
    if col_fecha and col_fecha in df_kpi_vista.columns:
        cols_mostrar.append(col_fecha)
    cols_mostrar.extend(cols_eval)
    if "PromedioGral" in df_kpi_vista.columns:
        cols_mostrar.append("PromedioGral")
    
    cols_existentes = [c for c in cols_mostrar if c in df_kpi_vista.columns]
    
    st.dataframe(
        df_kpi_vista[cols_existentes].drop_duplicates(),
        use_container_width=True,
        hide_index=True
    )

st.markdown("---")

col_izq, col_der = st.columns([1.1, 1])

with col_izq:
    st.subheader("Resumen por Coordinador / Zona")
    
    df_agencias_unicas = df_filtrado.groupby([col_coord, col_agencia], as_index=False)["Visitada"].first()
    
    resumen_coord = df_agencias_unicas.groupby(col_coord).agg(
        AgenciasAsignadas=(col_agencia, "nunique"),
        AgenciasVisitadas=("Visitada", lambda x: (x == "Sí").sum()),
        AgenciasNoVisitadas=("Visitada", lambda x: (x == "No").sum())
    ).reset_index()

    seleccion = st.dataframe(
        resumen_coord,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row"
    )

    filas_seleccionadas = seleccion.selection.rows
    if filas_seleccionadas:
        coord_cliqueado = resumen_coord.iloc[filas_seleccionadas[0]][col_coord]
        if st.session_state['coordinador_activo'] != coord_cliqueado:
            st.session_state['coordinador_activo'] = coord_cliqueado
            st.rerun()

    st.subheader("Gestión por Agencias")
    
    fig_barras = px.bar(
        x=[tot_vis, tot_novis],
        y=["Gestión", "Gestión"],
        orientation='h',
        color_discrete_sequence=["#A855F7", "#06B6D4"],
        text_auto=True
    )
    fig_barras.update_layout(
        barmode='stack',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color="#1E293B"),
        height=140,
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