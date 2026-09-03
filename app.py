import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
# LECTURA ROBUSTA Y PERSISTENCIA DE DATOS
# -----------------------------------------------------------------------------
def cargar_datos():
    path = ARCHIVO_EXCEL if os.path.exists(ARCHIVO_EXCEL) else ("Base_Visitas.xlsx" if os.path.exists("Base_Visitas.xlsx") else None)
    
    if not path:
        # Generar dataframe vacío estructurado en caso de no existir base inicial
        df = pd.DataFrame(columns=["Región", "Zona", "NombreAgencia", "FechaVisita", "Calidad", "ProcesosTransaccionales"])
    else:
        df = pd.read_excel(path)
    
    df.columns = [str(c).strip() for c in df.columns]

    col_region = "Región" if "Región" in df.columns else ("Region" if "Region" in df.columns else df.columns[0])
    col_coord = "Zona" if "Zona" in df.columns else ("Coordinador" if "Coordinador" in df.columns else df.columns[1])
    col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else ("Agencia" if "Agencia" in df.columns else df.columns[3])
    col_fecha = "FechaVisita" if "FechaVisita" in df.columns else ("Fecha de Visita" if "Fecha de Visita" in df.columns else "Fecha")

    # Tratamiento de Fechas y Estado de Visitada
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

df_base, col_agencia, col_region, col_coord, col_fecha, cols_eval = cargar_datos()

# Inicialización de estado de sesión
if 'region_activa' not in st.session_state:
    st.session_state['region_activa'] = "Todas"
if 'coordinador_activo' not in st.session_state:
    st.session_state['coordinador_activo'] = "Todos"
if 'filtro_kpi' not in st.session_state:
    st.session_state['filtro_kpi'] = "Todas"

# -----------------------------------------------------------------------------
# ESTRUCTURA DE PESTAÑAS (DASHBOARD VS ADMINISTRADOR)
# -----------------------------------------------------------------------------
tab_dash, tab_admin = st.tabs(["📊 Dashboard de Gestión", "⚙️ Carga de Datos (Admin)"])

# =============================================================================
# PESTAÑA 1: DASHBOARD
# =============================================================================
with tab_dash:
    # FILTROS SIDEBAR
    st.sidebar.title("📌 Control de Gestión")

    # Filtro por Región
    regiones = ["Todas"] + sorted([str(x) for x in df_base[col_region].dropna().unique() if str(x).strip() != ""])
    idx_reg = regiones.index(st.session_state['region_activa']) if st.session_state['region_activa'] in regiones else 0
    region_sel = st.sidebar.selectbox("Filtrar por Región:", regiones, index=idx_reg)

    if region_sel != st.session_state['region_activa']:
        st.session_state['region_activa'] = region_sel
        st.session_state['coordinador_activo'] = "Todos"
        st.session_state['filtro_kpi'] = "Todas"
        st.rerun()

    df_region = df_base[df_base[col_region].astype(str) == st.session_state['region_activa']] if st.session_state['region_activa'] != "Todas" else df_base.copy()

    # Filtro por Coordinador
    coordinadores = ["Todos"] + sorted([str(x) for x in df_region[col_coord].dropna().unique() if str(x).strip() != ""])
    idx_coord = coordinadores.index(st.session_state['coordinador_activo']) if st.session_state['coordinador_activo'] in coordinadores else 0
    coord_sel = st.sidebar.selectbox("Filtrar por Coordinador/Zona:", coordinadores, index=idx_coord)

    if coord_sel != st.session_state['coordinador_activo']:
        st.session_state['coordinador_activo'] = coord_sel
        st.session_state['filtro_kpi'] = "Todas"
        st.rerun()

    # Filtro por Rango de Fechas
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Rango de Fechas")
    fechas_validas = df_base["Fecha_DT"].dropna()
    if not fechas_validas.empty:
        min_date = fechas_validas.min().date()
        max_date = fechas_validas.max().date()
        f_inicio, f_fin = st.sidebar.date_input("Seleccionar Periodo:", [min_date, max_date])
    else:
        f_inicio, f_fin = None, None

    if st.sidebar.button("🔄 Restablecer Filtros"):
        st.session_state['region_activa'] = "Todas"
        st.session_state['coordinador_activo'] = "Todos"
        st.session_state['filtro_kpi'] = "Todas"
        st.rerun()

    # Aplicar filtros
    df_filtrado = df_base.copy()
    if st.session_state['region_activa'] != "Todas":
        df_filtrado = df_filtrado[df_filtrado[col_region].astype(str) == st.session_state['region_activa']]
    if st.session_state['coordinador_activo'] != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_coord].astype(str) == st.session_state['coordinador_activo']]
    if f_inicio and f_fin:
        df_filtrado = df_filtrado[(df_filtrado["Fecha_DT"].dt.date >= f_inicio) & (df_filtrado["Fecha_DT"].dt.date <= f_fin) | (df_filtrado["Fecha_DT"].isna())]

    # MÉTRICAS
    tot_asig = df_filtrado[col_agencia].nunique()
    tot_vis = df_filtrado[df_filtrado["Visitada"] == "Sí"][col_agencia].nunique()
    tot_novis = tot_asig - tot_vis
    pct_cob = (tot_vis / tot_asig * 100) if tot_asig > 0 else 0.0

    st.title("📊 Control de Gestión de Agencias")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f'<div class="kpi-card kpi-asig"><div class="kpi-title">Agencias Asignadas</div><div class="kpi-number">{tot_asig}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 Ver Asignadas", key="btn_asig", use_container_width=True): st.session_state['filtro_kpi'] = "Todas"
    with k2:
        st.markdown(f'<div class="kpi-card kpi-vis"><div class="kpi-title">Agencias Visitadas</div><div class="kpi-number">{tot_vis}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 Ver Visitadas", key="btn_vis", use_container_width=True): st.session_state['filtro_kpi'] = "Visitadas"
    with k3:
        st.markdown(f'<div class="kpi-card kpi-novis"><div class="kpi-title">Agencias No Visitadas</div><div class="kpi-number">{tot_novis}</div></div>', unsafe_allow_html=True)
        if st.button("🔍 Ver Pendientes", key="btn_novis", use_container_width=True): st.session_state['filtro_kpi'] = "No Visitadas"
    with k4:
        st.markdown(f'<div class="kpi-card kpi-cob"><div class="kpi-title">% Cobertura</div><div class="kpi-number">{pct_cob:.1f}%</div></div>', unsafe_allow_html=True)
        if st.button("🔍 Ver Cobertura", key="btn_cob", use_container_width=True): st.session_state['filtro_kpi'] = "Todas"

    st.markdown("<br>", unsafe_allow_html=True)

    # TABLA DE DETALLE SEGÚN BOTÓN CLIQUEADO
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
        cols_mostrar = [c for c in [col_region, col_coord, col_agencia, "Visitada", col_fecha, "PromedioGral"] if c in df_kpi_vista.columns]
        st.dataframe(df_kpi_vista[cols_mostrar].drop_duplicates(), use_container_width=True, hide_index=True)

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

        st.dataframe(resumen_coord, use_container_width=True, hide_index=True)

        # GRÁFICO 100% DINÁMICO POR AGENCIA Y FECHA
        st.subheader("Gestión Individual por Agencia")
        
        df_gestion_ag = df_filtrado.groupby(col_agencia, as_index=False)["Visitada"].first()
        
        fig_agencias = px.bar(
            df_gestion_ag,
            x=col_agencia,
            y=[1]*len(df_gestion_ag),
            color="Visitada",
            color_discrete_map={"Sí": "#16A34A", "No": "#DC2626"},
            title="Estado de Cobertura por Agencia (Selección Actual)",
            labels={col_agencia: "Agencia", "value": "Estado"},
            text="Visitada"
        )
        fig_agencias.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#1E293B"),
            height=280,
            yaxis_visible=False,
            margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_agencias, use_container_width=True)

    with col_der:
        v_col, nv_col = st.columns(2)
        with v_col:
            st.subheader("Agencias Visitadas")
            st.dataframe(df_filtrado[df_filtrado["Visitada"] == "Sí"][[col_region, col_agencia]].drop_duplicates(), use_container_width=True, hide_index=True, height=280)
        with nv_col:
            st.subheader("Agencias No Visitadas")
            st.dataframe(df_filtrado[df_filtrado["Visitada"] == "No"][[col_region, col_agencia]].drop_duplicates(), use_container_width=True, hide_index=True, height=280)

# =============================================================================
# PESTAÑA 2: MÓDULO ADMINISTRADOR (CARGA DIARIA DE EXCEL)
# =============================================================================
with tab_admin:
    st.title("⚙️ Módulo Administrador: Carga Diaria de Datos")
    st.info("Este módulo permite al administrador consolidar información diariamente sin sobrescribir los datos históricos.")

    clave = st.text_input("🔑 Ingrese Contraseña de Administrador:", type="password")

    if clave == "admin123":  # Contraseña por defecto
        st.success("Acceso Autorizado")
        
        archivo_nuevo = st.file_uploader("📂 Cargar nuevo archivo Excel diario (.xlsx)", type=["xlsx"])
        modo_carga = st.radio("Modo de integración:", ["Acumular (Agregar nuevos datos al histórico)", "Reemplazar base completa"])

        if archivo_nuevo and st.button("💾 Procesar y Guardar en la Base"):
            try:
                df_nuevo = pd.read_excel(archivo_nuevo)
                
                if modo_carga == "Acumular (Agregar nuevos datos al histórico)":
                    df_consolidado = pd.concat([df_base, df_nuevo], ignore_index=True).drop_duplicates()
                else:
                    df_consolidado = df_nuevo

                # Guardar cambios
                df_consolidado.to_excel(ARCHIVO_EXCEL, index=False)
                st.success("¡Base de datos actualizada con éxito! Refrescando la aplicación...")
                st.rerun()

            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")

    elif clave != "":
        st.error("Contraseña incorrecta. Acceso restringido.")