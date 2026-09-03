import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# Configuración de la página
st.set_page_config(
    page_title="Dashboard de Gestión de Agencias",
    page_icon="📊",
    layout="wide"
)

# Estilos CSS - Tema Claro tipo Power BI
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
# LECTURA DE DATOS
# -----------------------------------------------------------------------------
def cargar_datos():
    if not os.path.exists(ARCHIVO_EXCEL):
        st.error(f"No se encontró el archivo '{ARCHIVO_EXCEL}' en la carpeta.")
        st.stop()

    df = pd.read_excel(ARCHIVO_EXCEL)
    df.columns = [str(c).strip() for c in df.columns]

    col_agencia = "NombreAgencia" if "NombreAgencia" in df.columns else ("Agencia" if "Agencia" in df.columns else df.columns[0])
    col_region = "Región" if "Región" in df.columns else ("Region" if "Region" in df.columns else df.columns[1])
    col_coord = "Coordinador" if "Coordinador" in df.columns else "Zona"
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

if 'coordinador_activo' not in st.session_state:
    st.session_state['coordinador_activo'] = "Todos"

# -----------------------------------------------------------------------------
# GENERADOR DE INFORME EJECUTIVO EN WORD (.DOCX)
# -----------------------------------------------------------------------------
def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def generar_reporte_word(df_data, col_agencia, col_region, col_coord, col_fecha, cols_eval):
    doc = Document()
    
    # Márgenes del documento
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # ESTILOS Y COLORES CORPORATIVOS
    COLOR_AZUL = RGBColor(2, 132, 199)
    COLOR_TITULO = RGBColor(15, 23, 42)
    COLOR_TEXTO = RGBColor(51, 65, 85)

    # 1. PORTADA / ENCABEZADO PRINCIPAL
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("INFORME EJECUTIVO DE GESTIÓN Y AUDITORÍA DE AGENCIAS")
    run_title.font.name = 'Segoe UI'
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = COLOR_AZUL

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run("Evaluación Operativa, Control de Visitas y Recomendaciones Estratégicas")
    run_sub.font.name = 'Segoe UI'
    run_sub.font.size = Pt(13)
    run_sub.font.italic = True
    run_sub.font.color.rgb = COLOR_TEXTO

    doc.add_paragraph().paragraph_format.space_after = Pt(20)

    # 2. RESUMEN EJECUTIVO
    h1 = doc.add_heading(level=1)
    run_h1 = h1.add_run("1. Resumen Ejecutivo")
    run_h1.font.name = 'Segoe UI'
    run_h1.font.color.rgb = COLOR_TITULO

    tot_asig = len(df_data[col_agencia].unique())
    tot_vis = len(df_data[df_data["Visitada"] == "Sí"][col_agencia].unique())
    tot_novis = tot_asig - tot_vis
    pct_cob = (tot_vis / tot_asig * 100) if tot_asig > 0 else 0

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.line_spacing = 1.15
    p_intro.add_run(
        f"El presente informe consolida los resultados obtenidos durante el proceso de supervisión e inspección "
        f"de la red de agencias. A la fecha, se registra un total de "
    )
    r_bold1 = p_intro.add_run(f"{tot_asig} agencias asignadas")
    r_bold1.bold = True
    p_intro.add_run(", de las cuales ")
    r_bold2 = p_intro.add_run(f"{tot_vis} han sido supervisadas ({pct_cob:.1f}% de cobertura)")
    r_bold2.bold = True
    p_intro.add_run(f", quedando {tot_novis} agencias pendientes de inspección.")

    # Tabla Resumen de Cobertura
    table_kpi = doc.add_table(rows=2, cols=4)
    table_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table_kpi.rows[0].cells
    kpi_titles = ["Agencias Asignadas", "Agencias Visitadas", "Agencias No Visitadas", "% Cobertura"]
    for i, title in enumerate(kpi_titles):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1E293B")
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(9)

    val_cells = table_kpi.rows[1].cells
    values = [str(tot_asig), str(tot_vis), str(tot_novis), f"{pct_cob:.1f}%"]
    for i, val in enumerate(values):
        val_cells[i].text = val
        set_cell_background(val_cells[i], "F1F5F9")
        for paragraph in val_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = COLOR_AZUL

    doc.add_paragraph().paragraph_format.space_after = Pt(15)

    # 3. DETALLE DE AGENCIAS VISITADAS Y CALIFICACIONES
    h2 = doc.add_heading(level=1)
    run_h2 = h2.add_run("2. Evaluación Detallada por Agencia Visitada")
    run_h2.font.name = 'Segoe UI'
    run_h2.font.color.rgb = COLOR_TITULO

    df_vis = df_data[df_data["Visitada"] == "Sí"].copy()

    if not df_vis.empty:
        # Generación de la tabla detallada
        cols_headers = ["Agencia", "Coordinador", "Fecha Visita", "Promedio Gral.", "Estado"]
        table_det = doc.add_table(rows=1, cols=len(cols_headers))
        table_det.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_det = table_det.rows[0].cells
        for i, h_text in enumerate(cols_headers):
            hdr_det[i].text = h_text
            set_cell_background(hdr_det[i], "0284C7")
            for paragraph in hdr_det[i].paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.size = Pt(9)

        for _, row in df_vis.iterrows():
            row_cells = table_det.add_row().cells
            prom = row.get("PromedioGral", 0)
            
            if prom < 2.0:
                estado_txt = "Crítico"
                bg_color = "FEE2E2"
            elif prom < 2.8:
                estado_txt = "Aceptable"
                bg_color = "FEF3C7"
            else:
                estado_txt = "Excelente"
                bg_color = "DCFCE7"

            row_cells[0].text = str(row.get(col_agencia, ""))
            row_cells[1].text = str(row.get(col_coord, ""))
            row_cells[2].text = str(row.get(col_fecha, ""))
            row_cells[3].text = f"{prom:.2f}"
            row_cells[4].text = estado_txt

            for i, cell in enumerate(row_cells):
                set_cell_background(cell, bg_color if i == 4 else "FFFFFF")
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in [2, 3, 4] else WD_ALIGN_PARAGRAPH.LEFT
                    for r in p.runs:
                        r.font.size = Pt(8.5)

        doc.add_paragraph().paragraph_format.space_after = Pt(15)

        # 4. OBSERVACIONES Y HALLAZGOS
        h3 = doc.add_heading(level=1)
        run_h3 = h3.add_run("3. Observaciones y Hallazgos de Campo")
        run_h3.font.name = 'Segoe UI'
        run_h3.font.color.rgb = COLOR_TITULO

        for _, row in df_vis.iterrows():
            p_ag = doc.add_paragraph()
            r_ag = p_ag.add_run(f"• {row.get(col_agencia, '')} ({row.get(col_coord, '')}):")
            r_ag.bold = True
            
            p_obs1 = doc.add_paragraph()
            p_obs1.paragraph_format.left_indent = Inches(0.25)
            p_obs1.add_run("Observaciones de Calidad: ").bold = True
            p_obs1.add_run(str(row.get("ObsCalidad", "Sin hallazgos registrados.")))

            p_obs2 = doc.add_paragraph()
            p_obs2.paragraph_format.left_indent = Inches(0.25)
            p_obs2.add_run("Proceso Operativo: ").bold = True
            p_obs2.add_run(str(row.get("ObsProcesoOperativo", "Sin novedades en arqueo/bóveda.")))

        # 5. RECOMENDACIONES ESTRATÉGICAS Y PLAN DE ACCIÓN
        doc.add_paragraph().paragraph_format.space_after = Pt(10)
        h4 = doc.add_heading(level=1)
        run_h4 = h4.add_run("4. Recomendaciones Estratégicas y Plan de Acción")
        run_h4.font.name = 'Segoe UI'
        run_h4.font.color.rgb = COLOR_TITULO

        rec1 = doc.add_paragraph(style='List Bullet')
        rec1.add_run("Plan de Remediación Inmediata: ").bold = True
        rec1.add_run("Para agencias con promedios inferiores a 2.80, coordinar auditorías operativas en un plazo no mayor a 15 días hábiles.")

        rec2 = doc.add_paragraph(style='List Bullet')
        rec2.add_run("Estandarización de Controles Operativos: ").bold = True
        rec2.add_run("Reforzar los protocolos de cuadre de bóveda, cuadre de terminales y revisión de papeletas transaccionales.")

        rec3 = doc.add_paragraph(style='List Bullet')
        rec3.add_run("Capacitación y Talento: ").bold = True
        rec3.add_run("Impulsar talleres en servicio al cliente y cumplimiento de imagen corporativa para los colaboradores de ventanilla.")
        
        rec4 = doc.add_paragraph(style='List Bullet')
        rec4.add_run("Cierre de Cobertura: ").bold = True
        rec4.add_run(f"Priorizar la programación de visitas para las {tot_novis} agencias restantes en el cronograma del mes.")

    else:
        doc.add_paragraph("No se registran visitas completadas para el filtro seleccionado.")

    # Guardar en memoria
    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# -----------------------------------------------------------------------------
# BARRA LATERAL
# -----------------------------------------------------------------------------
st.sidebar.title("📌 Control de Gestión")

coordinadores = ["Todos"] + [str(x) for x in df_base[col_coord].dropna().unique() if str(x).strip() != ""]
idx_coord = coordinadores.index(st.session_state['coordinador_activo']) if st.session_state['coordinador_activo'] in coordinadores else 0

coord_seleccionado = st.sidebar.selectbox("Filtrar por Coordinador:", coordinadores, index=idx_coord)

if coord_seleccionado != st.session_state['coordinador_activo']:
    st.session_state['coordinador_activo'] = coord_seleccionado
    st.rerun()

if st.sidebar.button("🔄 Mostrar Todos los Datos"):
    st.session_state['coordinador_activo'] = "Todos"
    st.rerun()

st.sidebar.markdown("---")

# DESCARGA DE INFORME EN WORD (.DOCX)
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
else:
    st.sidebar.info("Registra visitas para habilitar la descarga del informe.")

st.sidebar.markdown("---")

# FORMULARIO DE REGISTRO EN EXCEL
with st.sidebar.expander("➕ Alimentar / Registrar Visita", expanded=False):
    with st.form("form_visita"):
        agencia_sel = st.selectbox("Agencia:", df_base[col_agencia].unique())
        fecha_in = st.date_input("Fecha de Visita")
        
        calif_c = st.number_input("Calidad (1.0 - 3.0)", min_value=1.0, max_value=3.0, value=3.0, step=0.1)
        calif_p = st.number_input("Procesos Transaccionales", min_value=1.0, max_value=3.0, value=3.0, step=0.1)
        calif_m = st.number_input("Manejo Efectivo Control", min_value=1.0, max_value=3.0, value=3.0, step=0.1)
        calif_t = st.number_input("Talento y Cultura", min_value=1.0, max_value=3.0, value=3.0, step=0.1)
        
        obs_cal = st.text_area("Obs. Calidad")
        obs_ope = st.text_area("Obs. Proceso Operativo")
        
        btn_guardar = st.form_submit_button("Guardar en Excel")
        
        if btn_guardar:
            idx = df_base[df_base[col_agencia] == agencia_sel].index
            if not idx.empty:
                i = idx[0]
                if col_fecha: df_base.loc[i, col_fecha] = str(fecha_in)
                
                for c, v in zip(["Calidad", "ProcesosTransaccionales", "ManejoEfectivoControl", "TalentoCultura"], 
                                [calif_c, calif_p, calif_m, calif_t]):
                    if c in df_base.columns:
                        df_base.loc[i, c] = v
                
                if "ObsCalidad" in df_base.columns: df_base.loc[i, "ObsCalidad"] = obs_cal
                if "ObsProcesoOperativo" in df_base.columns: df_base.loc[i, "ObsProcesoOperativo"] = obs_ope
                
                cols_originales = [c for c in df_base.columns if c not in ["Visitada", "PromedioGral"]]
                df_base[cols_originales].to_excel(ARCHIVO_EXCEL, index=False)
                
                st.success("¡Datos guardados correctamente!")
                st.rerun()

# Filtrado de DataFrame
if st.session_state['coordinador_activo'] != "Todos":
    df_filtrado = df_base[df_base[col_coord].astype(str) == st.session_state['coordinador_activo']]
else:
    df_filtrado = df_base.copy()

# -----------------------------------------------------------------------------
# DASHBOARD
# -----------------------------------------------------------------------------
pestana1, pestana2, pestana3 = st.tabs(["📊 Resumen de Gestión", "📝 Observaciones del Coordinador", "📄 Vista Previa del Informe Word"])

with pestana1:
    st.title("Resumen de Gestión de Agencias")

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
        st.subheader("Resumen por Coordinador")
        resumen_coord = df_base.groupby(col_coord).agg(
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
            barmode='stack', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#1E293B"), height=140, showlegend=False, margin=dict(l=0, r=0, t=10, b=0)
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

    st.markdown("---")
    st.subheader("Detalle de Respuestas por Agencia")
    cols_mostrar = [col_agencia, col_fecha] + cols_eval + ["PromedioGral"]
    cols_existentes = [c for c in cols_mostrar if c in df_filtrado.columns]
    
    st.dataframe(
        df_filtrado[df_filtrado["Visitada"] == "Sí"][cols_existentes],
        use_container_width=True, hide_index=True
    )

with pestana2:
    st.title("Observaciones del Coordinador")
    st.markdown("---")
    
    df_obs = df_filtrado[df_filtrado["Visitada"] == "Sí"]
    cols_obs = [c for c in [col_agencia, col_region, col_coord, "ObsCalidad", "ObsProcesoOperativo"] if c in df_filtrado.columns]
    
    if not df_obs.empty:
        st.dataframe(df_obs[cols_obs], use_container_width=True, hide_index=True)
    else:
        st.info("No hay observaciones registradas para la selección actual.")

with pestana3:
    st.title("📄 Estructura del Informe Ejecutivo (.docx)")
    st.markdown("---")
    st.markdown("""
    ### Secciones incluidas en el archivo Word generado:
    1. **Encabezado y Titular Institucional**
    2. **Resumen Ejecutivo:**
       - Métricas globales de cobertura (% Cobertura, Agencias Auditadas vs Pendientes).
       - Tabla resumen resaltada con estilo corporativo.
    3. **Evaluación Detallada:**
       - Tabla formateada con semaforización de calificaciones (Crítico, Aceptable, Excelente).
    4. **Observaciones y Hallazgos:**
       - Agrupados por agencia, desglosando Calidad y Proceso Operativo.
    5. **Recomendaciones Estratégicas:**
       - Plan de remediación e intervención operativa.
    """)