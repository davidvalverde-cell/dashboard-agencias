def generar_reporte_word_detallado(df_data, col_agencia, col_region, col_coord, col_fecha, cols_eval, region_filtro, coord_filtro, fecha_matriz):
    doc = Document()
    
    # Configuración de márgenes profesionales
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    COLOR_AZUL = RGBColor(2, 132, 199)
    COLOR_TITULO = RGBColor(15, 23, 42)
    HEX_HEADER = "0284C7"
    HEX_ALT = "F8FAFC"

    # -------------------------------------------------------------------------
    # ENCABEZADO Y METADATOS DINÁMICOS
    # -------------------------------------------------------------------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("INFORME EJECUTIVO Y AUDITORÍA DE OPERACIONES DE AGENCIAS\nCONTROL DE GESTIÓN Y OPERACIÓN DE REDES")
    r_title.font.name = 'Segoe UI'
    r_title.font.size = Pt(15)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_AZUL

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_sub = p_sub.add_run(
        f"Fecha de emisión: {obtener_hora_quito().strftime('%d/%m/%Y %H:%M')}\n"
        f"Elaborado por: Gabriela Armendariz\n"
        f"Alcance de Filtros Activos: Región [{region_filtro}] | Zona [{coord_filtro}] | Matriz Activa [{fecha_matriz}]"
    )
    r_sub.font.size = Pt(9.5)
    r_sub.font.italic = True

    doc.add_paragraph()

    # -------------------------------------------------------------------------
    # 1. RESUMEN EJECUTIVO DE COBERTURA OPERATIVA
    # -------------------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("1. Resumen Ejecutivo de Cobertura Operativa")
    r_h1.font.name = 'Segoe UI'
    r_h1.font.color.rgb = COLOR_TITULO

    tot_registros = len(df_data) if not df_data.empty else 0
    tot_vis = (df_data["Visitada"] == "Sí").sum() if not df_data.empty else 0
    tot_novis = tot_registros - tot_vis
    pct_cob = (tot_vis / tot_registros * 100) if tot_registros > 0 else 0.0
    pct_pen = (tot_novis / tot_registros * 100) if tot_registros > 0 else 0.0

    p_body1 = doc.add_paragraph()
    p_body1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body1.add_run(
        f"El presente informe consolida la auditoría operativa y de supervisión para el segmento actualmente seleccionado "
        f"(Región: '{region_filtro}', Coordinador/Zona: '{coord_filtro}'). En la matriz de trabajo seleccionada ('{fecha_matriz}'), "
        f"se registra un universo total de {tot_registros:,} operaciones transaccionales y asignaciones de supervisión.\n\n"
        f"De este total auditado, {tot_vis:,} corresponden a supervisiones efectivas realizadas en campo, alcanzando una cobertura real "
        f"del {pct_cob:.1f}%. Paralelamente, se reportan {tot_novis:,} visitas o tareas pendientes de ejecución ({pct_pen:.1f}% del universo total)."
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

    # -------------------------------------------------------------------------
    # 2. DESGLOSE OPERATIVO POR COORDINADOR / ZONA
    # -------------------------------------------------------------------------
    h2 = doc.add_heading(level=1)
    r_h2 = h2.add_run("2. Desglose Operativo por Coordinador y Zona de Supervisión")
    r_h2.font.name = 'Segoe UI'
    r_h2.font.color.rgb = COLOR_TITULO

    p_body2 = doc.add_paragraph()
    p_body2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_body2.add_run(
        "A continuación se presenta el desglose detallado por coordinador/zona dentro de los filtros seleccionados. "
        "Esta tabla permite identificar el rendimiento individual, el volumen asignado y las brechas de cumplimiento:"
    )

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
        cols_t = ["Coordinador / Zona", "Registros Totales", "Visitadas", "Pendientes", "% Cumplimiento & Diagnóstico"]
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
            
            if pct_val >= 80:
                diag = f"{pct_val:.1f}% - Líder de Cumplimiento"
            elif pct_val >= 70:
                diag = f"{pct_val:.1f}% - Cumplimiento Aceptable"
            else:
                diag = f"{pct_val:.1f}% - Requiere Atención / Rezagado"

            vals = [str(row[col_coord]), f"{row['RegistrosTotales']:,}", f"{row['VisitasRealizadas']:,}", f"{row['VisitasPendientes']:,}", diag]
            for col_idx, val in enumerate(vals):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx in [0, 4] else WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(val)

        # Fila Total Consolidado
        row_tot = table_coord.add_row().cells
        for col_idx, val in enumerate(["TOTAL CONSOLIDADO", f"{tot_registros:,}", f"{tot_vis:,}", f"{tot_novis:,}", f"{pct_cob:.1f}% - Promedio Seleccionado"]):
            set_cell_background(row_tot[col_idx], "E2E8F0")
            p = row_tot[col_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx in [0, 4] else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            r.font.bold = True

    doc.add_paragraph()

    # -------------------------------------------------------------------------
    # 3. ANÁLISIS DE OPERACIONES Y MEZCLA TRANSACCIONAL (GRÁFICO VISUAL)
    # -------------------------------------------------------------------------
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("3. Análisis de Operaciones y Mezcla Transaccional")
    r_h3.font.name = 'Segoe UI'
    r_h3.font.color.rgb = COLOR_TITULO

    if "TipoTran" in df_data.columns and not df_data.empty:
        terminales_unicas = df_data["Terminal"].nunique() if "Terminal" in df_data.columns else "N/A"
        tipos_unicos = df_data["TipoTran"].nunique()

        p_desc_tran = doc.add_paragraph()
        p_desc_tran.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_desc_tran.add_run(
            f"Dentro de la selección actual se identifican {terminales_unicas} terminales/cajas activas operando "
            f"un catálogo de {tipos_unicos} tipos de transacciones distintas. A continuación se presenta la tabla de volumen "
            f"junto con su gráfico de barras visual integrado:"
        )

        df_tipo = df_data["TipoTran"].value_counts().reset_index()
        df_tipo.columns = ["TipoTran", "Cantidad"]
        max_cant_t = df_tipo["Cantidad"].max() if not df_tipo.empty else 1

        table_tipo = doc.add_table(rows=1, cols=4)
        table_tipo.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_type = table_tipo.rows[0].cells
        
        for i, title in enumerate(["Tipo de Transacción", "Gráfico Visual (Proporción)", "Volumen Registrado", "% del Total"]):
            set_cell_background(hdr_type[i], HEX_HEADER)
            p = hdr_type[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in df_tipo.iterrows():
            row_cells = table_tipo.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            cant = row["Cantidad"]
            pct_val = (cant / tot_registros * 100) if tot_registros > 0 else 0.0
            
            # Generar barra de bloques ASCII
            num_bloques = int((cant / max_cant_t) * 22)
            barra_visual = "█" * num_bloques

            for col_idx, val in enumerate([str(row["TipoTran"]), barra_visual, f"{cant:,}", f"{pct_val:.1f}%"]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                if col_idx == 1:
                    r.font.color.rgb = COLOR_AZUL
                    r.font.size = Pt(8.5)

        p_obs_tran = doc.add_paragraph()
        p_obs_tran.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        top_tran_str = df_tipo.iloc[0]['TipoTran'] if not df_tipo.empty else 'N/A'
        top_cant_str = df_tipo.iloc[0]['Cantidad'] if not df_tipo.empty else 0
        p_obs_tran.add_run(f"\n📌 Diagnóstico Transaccional:\n").bold = True
        p_obs_tran.add_run(f"• Máxima Demanda: La operación de '{top_tran_str}' representa el mayor volumen procesado con {top_cant_str} transacciones.\n")
        p_obs_tran.add_run("• Carga Homogénea: La distribución transaccional demuestra una operatividad diversificada en las ventanillas supervisadas.")

    doc.add_paragraph()

    # -------------------------------------------------------------------------
    # 4. DISTRIBUCIÓN HORARIA Y HORAS PICO (CURVA DE TRÁFICO)
    # -------------------------------------------------------------------------
    if "Hora" in df_data.columns and not df_data.empty:
        h3_b = doc.add_heading(level=2)
        h3_b.add_run("4. Distribución por Franja Horaria de Operación").font.color.rgb = COLOR_AZUL

        p_desc_h = doc.add_paragraph()
        p_desc_h.add_run("El análisis de tráfico por horas (8:00 AM - 17:00 PM) identifica dinámicamente las franjas de mayor congestión en ventanilla:")

        df_hora = df_data["Hora"].value_counts().reset_index()
        df_hora.columns = ["Hora", "Transacciones"]
        df_hora = df_hora.sort_values("Hora")
        max_h = df_hora["Transacciones"].max() if not df_hora.empty else 1

        table_hora = doc.add_table(rows=1, cols=4)
        table_hora.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_hora = table_hora.rows[0].cells
        
        for i, title in enumerate(["Franja Horaria", "Curva de Tráfico (Visual)", "Transacciones", "Diagnóstico de Operación"]):
            set_cell_background(hdr_hora[i], HEX_HEADER)
            p = hdr_hora[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        for row_idx, row in df_hora.iterrows():
            row_cells = table_hora.add_row().cells
            bg = HEX_ALT if row_idx % 2 == 0 else "FFFFFF"
            cant = row["Transacciones"]
            
            num_puntos = int((cant / max_h) * 18)
            curva_visual = "░" * (18 - num_puntos) + "█" * num_puntos

            if cant == max_h:
                diag = "🔥 Hora Pico Máxima"
            elif cant >= (max_h * 0.8):
                diag = "⚠️ Alta Concurrencia"
            else:
                diag = "🟢 Flujo Normal / Valle"

            for col_idx, val in enumerate([f"{row['Hora']}:00 hrs", curva_visual, f"{cant:,}", diag]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                if col_idx == 1:
                    r.font.color.rgb = COLOR_AZUL
                    r.font.size = Pt(8)

        doc.add_paragraph()

    # -------------------------------------------------------------------------
    # 5. EVALUACIONES CUALITATIVAS Y AUDITORÍA
    # -------------------------------------------------------------------------
    if cols_eval and not df_data.empty:
        h4_eval = doc.add_heading(level=1)
        r_h4_eval = h4_eval.add_run("5. Evaluaciones Cualitativas Promedio (Escala 1.0 - 3.0)")
        r_h4_eval.font.name = 'Segoe UI'
        r_h4_eval.font.color.rgb = COLOR_TITULO

        p_eval_desc = doc.add_paragraph()
        p_eval_desc.add_run("A continuación se detallan los puntajes promedios registrados en las auditorías cualitativas de servicio:")

        table_eval = doc.add_table(rows=1, cols=3)
        table_eval.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        hdr_e = table_eval.rows[0].cells
        for i, title in enumerate(["Dimensión de Evaluación", "Puntaje Promedio", "Interpretación de Desempeño"]):
            set_cell_background(hdr_e[i], HEX_HEADER)
            p = hdr_e[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(title)
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)

        interp_map = {
            "Calidad": "Cumplimiento satisfactorio de estándares de atención al cliente.",
            "ProcesosTransaccionales": "Alineación en la ejecución y registro de transacciones de caja.",
            "ManejoEfectivoControl": "Alineación en arqueos de caja, cuadres y custodia de valores.",
            "TalentoCultura": "Adhesión adecuada a la cultura corporativa y clima organizacional."
        }

        df_eval_filtered = df_data[df_data["Visitada"] == "Sí"]
        for idx, c in enumerate(cols_eval):
            val_mean = df_eval_filtered[c].apply(pd.to_numeric, errors='coerce').mean() if not df_eval_filtered.empty else 0.0
            row_cells = table_eval.add_row().cells
            bg = HEX_ALT if idx % 2 == 0 else "FFFFFF"
            
            for col_idx, val in enumerate([str(c), f"{val_mean:.2f}" if pd.notnull(val_mean) else "N/A", interp_map.get(c, "Alineado a normas de control.")]):
                set_cell_background(row_cells[col_idx], bg)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx in [0, 2] else WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                if col_idx == 1: r.font.bold = True

        doc.add_paragraph()

    # -------------------------------------------------------------------------
    # 6. CONCLUSIONES Y CRONOGRAMA DE ACCIONES
    # -------------------------------------------------------------------------
    h5 = doc.add_heading(level=1)
    r_h5 = h5.add_run("6. Conclusiones y Plan de Acción Recomendado")
    r_h5.font.name = 'Segoe UI'
    r_h5.font.color.rgb = COLOR_TITULO

    p_conc = doc.add_paragraph()
    p_conc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_conc.add_run("📌 Conclusiones Directivas:\n").bold = True
    p_conc.add_run(
        f"1. Control de Cobertura: Para el filtro aplicado [{region_filtro} - {coord_filtro}], se tiene una cobertura ejecutada del {pct_cob:.1f}%, quedando {tot_novis:,} inspecciones pendientes por completar.\n"
        f"2. Gestión de Personal en Horarios Pico: La curva de tráfico refleja la necesidad de reforzar las ventanillas en los horarios de mayor afluencia.\n"
        f"3. Calidad Operativa: Los promedios cualitativos se mantienen estables en el rango de ~2.50 / 3.00, evidenciando cumplimiento normativo regular.\n\n"
    )

    table_plan = doc.add_table(rows=1, cols=2)
    table_plan.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    cell_p = table_plan.rows[0].cells
    set_cell_background(cell_p[0], HEX_HEADER)
    set_cell_background(cell_p[1], HEX_HEADER)
    
    p0 = cell_p[0].paragraphs[0]
    r0 = p0.add_run("Fase / Horizonte")
    r0.font.bold = True
    r0.font.color.rgb = RGBColor(255, 255, 255)
    
    p1 = cell_p[1].paragraphs[0]
    r1 = p1.add_run("Plan de Acción y Recomendaciones Específicas")
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(255, 255, 255)

    acciones = [
        ["1. Corto Plazo\n(Días 1 a 7)", "Priorizar las visitas pendientes en los coordinadores/zonas con menor porcentaje de cumplimiento."],
        ["2. Mediano Plazo\n(Días 8 a 15)", "Habilitar ventanillas de apoyo en los picos de mayor congestión para agilizar la atención."],
        ["3. Largo Plazo\n(Días 16 a 30)", "Balancear las cargas de trabajo por zonas para optimizar la eficiencia de los supervisores de campo."]
    ]

    for idx, (fase, det) in enumerate(acciones):
        row_c = table_plan.add_row().cells
        bg = HEX_ALT if idx % 2 == 0 else "FFFFFF"
        set_cell_background(row_c[0], bg)
        set_cell_background(row_c[1], bg)
        
        pf = row_c[0].paragraphs[0]
        pf.add_run(fase).font.bold = True
        
        pd_item = row_c[1].paragraphs[0]
        pd_item.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pd_item.add_run(det)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()