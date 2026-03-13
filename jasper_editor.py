import streamlit as st
import io
import pandas as pd
import re
 
# --- Utility: normalize field names ---
def normalize_field_name(s):
    s = str(s)
    s = s.replace("%", "PRC")
    s = s.replace("/", "")
    s = s.replace("\\", "")
    s = s.replace(".", "")
    s = s.replace("-", "")
    s = s.replace("(", "").replace(")", "").replace("（", "").replace("）", "")
    s = s.replace(" ", "")  # Remove all spaces at the end
    return s.upper()
 
def deduplicate_field_names(field_names):
    seen = {}
    result = []
    for name in field_names:
        base = name
        if base in seen:
            seen[base] += 1
            name = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        result.append(name)
    return result
 
st.set_page_config(layout="wide")
st.title("JasperReport Maker")
 
# -----------------------------
# Step 1: Upload Excel/CSV file
# -----------------------------
uploaded_file = st.file_uploader("Upload an Excel file with field data", type=["xlsx", "xls", "csv"])
 
if uploaded_file is not None:
    # --- Detect sheets if Excel ---
    if uploaded_file.name.endswith((".xlsx", ".xls")):
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            st.session_state['excel_file'] = excel_file
            st.session_state['excel_sheets'] = excel_file.sheet_names
 
            if 'selected_sheet' not in st.session_state:
                st.session_state['selected_sheet'] = excel_file.sheet_names[0]
 
            # --- Sheet selection UI ---
            sheet = st.selectbox(
                "Select sheet to preview",
                st.session_state['excel_sheets'],
                index=st.session_state['excel_sheets'].index(st.session_state['selected_sheet']),
                key="sheet_selector"
            )
 
            if sheet != st.session_state['selected_sheet']:
                st.session_state['selected_sheet'] = sheet
 
            # Load selected sheet
            df_raw = excel_file.parse(sheet, dtype=str, header=None)
 
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            df_raw = pd.DataFrame()
    else:
        # CSV
        df_raw = pd.read_csv(uploaded_file, dtype=str, header=None)
 
    st.session_state['excel_df'] = df_raw.copy()
 
    # -------------------------------
    # Step 2: Preview Original Data
    # -------------------------------
    st.subheader("Preview Original Data")
    st.dataframe(st.session_state['excel_df'], use_container_width=True)
    st.markdown("---")
 
    # -------------------------------
    # Step 3: Delete Rows / Columns
    # -------------------------------
    st.subheader("🧹 Delete Rows / Columns")
    rows_to_delete = st.multiselect(
        "Select rows to delete (by index):",
        options=df_raw.index.tolist(),
        default=[]
    )
 
    col_indices = list(range(len(df_raw.columns)))
    col_labels = [f"{i}: {col}" for i, col in enumerate(df_raw.columns)]
    cols_to_delete_idx = st.multiselect(
        "Select columns to delete (by index):",
        options=df_raw.index.tolist(),
        default=[]
    )
    cols_to_delete = [df_raw.columns[i] for i in cols_to_delete_idx]
 
    if st.button("Apply Deletion"):
        df_cleaned = df_raw.drop(index=rows_to_delete, columns=cols_to_delete)
        st.session_state['cleaned_df'] = df_cleaned.copy()
        st.success(f"✅ Deleted {len(rows_to_delete)} rows and {len(cols_to_delete)} columns.")
 
# -----------------------------
# Step 4: Show cleaned data if available
# -----------------------------
df_to_use = st.session_state.get('cleaned_df', st.session_state.get('excel_df', pd.DataFrame()))
 
if 'cleaned_df' in st.session_state:
    st.subheader("Cleaned Data Preview (After Deletion)")
    cleaned_df = st.session_state['cleaned_df']
 
    if cleaned_df.shape[0] > 0:
        raw_names = [normalize_field_name(v) if pd.notna(v) and v != '' else f"COLUMN{i}" for i, v in enumerate(cleaned_df.iloc[0])]
        cleaned_field_names = deduplicate_field_names(raw_names)
        cleaned_df.columns = cleaned_field_names
        st.session_state['cleaned_df'] = cleaned_df.copy()
    st.dataframe(cleaned_df, use_container_width=True)
 
# -----------------------------
# Step 5: Preview Fields (First Row)
# -----------------------------
if not df_to_use.empty:
    st.subheader("Fields name (First Row Preview)")
    first_row = df_to_use.iloc[0]
    raw_names = [normalize_field_name(v) if pd.notna(v) and v != '' else f"COLUMN{i}" for i, v in enumerate(first_row)]
    cols_to_show = deduplicate_field_names(raw_names)
 
    st.markdown("**Fields from First Row of Excel File**")
    header_cols = st.columns([1, 3, 4, 2])
    header_cols[0].markdown("**No.**")
    header_cols[1].markdown("**Field Name**")
    header_cols[2].markdown("**Original Value**")
    header_cols[3].markdown("")
 
    for idx, col_name in enumerate(cols_to_show, 1):
        row_cols = st.columns([1, 3, 4, 2])
        row_cols[0].write(idx)
        row_cols[1].write(col_name)
        val = first_row.iloc[idx - 1]
        if pd.isna(val) or val == '':
            val = "No Value"
        row_cols[2].write(val)
        col_key = f"delete_{idx}_{col_name}"
        if row_cols[3].button("❌", key=col_key):
            new_df = df_to_use.copy().drop(columns=[df_to_use.columns[idx - 1]])
            st.session_state['cleaned_df'] = new_df
            st.rerun()
 
# -----------------------------
# Step 6: Export JRXML and SQL
# -----------------------------
def export_to_jrxml(df, report_name="report"):
    import xml.etree.ElementTree as ET
    import io
 
    ns = "http://jasperreports.sourceforge.net/jasperreports"
    ET.register_namespace('', ns)
    default_col_width = 120
    width_per_col = default_col_width
    left_margin = 20
    right_margin = 20
    total_width = width_per_col * len(df.columns) + left_margin + right_margin  # Fit exactly
    column_width = total_width - left_margin - right_margin
 
    root = ET.Element("jasperReport", attrib={
        "xmlns": ns,
        "name": report_name,
        "pageWidth": str(total_width),
        "pageHeight": "842",
        "columnWidth": str(column_width),
        "leftMargin": str(left_margin),
        "rightMargin": str(right_margin),
        "topMargin": "20",
        "bottomMargin": "20",
        "language": "plsql"
    })
 
    ET.SubElement(root, "property", attrib={
        "name": "com.jaspersoft.studio.data.defaultdataadapter",
        "value": "Multifonds"
    })
 
    query_string = ET.SubElement(root, "queryString", attrib={"language": "plsql"})
    proc_call = f"{{call SP_RPT_{report_name.upper()}($P{{ASATDATE}}, $P{{PRODUCTCODE}}, $P{{ORACLE_REF_CURSOR}})}}"
    query_string.text = proc_call
 
    # Allow duplicate field names by using unique XML field names (e.g., field_0, field_1, ...)
    field_xml_names = [f"field_{i}" for i in range(len(df.columns))]
    field_display_names = list(df.columns)
 
    for xml_name in field_xml_names:
        ET.SubElement(root, "field", attrib={
            "name": xml_name,
            "class": "java.lang.String"
        })

    # --- Bands in JasperReports order (without title band) ---

    # 1. Page Header (custom structure)
    page_header = ET.Element("pageHeader")
    band = ET.SubElement(page_header, "band", attrib={"height": "109"})
    ET.SubElement(band, "property", attrib={"name": "com.jaspersoft.studio.unit.height", "value": "px"})

    # TextField: title (use jrxml file name)
    text_field1 = ET.SubElement(band, "textField")
    report_element1 = ET.SubElement(text_field1, "reportElement", attrib={
        "mode": "Transparent", "x": "0", "y": "50", "width": "802", "height": "14", "backcolor": "#FFFFFF",
        "uuid": "9237bf77-49a3-4a1d-8861-737d431eecdb"
    })
    ET.SubElement(report_element1, "property", attrib={"name": "com.jaspersoft.studio.unit.y", "value": "px"})
    ET.SubElement(report_element1, "property", attrib={"name": "com.jaspersoft.studio.unit.x", "value": "px"})
    box1 = ET.SubElement(text_field1, "box")
    for side in ["topPen", "leftPen", "bottomPen", "rightPen"]:
        ET.SubElement(box1, side, attrib={"lineWidth": "0.0", "lineStyle": "Solid", "lineColor": "#000000"})
    text_element1 = ET.SubElement(text_field1, "textElement", attrib={"textAlignment": "Center", "verticalAlignment": "Middle"})
    ET.SubElement(text_element1, "font", attrib={"fontName": "SansSerif", "size": "9", "isBold": "true"})
    expr1 = ET.SubElement(text_field1, "textFieldExpression")
    expr1.text = f'"{report_name}"'

    # TextField: "FROM " + $P{DateFrom} + " TO " + $P{DateTo}
    text_field2 = ET.SubElement(band, "textField")
    report_element2 = ET.SubElement(text_field2, "reportElement", attrib={
        "mode": "Transparent", "x": "0", "y": "64", "width": "802", "height": "16", "backcolor": "#FFFFFF",
        "uuid": "04d52860-f730-46f8-b954-ea5788f6d82d"
    })
    ET.SubElement(report_element2, "property", attrib={"name": "com.jaspersoft.studio.unit.y", "value": "px"})
    ET.SubElement(report_element2, "property", attrib={"name": "com.jaspersoft.studio.unit.x", "value": "px"})
    box2 = ET.SubElement(text_field2, "box")
    for side in ["topPen", "leftPen", "bottomPen", "rightPen"]:
        ET.SubElement(box2, side, attrib={"lineWidth": "0.0", "lineStyle": "Solid", "lineColor": "#000000"})
    text_element2 = ET.SubElement(text_field2, "textElement", attrib={"textAlignment": "Center"})
    ET.SubElement(text_element2, "font", attrib={"fontName": "SansSerif", "size": "9", "isBold": "true"})
    expr2 = ET.SubElement(text_field2, "textFieldExpression")
    expr2.text = '\"FROM \" + $P{DateFrom} + \" TO \" + $P{DateTo}'

    # TextField: $P{PORTFOLIOCODE}
    text_field3 = ET.SubElement(band, "textField")
    report_element3 = ET.SubElement(text_field3, "reportElement", attrib={
        "mode": "Transparent", "x": "0", "y": "80", "width": "802", "height": "14", "backcolor": "#FFFFFF",
        "uuid": "27c361f5-4435-448c-8da0-a36338170577"
    })
    ET.SubElement(report_element3, "property", attrib={"name": "com.jaspersoft.studio.unit.height", "value": "px"})
    ET.SubElement(report_element3, "property", attrib={"name": "com.jaspersoft.studio.unit.x", "value": "px"})
    box3 = ET.SubElement(text_field3, "box")
    for side in ["topPen", "leftPen", "bottomPen", "rightPen"]:
        ET.SubElement(box3, side, attrib={"lineWidth": "0.0", "lineStyle": "Solid", "lineColor": "#000000"})
    text_element3 = ET.SubElement(text_field3, "textElement", attrib={"verticalAlignment": "Middle"})
    ET.SubElement(text_element3, "font", attrib={"size": "9", "isBold": "true"})
    expr3 = ET.SubElement(text_field3, "textFieldExpression")
    expr3.text = '$P{PORTFOLIOCODE}'

    # TextField: "Base Currency : " + $F{FUNCTIONALCURRENCYISOCODE}
    text_field4 = ET.SubElement(band, "textField")
    report_element4 = ET.SubElement(text_field4, "reportElement", attrib={
        "mode": "Transparent", "x": "0", "y": "94", "width": "802", "height": "15", "backcolor": "#FFFFFF",
        "uuid": "93799b37-aa88-4acd-9fac-b77356d69484"
    })
    ET.SubElement(report_element4, "property", attrib={"name": "com.jaspersoft.studio.unit.height", "value": "px"})
    ET.SubElement(report_element4, "property", attrib={"name": "com.jaspersoft.studio.unit.x", "value": "px"})
    box4 = ET.SubElement(text_field4, "box")
    for side in ["topPen", "leftPen", "bottomPen", "rightPen"]:
        ET.SubElement(box4, side, attrib={"lineWidth": "0.0", "lineStyle": "Solid", "lineColor": "#000000"})
    text_element4 = ET.SubElement(text_field4, "textElement", attrib={"verticalAlignment": "Middle"})
    ET.SubElement(text_element4, "font", attrib={"size": "9", "isBold": "true"})
    expr4 = ET.SubElement(text_field4, "textFieldExpression")
    expr4.text = '"Base Currency : " + $F{FUNCTIONALCURRENCYISOCODE}'

    # Subreport
    subreport = ET.SubElement(band, "subreport")
    report_element_sub = ET.SubElement(subreport, "reportElement", attrib={
        "x": "0", "y": "0", "width": "802", "height": "50", "uuid": "b069e942-d6f7-4735-84e8-a69a086ab4cc"
    })
    ET.SubElement(report_element_sub, "property", attrib={"name": "com.jaspersoft.studio.unit.height", "value": "px"})
    conn_expr = ET.SubElement(subreport, "connectionExpression")
    conn_expr.text = "$P{REPORT_CONNECTION}"
    subreport_expr = ET.SubElement(subreport, "subreportExpression")
    # Use relative path for Jasper preview compatibility
    subreport_expr.text = '"AHAM_Header.jrxml"'

    root.append(page_header)

    # 2. Column Header
    column_header = ET.Element("columnHeader")
    band = ET.SubElement(column_header, "band", attrib={"height": "30"})
    x = 0
    for display_name in field_display_names:
        static_text = ET.SubElement(band, "staticText")
        ET.SubElement(static_text, "reportElement", attrib={"x": str(x), "y": "0", "width": str(width_per_col), "height": "30"})
        ET.SubElement(static_text, "textElement")
        text = ET.SubElement(static_text, "text")
        text.text = str(display_name)
        x += width_per_col
    root.append(column_header)

    # 3. Detail
    detail = ET.Element("detail")
    detail_band = ET.SubElement(detail, "band", attrib={"height": "30"})
    x = 0
    for i, xml_name in enumerate(field_xml_names):
        text_field = ET.SubElement(detail_band, "textField")
        ET.SubElement(text_field, "reportElement", attrib={"x": str(x), "y": "0", "width": str(width_per_col), "height": "30"})
        ET.SubElement(text_field, "textElement")
        text_field_expr = ET.SubElement(text_field, "textFieldExpression")
        text_field_expr.text = f"$F{{{field_display_names[i]}}}"
        x += width_per_col
    root.append(detail)

    # 4. Column Footer
    column_footer = ET.Element("columnFooter")
    band = ET.SubElement(column_footer, "band", attrib={"height": "0"})
    root.append(column_footer)

    # 5. Page Footer
    page_footer = ET.Element("pageFooter")
    band = ET.SubElement(page_footer, "band", attrib={"height": "0"})
    root.append(page_footer)

    # 6. Last Page Footer
    last_page_footer = ET.Element("lastPageFooter")
    band = ET.SubElement(last_page_footer, "band", attrib={"height": "0"})
    root.append(last_page_footer)

    # 7. Summary
    summary = ET.Element("summary")
    band = ET.SubElement(summary, "band", attrib={"height": "0"})
    root.append(summary)

    # 8. No Data
    nodata = ET.Element("noData")
    band = ET.SubElement(nodata, "band", attrib={"height": "0"})
    root.append(nodata)

    buf = io.BytesIO()
    tree = ET.ElementTree(root)
    tree.write(buf, encoding="utf-8", xml_declaration=True)
    buf.seek(0)
    return buf
 
# Export buttons
if not df_to_use.empty:
    # --- Use first row as field names for export, allow duplicates ---
    export_df = df_to_use.copy()
    raw_names = [normalize_field_name(v) if pd.notna(v) and v != '' else f"COLUMN{i}" for i, v in enumerate(export_df.iloc[0])]
    export_field_names = deduplicate_field_names(raw_names)
    export_df.columns = export_field_names
    export_df = export_df.iloc[1:].reset_index(drop=True)  # Remove the first row (header row) from data
 
    default_name = "exported_report"
    file_name = st.text_input("JRXML file name (without .jrxml extension):", value=default_name, key="jrxml_filename")
    jrxml_buf = export_to_jrxml(export_df, report_name=file_name)
 
    proc_name = f"SP_RPT_{file_name.upper()}"
    field_list = ',\n    '.join([str(col) for col in export_df.columns])
    null_field_list = ',\n    '.join([f'NULL AS {col}' for col in export_df.columns])
    create_table_fields = ',\n    '.join([f"{col} VARCHAR2(255)" for col in export_df.columns])
    create_table_sql = f"""CREATE TABLE table_name (
    {create_table_fields}
);
"""
    sql_code = f"""create or replace PROCEDURE {proc_name} (
    p_AsOfDate IN VARCHAR2,
    p_ProductCode IN VARCHAR2,
    p_result OUT SYS_REFCURSOR  
    )IS
    BEGIN
    OPEN p_result FOR

    SELECT
    {field_list}
    vrr.Manager,
    vrr.Path,
    'Data' as TYPE
   
    FROM VW_TABLENAME
    LEFT JOIN VW_REPORTDUMMYROW vrr ON d.PORTFOLIOID = vrr.PORTFOLIOID

    UNION ALL

    SELECT
    {null_field_list}
    vrr.MANAGER AS MANAGER,
    vrr.PATH AS PATH,
    'Header' AS TYPE
    
    FROM VW_REPORTDUMMYROW vrr
    WHERE vrr.portfolioid = v_portfolioid;

    END {proc_name};
"""
    sql_buf = io.BytesIO(sql_code.encode('utf-8'))
 
    tbl_name = f'TBL_RPT_{file_name.upper()}'
    tbl_fields = ',\n    '.join([f'"{col}" VARCHAR2(100 BYTE)' for col in export_df.columns])
    tbl_sql = f'''CREATE TABLE "CUSTOMER"."{tbl_name}" \n(\n    {tbl_fields}\n);\n'''
    tbl_sql_buf = io.BytesIO(tbl_sql.encode('utf-8'))
 
    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button(
            label="Download as JRXML",
            data=jrxml_buf,
            file_name=f"{file_name}.jrxml",
            mime="application/xml"
        )
    with col2:
        st.download_button(
            label="Download SQL Procedure",
            data=sql_buf,
            file_name=f"SP_RPT_{file_name.upper()}.sql",
            mime="text/x-sql"
        )
    with col3:
        st.download_button(
            label="Download Table SQL",
            data=tbl_sql_buf,
            file_name=f"{tbl_name}.sql",
            mime="text/x-sql"
        )
else:
    st.info("Import .xlsx or .csv file to continue.")