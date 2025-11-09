import io
import re
import numpy as np
import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# === Streamlit basic config ===
st.set_page_config(page_title="Metadata Automation", page_icon="📊", layout="centered")

st.title("📊 Metadata → PowerPoint Generator")
st.caption("Upload your Excel file and PowerPoint templates to automatically generate Dictionary and TOC slides.")

# --- Sidebar selection ---
tool = st.sidebar.radio("Select a tool:", ["📗 Dictionary Generator", "📑 TOC Generator"])

# --- File upload ---
pptx_template = st.file_uploader("Upload PowerPoint Template (.pptx)", type=["pptx"])
data_file = st.file_uploader("Upload Excel File (.xlsx or .xls)", type=["xlsx", "xls"])

# --- Load Excel and show sheet dropdown ---
sheet_name = None
if data_file:
    try:
        xls = pd.ExcelFile(data_file)
        sheet_name = st.selectbox("Select Worksheet", xls.sheet_names)
        df = pd.read_excel(data_file, sheet_name=sheet_name)
        st.success(f"✅ Loaded data: {df.shape[0]} rows × {df.shape[1]} columns")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()
else:
    st.info("Please upload an Excel file to continue.")
    st.stop()

# --- colors ---
COLOR_DARK_BLUE = RGBColor(31, 78, 121)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_GREEN = RGBColor(82, 158, 69)
COLOR_ORANGE = RGBColor(237, 125, 49)

# === Common helpers ===
def _clear_cell(cell):
    tf = cell.text_frame
    tf.clear()
    return tf

def _set_alignment(cell, halign=None, valign_middle=False):
    if valign_middle:
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if cell.text_frame.paragraphs and halign is not None:
        cell.text_frame.paragraphs[0].alignment = halign

def _add_run(p, text, size_pt=11, color=COLOR_BLACK, bold=False):
    run = p.add_run()
    run.text = "" if text is None else str(text)
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color
    run.font.bold = bold

def get_first_table(slide):
    for sh in slide.shapes:
        if sh.has_table:
            return sh.table
    raise RuntimeError("No table found in this slide.")

def set_cell_rtl(cell, rtl=True):
    txBody = cell._tc.txBody
    bodyPr = txBody.bodyPr
    bodyPr.set(qn("a:rtlCol"), "1" if rtl else "0")
    for p in cell.text_frame.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set(qn("a:rtl"), "1" if rtl else "0")

# === Page 1: Dictionary Generator ===
if tool == "📗 Dictionary Generator":
    st.header("📗 Dictionary Generator")

    colnames = df.columns.tolist()
    st.markdown("#### Select columns from your Excel file:")
    code_col = st.selectbox("Code Column", colnames)
    en_name_col = st.selectbox("English Name Column", colnames)
    en_def_col = st.selectbox("English Definition Column", colnames)
    classification_col = st.selectbox("Classification Column", colnames)
    ar_name_col = st.selectbox("Arabic Name Column", colnames)
    ar_def_col = st.selectbox("Arabic Definition Column", colnames)

    st.markdown("#### Enter fixed values:")
    owner_en = st.text_input("Data Owner (English)", "Legal Department")
    owner_ar = st.text_input("Data Owner (Arabic)", "الإدارة القانونية")

    if st.button("Generate Dictionary ✅", use_container_width=True):
        try:
            prs = Presentation(io.BytesIO(pptx_template.read()))
            ROWS_PER_SLIDE = 9

            chunks = [df.iloc[i:i+ROWS_PER_SLIDE] for i in range(0, len(df), ROWS_PER_SLIDE)]
            if len(prs.slides) < len(chunks):
                st.warning(f"⚠ Template has {len(prs.slides)} slides but you need {len(chunks)}.")

            for idx, chunk in enumerate(chunks):
                slide = prs.slides[idx]
                table = get_first_table(slide)
                for r in range(1, len(table.rows)):
                    for c in range(len(table.columns)):
                        table.cell(r, c).text = ""

                for r, (_, row) in enumerate(chunk.iterrows(), start=1):
                    if r >= len(table.rows):
                        break
                    vals = [
                        row[code_col],  # col1
                        f"{row[en_name_col]}:\n{row[en_def_col]}",  # col2
                        owner_en,  # col3
                        row[classification_col],  # col4
                        "",  # col5 (Personal Data)
                        "",  # col6 (Sensitive)
                        row[classification_col].replace("Public", "عام").replace("Restricted", "مقيد"),  # col7 Arabic classification
                        owner_ar,  # col8
                        f"{row[ar_name_col]}:\n{row[ar_def_col]}"  # col9 Arabic Name+Def
                    ]
                    for c, val in enumerate(vals):
                        tf = _clear_cell(table.cell(r, c))
                        p = tf.paragraphs[0]
                        _add_run(p, val, size_pt=11, color=COLOR_BLACK, bold=False)

            out_buf = io.BytesIO()
            prs.save(out_buf)
            out_buf.seek(0)
            st.success("✅ Dictionary file created successfully!")
            st.download_button("⬇ Download Dictionary (PPTX)", data=out_buf,
                               file_name="dictionary_output.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"Error: {e}")

# === Page 2: TOC Generator ===
if tool == "📑 TOC Generator":
    st.header("📑 TOC Generator")

    colnames = df.columns.tolist()
    code_col = st.selectbox("Code Column", colnames)
    name_en_col = st.selectbox("English Name Column", colnames)
    name_ar_col = st.selectbox("Arabic Name Column", colnames)

    if st.button("Generate TOC ✅", use_container_width=True):
        try:
            prs = Presentation(io.BytesIO(pptx_template.read()))
            ROWS_PER_TABLE = 15
            COLS_PER_TABLE = 4

            toc_matrix = []
            for _, row in df.iterrows():
                toc_matrix.append([row[code_col], row[name_en_col], "", row[name_ar_col]])

            def get_tables_sorted_by_x(slide):
                tables = []
                for sh in slide.shapes:
                    if sh.has_table:
                        tables.append((sh.left, sh.table))
                tables.sort(key=lambda t: t[0], reverse=True)
                return [t[1] for t in tables]

            def clear_table_data(table):
                max_rows = len(table.rows) - 1
                max_cols = len(table.columns)
                for r in range(1, max_rows+1):
                    for c in range(max_cols):
                        table.cell(r, c).text = ""

            cur_idx = 0
            for slide in prs.slides:
                tables = get_tables_sorted_by_x(slide)
                if len(tables) < 2:
                    continue
                right_table, left_table = tables[0], tables[1]
                clear_table_data(right_table)
                clear_table_data(left_table)

                for tbl in [right_table, left_table]:
                    for r in range(1, ROWS_PER_TABLE+1):
                        if cur_idx >= len(toc_matrix):
                            break
                        vals = toc_matrix[cur_idx]
                        for c in range(COLS_PER_TABLE):
                            tf = _clear_cell(tbl.cell(r, c))
                            p = tf.paragraphs[0]
                            _add_run(p, vals[c], size_pt=11, color=COLOR_BLACK, bold=False)
                        cur_idx += 1
                if cur_idx >= len(toc_matrix):
                    break

            if cur_idx < len(toc_matrix):
                st.warning(f"⚠ Template slides not enough: only {cur_idx} of {len(toc_matrix)} rows written.")

            out_buf = io.BytesIO()
            prs.save(out_buf)
            out_buf.seek(0)
            st.success("✅ TOC file created successfully!")
            st.download_button("⬇ Download TOC (PPTX)", data=out_buf,
                               file_name="toc_output.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"Error: {e}")
