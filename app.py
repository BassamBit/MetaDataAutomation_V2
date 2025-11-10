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

# --- File uploads ---
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

# --- Colors ---
COLOR_DARK_BLUE = RGBColor(31, 78, 121)
COLOR_BLACK     = RGBColor(0, 0, 0)
COLOR_GREEN     = RGBColor(82, 158, 69)
COLOR_ORANGE    = RGBColor(237, 125, 49)

# === Helpers (text & table) ===
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

def set_cell_rtl(cell, rtl=True):
    txBody = cell._tc.txBody
    bodyPr = txBody.bodyPr
    bodyPr.set(qn("a:rtlCol"), "1" if rtl else "0")
    for p in cell.text_frame.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set(qn("a:rtl"), "1" if rtl else "0")

def first_line_and_rest(text):
    """Return (first_line, rest_lines_list)."""
    if text is None:
        return "", []
    s = str(text)
    lines = re.split(r"\r?\n", s)
    if not lines:
        return "", []
    return lines[0], lines[1:]

def get_first_table(slide):
    for sh in slide.shapes:
        if sh.has_table:
            return sh.table
    raise RuntimeError("No table found in this slide.")

def get_tables_sorted_by_x(slide):
    tables = []
    for sh in slide.shapes:
        if sh.has_table:
            tables.append((sh.left, sh.table))
    tables.sort(key=lambda t: t[0], reverse=True)  # right-most first
    return [t[1] for t in tables]

def clear_table_data(table):
    max_rows = len(table.rows) - 1
    max_cols = len(table.columns)
    for r in range(1, max_rows+1):
        for c in range(max_cols):
            table.cell(r, c).text = ""

# === Column-based formatters for Dictionary ===
def fmt_ar_combined(cell, value):
    """Arabic Name+Definition: first line 12pt dark-blue bold, rest 11pt black; RTL; right aligned; vertically middle."""
    tf = _clear_cell(cell)
    tf.word_wrap = True
    _set_alignment(cell, halign=PP_ALIGN.RIGHT, valign_middle=True)
    set_cell_rtl(cell, True)
    first, rest = first_line_and_rest(value)
    p = tf.paragraphs[0]
    _add_run(p, first, size_pt=12, color=COLOR_DARK_BLUE, bold=True)
    for line in rest:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.RIGHT
        _add_run(p2, line, size_pt=11, color=COLOR_BLACK, bold=False)

def fmt_en_combined(cell, value):
    """English Name+Definition: first line 12pt dark-blue bold, rest 11pt black; LTR; left aligned; vertically middle."""
    tf = _clear_cell(cell)
    tf.word_wrap = True
    _set_alignment(cell, halign=PP_ALIGN.LEFT, valign_middle=True)
    first, rest = first_line_and_rest(value)
    p = tf.paragraphs[0]
    _add_run(p, first, size_pt=12, color=COLOR_DARK_BLUE, bold=True)
    for line in rest:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.LEFT
        _add_run(p2, line, size_pt=11, color=COLOR_BLACK, bold=False)

def fmt_center(cell, value):
    tf = _clear_cell(cell)
    _set_alignment(cell, halign=PP_ALIGN.CENTER, valign_middle=True)
    p = tf.paragraphs[0]
    _add_run(p, "" if value is None else str(value), size_pt=11, color=COLOR_BLACK, bold=False)

def fmt_class(cell, value, lang="en"):
    """Classification colored:
       - AR: 'عام' -> green, 'مقيد' -> orange
       - EN: 'public' -> green, 'restricted' -> orange
    """
    txt = "" if value is None else str(value).strip()
    low = txt.lower()
    color = COLOR_BLACK
    if lang == "ar":
        if "عام" in txt:
            color = COLOR_GREEN
        elif "مقيد" in txt:
            color = COLOR_ORANGE
    else:
        if "public" in low:
            color = COLOR_GREEN
        elif "restricted" in low:
            color = COLOR_ORANGE
    tf = _clear_cell(cell)
    _set_alignment(cell, halign=PP_ALIGN.CENTER, valign_middle=True)
    p = tf.paragraphs[0]
    _add_run(p, txt, size_pt=11, color=color, bold=False)

# === PAGE: Dictionary ===
if tool == "📗 Dictionary Generator":
    st.header("📗 Dictionary Generator")

    colnames = df.columns.tolist()
    # Dropdowns (keep Excel order)
    code_col          = st.selectbox("Code column", colnames)
    en_name_col       = st.selectbox("English Name column", colnames)
    en_def_col        = st.selectbox("English Definition column", colnames)
    classification_col= st.selectbox("Classification column", colnames)
    ar_name_col       = st.selectbox("Arabic Name column", colnames)
    ar_def_col        = st.selectbox("Arabic Definition column", colnames)

    owner_en = st.text_input("Data Owner (English)", "Legal Department")
    owner_ar = st.text_input("Data Owner (Arabic)", "الإدارة القانونية")

    st.caption("Dictionary expects your template slide table to have 9 columns in this order: "
               "[AR Name+Def, Owner AR, Class AR, Code, Owner EN, Class EN, PersonalData, EN Name+Def, PersonalData(EN)].")

    if st.button("Generate Dictionary ✅", use_container_width=True):
        try:
            prs = Presentation(io.BytesIO(pptx_template.read()))
            ROWS_PER_SLIDE = 9

            # Build combined strings respecting 'Name: \nDefinition'
            def combine(n, d, add_colon=True):
                n = "" if pd.isna(n) else str(n).strip()
                d = "" if pd.isna(d) else str(d).strip()
                if not n and not d:
                    return ""
                if add_colon:
                    return f"{n}:\n{d}" if d else n
                return f"{n}\n{d}" if d else n

            df_local = df.copy()
            en_combined = [combine(n, d, add_colon=True) for n, d in zip(df_local[en_name_col], df_local[en_def_col])]
            ar_combined = [combine(n, d, add_colon=True) for n, d in zip(df_local[ar_name_col], df_local[ar_def_col])]

            # Arabic classification by mapping EN values from same column
            ar_class_series = df_local[classification_col].astype(str).replace({"Public": "عام", "Restricted": "مقيد"})

            # Prepare chunks
            chunks = [df_local.iloc[i:i+ROWS_PER_SLIDE] for i in range(0, len(df_local), ROWS_PER_SLIDE)]
            needed_slides = len(chunks)
            available_slides = len(prs.slides)
            if available_slides < needed_slides:
                st.warning(f"⚠ Template has {available_slides} slides but you need {needed_slides}. Only filling available slides.")
                chunks = chunks[:available_slides]

            for idx, chunk in enumerate(chunks):
                slide = prs.slides[idx]
                table = get_first_table(slide)

                max_rows_data = len(table.rows) - 1  # data rows (excluding header)
                max_cols = len(table.columns)
                if max_cols < 9:
                    st.warning(f"⚠ Slide #{idx+1}: table has {max_cols} columns, expected 9. Filling up to {max_cols}.")
                rows_to_write = min(len(chunk), max_rows_data)

                # clear existing data rows
                for rr in range(1, len(table.rows)):
                    for cc in range(len(table.columns)):
                        table.cell(rr, cc).text = ""

                for i_row in range(rows_to_write):
                    src = chunk.iloc[i_row]
                    # Compose each column value according to your spec & formatting
                    # col0: AR combined
                    fmt_ar_combined(table.cell(i_row+1, 0), ar_combined[chunk.index[i_row]])
                    # col1: Owner AR
                    fmt_center(table.cell(i_row+1, 1), owner_ar)
                    # col2: Class AR (colored)
                    fmt_class(table.cell(i_row+1, 2), ar_class_series.iloc[chunk.index[i_row]], lang="ar")
                    # col3: Code
                    fmt_center(table.cell(i_row+1, 3), src[code_col])
                    # col4: Owner EN
                    fmt_center(table.cell(i_row+1, 4), owner_en)
                    # col5: Class EN (colored)
                    fmt_class(table.cell(i_row+1, 5), src[classification_col], lang="en")
                    # col6: Personal Data (empty)
                    fmt_center(table.cell(i_row+1, 6), "")
                    # col7: EN combined (LTR)
                    fmt_en_combined(table.cell(i_row+1, 7), en_combined[chunk.index[i_row]])
                    # col8: Personal Data EN (empty)
                    fmt_center(table.cell(i_row+1, 8), "")

            out_buf = io.BytesIO()
            prs.save(out_buf)
            out_buf.seek(0)
            st.success("✅ Dictionary file created successfully!")
            st.download_button("⬇ Download Dictionary (PPTX)", data=out_buf,
                               file_name="dictionary_output.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"Error: {e}")

# === PAGE: TOC ===
if tool == "📑 TOC Generator":
    st.header("📑 TOC Generator")

    colnames = df.columns.tolist()
    code_col   = st.selectbox("Code column", colnames, key="toc_code")
    name_en_col= st.selectbox("English Name column", colnames, key="toc_name_en")
    name_ar_col= st.selectbox("Arabic Name column", colnames, key="toc_name_ar")

    st.caption("TOC expects each slide to contain **two tables** (right then left), each with header row + 15 data rows and 4 columns.")

    if st.button("Generate TOC ✅", use_container_width=True):
        try:
            prs = Presentation(io.BytesIO(pptx_template.read()))
            ROWS_PER_TABLE = 15
            COLS_PER_TABLE_EXPECTED = 4

            # Build matrix: [Code, EN Name, PageNumber(blank), AR Name]
            toc_matrix = []
            for _, row in df.iterrows():
                toc_matrix.append([row[code_col], row[name_en_col], "", row[name_ar_col]])

            cur_idx = 0
            total_needed = len(toc_matrix)

            for slide_idx, slide in enumerate(prs.slides, start=1):
                tables = get_tables_sorted_by_x(slide)
                if not tables:
                    continue

                # If template sometimes has only one table, handle gracefully
                # Expected: 2 tables (right, left). If 1 -> fill it; if 2 -> fill both in order.
                target_tables = tables[:2]

                for tbl in target_tables:
                    clear_table_data(tbl)
                    header_rows = 1
                    max_data_rows_tbl = len(tbl.rows) - header_rows
                    if max_data_rows_tbl <= 0:
                        continue
                    max_cols_tbl = len(tbl.columns)

                    rows_to_fill = min(ROWS_PER_TABLE, max_data_rows_tbl)
                    cols_to_fill = min(COLS_PER_TABLE_EXPECTED, max_cols_tbl)

                    for r in range(rows_to_fill):
                        if cur_idx >= total_needed:
                            break
                        vals = toc_matrix[cur_idx]
                        for c in range(cols_to_fill):
                            tf = _clear_cell(tbl.cell(r+header_rows, c))
                            p = tf.paragraphs[0]
                            _set_alignment(tbl.cell(r+header_rows, c), PP_ALIGN.CENTER, True)
                            _add_run(p, vals[c], size_pt=11, color=COLOR_BLACK, bold=False)
                        cur_idx += 1
                    if cur_idx >= total_needed:
                        break
                if cur_idx >= total_needed:
                    break

            if cur_idx < total_needed:
                st.warning(f"⚠ Template slides not enough: wrote {cur_idx} of {total_needed} rows.")

            out_buf = io.BytesIO()
            prs.save(out_buf)
            out_buf.seek(0)
            st.success("✅ TOC file created successfully!")
            st.download_button("⬇ Download TOC (PPTX)", data=out_buf,
                               file_name="toc_output.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            st.error(f"Error: {e}")
