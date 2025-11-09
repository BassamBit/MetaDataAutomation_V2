import io, re, numpy as np, pandas as pd, streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

st.set_page_config(page_title="Metadata Automation", page_icon="📊", layout="centered")
st.title("📘 Metadata Automation Tool")
st.caption("Generate Dictionary and TOC PowerPoint files from your Excel dataset.")

# --------------------------------------------------------------------
# Sidebar settings
# --------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Settings")
    mode = st.radio("Select Mode", ["Dictionary", "TOC"])
    st.markdown("---")
    st.caption("Upload your Excel and PowerPoint template below.")

# --------------------------------------------------------------------
# File uploads
# --------------------------------------------------------------------
pptx_file = st.file_uploader("📎 Upload PowerPoint Template", type=["pptx"])
data_file = st.file_uploader("📎 Upload Excel File", type=["xlsx", "xls", "csv"])

sheet_name = None
if data_file is not None and data_file.name.lower().endswith((".xlsx", ".xls")):
    try:
        xls = pd.ExcelFile(data_file)
        sheet_name = st.selectbox("Select sheet", options=xls.sheet_names, index=0)
    except Exception:
        sheet_name = None

# --------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------
COLOR_DARK_BLUE = RGBColor(31, 78, 121)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_GREEN = RGBColor(82, 158, 69)
COLOR_ORANGE = RGBColor(237, 125, 49)

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
    bodyPr.set(qn('a:rtlCol'), '1' if rtl else '0')
    for p in cell.text_frame.paragraphs:
        pPr = p._p.get_or_add_pPr()
        pPr.set(qn('a:rtl'), '1' if rtl else '0')

def combine_name_definition(name_series, def_series, add_colon=True):
    name_series = name_series.fillna("").astype(str)
    def_series = def_series.fillna("").astype(str)
    result = []
    for n, d in zip(name_series, def_series):
        n = n.strip()
        d = d.strip()
        if not n and not d:
            result.append("")
        elif add_colon:
            result.append(f"{n}:\n{d}" if d else n)
        else:
            result.append(f"{n}\n{d}" if d else n)
    return result

# --------------------------------------------------------------------
# Execution button
# --------------------------------------------------------------------
run = st.button("🚀 Generate", type="primary", use_container_width=True)

if run:
    if data_file is None or pptx_file is None:
        st.error("Please upload both the Excel file and PowerPoint template.")
        st.stop()

    # Load Excel
    if data_file.name.lower().endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file, sheet_name=sheet_name)

    st.success(f"Loaded {df.shape[0]} rows × {df.shape[1]} columns.")

    prs = Presentation(io.BytesIO(pptx_file.read()))

    if mode == "Dictionary":
        st.subheader("📗 Dictionary Settings")

        owner_ar = st.text_input("Data Owner (Arabic)", value="الإدارة القانونية")
        owner_en = st.text_input("Data Owner (English)", value="Legal Department")

        # Build columns
        df["col1"] = combine_name_definition(df["Name in arabic"], df["Definition in arabic"])
        df["col2"] = owner_ar
        df["col3"] = df["Classification"].replace({
            "Public": "عام", "Restricted": "مقيد"
        })
        df["col4"] = df["Code"]
        df["col5"] = owner_en
        df["col6"] = df["Classification"]
        df["col7"] = ""
        df["col8"] = combine_name_definition(df["Name"], df["Definition"])
        df["col9"] = ""

        df_out = df[["col1","col2","col3","col4","col5","col6","col7","col8","col9"]]

        # chunk size (9 per slide)
        ROWS_PER_SLIDE = 9
        chunks = [df_out.iloc[i:i+ROWS_PER_SLIDE] for i in range(0, len(df_out), ROWS_PER_SLIDE)]
        needed_slides = len(chunks)
        available_slides = len(prs.slides)

        if available_slides < needed_slides:
            st.warning(f"⚠️ Not enough slides ({available_slides} < {needed_slides}). Only filling available ones.")
            chunks = chunks[:available_slides]

        # formatters simplified
        def format_cell(cell, text, color=COLOR_BLACK, bold=False, align=PP_ALIGN.CENTER, rtl=False, size=11):
            tf = _clear_cell(cell)
            p = tf.paragraphs[0]
            _set_alignment(cell, align, True)
            _add_run(p, text, size, color, bold)
            if rtl:
                set_cell_rtl(cell, True)

        # Fill slides
        for i, chunk in enumerate(chunks):
            table = None
            for sh in prs.slides[i].shapes:
                if sh.has_table:
                    table = sh.table
                    break
            if table is None:
                continue

            for r in range(1, len(table.rows)):
                for c in range(len(table.columns)):
                    table.cell(r, c).text = ""

            for r in range(len(chunk)):
                row = chunk.iloc[r]
                # Arabic term/def
                format_cell(table.cell(r+1,0), row["col1"], rtl=True, align=PP_ALIGN.RIGHT)
                format_cell(table.cell(r+1,1), row["col2"])
                # classification AR colors
                color3 = COLOR_GREEN if "عام" in row["col3"] else COLOR_ORANGE if "مقيد" in row["col3"] else COLOR_BLACK
                format_cell(table.cell(r+1,2), row["col3"], color3)
                format_cell(table.cell(r+1,3), row["col4"])
                format_cell(table.cell(r+1,4), row["col5"])
                color6 = COLOR_GREEN if "Public" in str(row["col6"]) else COLOR_ORANGE if "Restricted" in str(row["col6"]) else COLOR_BLACK
                format_cell(table.cell(r+1,5), row["col6"], color6)
                format_cell(table.cell(r+1,6), row["col7"])
                format_cell(table.cell(r+1,7), row["col8"], align=PP_ALIGN.LEFT)
                format_cell(table.cell(r+1,8), row["col9"])

        out_buf = io.BytesIO()
        prs.save(out_buf)
        st.download_button("⬇️ Download Dictionary", data=out_buf.getvalue(),
            file_name="dictionary_output.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    elif mode == "TOC":
        st.subheader("📘 TOC Settings")

        df["col1"] = df["Code"]
        df["col2"] = df["Name"]
        df["col3"] = ""
        df["col4"] = df["Name in arabic"]

        toc_matrix = list(zip(df["col1"], df["col2"], df["col3"], df["col4"]))

        ROWS_PER_TABLE = 15
        COLS_PER_TABLE = 4
        rows_per_slide = ROWS_PER_TABLE * 2
        total_needed = len(toc_matrix)

        cur_idx = 0
        for slide in prs.slides:
            tables = [sh.table for sh in slide.shapes if sh.has_table]
            if len(tables) < 2:
                continue
            right, left = tables[0], tables[1]

            for t in [right, left]:
                for r in range(1, len(t.rows)):
                    for c in range(len(t.columns)):
                        t.cell(r,c).text = ""

            for table in [right, left]:
                for r in range(ROWS_PER_TABLE):
                    if cur_idx >= total_needed:
                        break
                    vals = toc_matrix[cur_idx]
                    for c in range(COLS_PER_TABLE):
                        table.cell(r,c).text = "" if vals[c] is None else str(vals[c])
                        _set_alignment(table.cell(r,c), PP_ALIGN.CENTER, True)
                    cur_idx += 1
                if cur_idx >= total_needed:
                    break
            if cur_idx >= total_needed:
                break

        if cur_idx < total_needed:
            st.warning(f"⚠️ Not enough slides: only {cur_idx} of {total_needed} rows written.")

        out_buf = io.BytesIO()
        prs.save(out_buf)
        st.download_button("⬇️ Download TOC", data=out_buf.getvalue(),
            file_name="toc_output.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

st.markdown("---")
