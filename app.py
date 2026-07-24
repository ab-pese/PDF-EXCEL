import streamlit as st
import pandas as pd
import os
import io
import re
import zipfile
import traceback
from fractions import Fraction

import pymupdf  # PyMuPDF
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# ============================================================
# Session State Init
# ============================================================
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.page_files = None      # {page_num: xlsx_bytes}
    st.session_state.page_summary = None    # [(page_num, kept_table_count, skipped_count)]
    st.session_state.basename = None
    st.session_state.preview_tables = None  # {page_num: [df, df, ...]}

st.set_page_config(page_title="PDF Table Extractor", layout="wide")

# ============================================================
# Dimension Conversion Helpers (Architectural <-> Decimal)
# ============================================================
# Recognizes things like: 12'-6"   12'6"   6"   3/4"   12'   3'-0 1/2"
ARCH_RE = re.compile(r"""
    ^\s*
    (?:(?P<feet>\d+(?:\.\d+)?)\s*'\s*-?\s*)?
    (?:(?P<inch_whole>\d+(?:\.\d+)?)\s*)?
    (?:(?P<frac_num>\d+)\s*/\s*(?P<frac_den>\d+)\s*)?
    (?P<inch_mark>")?
    \s*$
""", re.VERBOSE)

# A plain decimal number, e.g. 12.5, -3.25 (used to detect candidates for
# conversion INTO architectural notation). Requires a decimal point so we
# don't accidentally treat plain integer counts/quantities as dimensions.
DECIMAL_ONLY_RE = re.compile(r"^[-+]?\d+\.\d+$")


def parse_architectural(text):
    """Parse an architectural dimension string into decimal feet.
    Returns None if the string doesn't look like an architectural dimension."""
    if not text:
        return None
    m = ARCH_RE.match(text)
    if not m:
        return None

    feet = m.group("feet")
    inch_mark = m.group("inch_mark")
    inch_whole = m.group("inch_whole")
    frac_num = m.group("frac_num")
    frac_den = m.group("frac_den")

    # Require at least one architectural marker (foot mark, inch mark, or a
    # fraction) so plain numbers like "12.5" or "100" are never misread as
    # architectural dimensions.
    if feet is None and inch_mark is None and frac_num is None:
        return None

    feet_val = float(feet) if feet else 0.0
    inch_val = float(inch_whole) if inch_whole else 0.0
    if frac_num and frac_den:
        try:
            inch_val += float(frac_num) / float(frac_den)
        except ZeroDivisionError:
            pass

    return feet_val + inch_val / 12.0


def decimal_to_architectural(value_feet, denom=16):
    """Convert decimal feet into architectural feet-inches-fraction notation,
    rounding the fractional inch to the nearest 1/denom."""
    sign = "-" if value_feet < 0 else ""
    value_feet = abs(value_feet)
    feet = int(value_feet)
    inches_total = (value_feet - feet) * 12
    whole_inches = int(inches_total)
    frac_remainder = inches_total - whole_inches

    frac = Fraction(frac_remainder).limit_denominator(denom)
    if frac.numerator == frac.denominator:
        whole_inches += 1
        frac = Fraction(0)
    if whole_inches == 12:
        feet += 1
        whole_inches = 0

    if frac.numerator == 0:
        inch_str = f"{whole_inches}"
    else:
        inch_str = (
            f"{whole_inches} {frac.numerator}/{frac.denominator}"
            if whole_inches
            else f"{frac.numerator}/{frac.denominator}"
        )

    return f'{sign}{feet}\'-{inch_str}"'


def format_decimal(value_feet, precision=4):
    return f"{round(value_feet, precision):g}"


def convert_cell_text(value, target_unit, frac_denom=16):
    """Convert a single cell's text to the target dimension unit.
    Leaves non-dimension-looking text untouched."""
    if value is None:
        return value
    text_str = str(value).strip()
    if text_str == "" or text_str.lower() == "nan":
        return text_str

    arch_val = parse_architectural(text_str)
    if arch_val is not None:
        if target_unit == "decimal":
            return format_decimal(arch_val)
        return text_str  # already architectural

    if DECIMAL_ONLY_RE.match(text_str) and target_unit == "architectural":
        return decimal_to_architectural(float(text_str), denom=frac_denom)

    return text_str


def convert_dataframe(df, target_unit, frac_denom=16):
    return df.apply(lambda col: col.map(lambda v: convert_cell_text(v, target_unit, frac_denom)))


# ============================================================
# PDF Table Extraction
# ============================================================

def _clean_cell(v):
    if v is None:
        return ""
    return str(v).strip()


def extract_tables_from_page(page):
    """Returns (kept_tables, skipped_count).
    kept_tables is a list of pandas DataFrames containing ONLY the header
    row and data rows (nothing else). Tables that end up with zero data
    rows after cleaning are dropped entirely and counted in skipped_count."""
    kept = []
    skipped = 0
    try:
        tab_finder = page.find_tables()
    except Exception:
        return kept, skipped

    for tab in tab_finder.tables:
        try:
            data = tab.extract()
            if not data or len(data) < 1:
                continue
            header = data[0]
            rows = data[1:]

            clean_header = []
            for i, h in enumerate(header):
                h_str = _clean_cell(h)
                clean_header.append(h_str if h_str else f"Column_{i + 1}")
            # Guard against duplicate column names (Excel/pandas dislike them)
            seen = {}
            for i, h in enumerate(clean_header):
                if h in seen:
                    seen[h] += 1
                    clean_header[i] = f"{h}_{seen[h]}"
                else:
                    seen[h] = 0

            df = pd.DataFrame(rows, columns=clean_header)
            # Keep only header + data: strip whitespace and drop rows that
            # are entirely empty (no real data in any column).
            df = df.apply(lambda col: col.map(_clean_cell))
            df = df[~df.apply(lambda row: all(cell == "" for cell in row), axis=1)]
            df = df.reset_index(drop=True)

            if df.shape[0] == 0:
                # Header-only table (no data rows) -> ignore entirely
                skipped += 1
                continue

            kept.append(df)
        except Exception:
            continue
    return kept, skipped


def build_page_workbook(tables, target_unit, frac_denom=16, as_excel_table=False):
    """Builds one xlsx (in-memory) for a single page, one sheet per table.
    If as_excel_table is True, each sheet's range is formatted as a proper
    Excel Table (filter dropdowns + banded rows)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for i, df in enumerate(tables, start=1):
            converted = convert_dataframe(df, target_unit, frac_denom)
            sheet_name = f"Table_{i}"[:31]
            converted.to_excel(writer, sheet_name=sheet_name, index=False)

            if as_excel_table:
                ws = writer.sheets[sheet_name]
                n_rows = converted.shape[0] + 1  # + header row
                n_cols = converted.shape[1]
                if n_rows >= 2 and n_cols >= 1:
                    last_col = get_column_letter(n_cols)
                    ref = f"A1:{last_col}{n_rows}"
                    excel_table = Table(displayName=f"Tbl_{i}", ref=ref)
                    excel_table.tableStyleInfo = TableStyleInfo(
                        name="TableStyleMedium9",
                        showRowStripes=True,
                        showFirstColumn=False,
                        showLastColumn=False,
                        showColumnStripes=False,
                    )
                    ws.add_table(excel_table)
    output.seek(0)
    return output.getvalue()


# ============================================================
# UI
# ============================================================

st.title("📐 PDF Table Extractor → XLSX")
st.info(
    """
**How to use this tool:**
1. **Upload a PDF** below (e.g. drawings/specs with tables and dimensions).
2. Choose the **dimension unit** you want values converted to — Decimal or Architectural.
3. Choose whether extracted data should be formatted as an Excel Table.
4. Click **Extract Tables**. You'll get **one XLSX file per page**, with **one sheet per table**
   found on that page — only the header row and its data rows are included. Tables that have a
   header but no data rows are skipped entirely. Download individually or all together as a ZIP.
"""
)

st.divider()

st.subheader("1. Upload PDF")
uploaded_pdf = st.file_uploader("Upload the PDF you want to convert", type=["pdf"])

st.divider()

st.subheader("2. Dimension Unit")
unit_choice = st.radio(
    "Convert dimension-looking values in the tables to:",
    ["Decimal", "Architectural"],
    horizontal=True,
    help=(
        "Decimal: 12'-6\" -> 12.5. "
        "Architectural: 12.5 -> 12'-6\". "
        "Only values that look like dimensions are converted; other text/numbers are left as-is."
    ),
)
target_unit = "decimal" if unit_choice == "Decimal" else "architectural"

frac_denom = 16
if target_unit == "architectural":
    denom_choice = st.selectbox(
        "Round fractional inches to the nearest:",
        ["1/8\"", "1/16\"", "1/32\""],
        index=1,
    )
    frac_denom = int(denom_choice.split("/")[1].rstrip('"'))

st.divider()

st.subheader("3. Excel Table Formatting")
table_format_choice = st.radio(
    "Would you like the extracted data converted into a formatted Excel Table "
    "(adds column filter dropdowns and banded row styling)?",
    ["Yes", "No"],
    horizontal=True,
    help="If No, data is written as plain cells with a bold header row only.",
)
as_excel_table = table_format_choice == "Yes"

st.divider()

st.subheader("4. Extract")

if uploaded_pdf is None:
    st.warning("Upload a PDF above to continue.")
else:
    if st.button("✨ Extract Tables", type="primary", use_container_width=True):
        try:
            with st.spinner("Scanning PDF for tables..."):
                input_bytes = uploaded_pdf.getvalue()
                doc = pymupdf.open(stream=input_bytes, filetype="pdf")
                basename = os.path.splitext(uploaded_pdf.name)[0]

                page_files = {}
                page_summary = []  # (page_num, kept_count, skipped_count)
                preview_tables = {}

                for page_index in range(len(doc)):
                    page = doc.load_page(page_index)
                    tables, skipped_count = extract_tables_from_page(page)
                    page_num = page_index + 1
                    page_summary.append((page_num, len(tables), skipped_count))
                    if tables:
                        converted_tables = [convert_dataframe(df, target_unit, frac_denom) for df in tables]
                        wb_bytes = build_page_workbook(
                            tables, target_unit, frac_denom, as_excel_table=as_excel_table
                        )
                        page_files[page_num] = wb_bytes
                        preview_tables[page_num] = converted_tables

                doc.close()

            st.session_state.page_files = page_files
            st.session_state.page_summary = page_summary
            st.session_state.basename = basename
            st.session_state.preview_tables = preview_tables

        except Exception as e:
            st.error(f"Extraction failed: {type(e).__name__}: {e}")
            with st.expander("Show detailed error logs"):
                st.code(traceback.format_exc())

# ============================================================
# Results
# ============================================================

if st.session_state.page_files is not None:
    st.divider()
    st.subheader("5. Results")

    page_files = st.session_state.page_files
    page_summary = st.session_state.page_summary
    basename = st.session_state.basename
    preview_tables = st.session_state.preview_tables

    total_pages = len(page_summary)
    pages_with_tables = len(page_files)
    total_tables = sum(kept for _, kept, _ in page_summary)
    total_skipped = sum(skipped for _, _, skipped in page_summary)

    st.success(
        f"Scanned {total_pages} page(s). Kept {total_tables} table(s) with data across "
        f"{pages_with_tables} page(s)."
    )
    if total_skipped:
        st.caption(f"Ignored {total_skipped} header-only table(s) that had no data rows.")

    with st.expander("Per-page table counts"):
        for pg, kept, skipped in page_summary:
            note = f" ({skipped} header-only table(s) ignored)" if skipped else ""
            st.write(f"• Page {pg}: {kept} table(s) kept{note}")

    no_table_pages = [pg for pg, kept, _ in page_summary if kept == 0]
    if no_table_pages:
        st.warning(
            "No usable tables found on page(s): " + ", ".join(str(p) for p in no_table_pages) +
            ". These pages were skipped (no file generated)."
        )

    if page_files:
        # Build a ZIP with one xlsx per page
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for pg, data in page_files.items():
                zf.writestr(f"{basename}_page{pg}.xlsx", data)
        zip_buffer.seek(0)

        st.download_button(
            "📦 Download All Pages (ZIP)",
            data=zip_buffer,
            file_name=f"{basename}_tables.zip",
            mime="application/zip",
            use_container_width=True,
        )

        st.markdown("**Or download individual page files:**")
        kept_lookup = {pg: kept for pg, kept, _ in page_summary}
        for pg, data in page_files.items():
            table_count = kept_lookup[pg]
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"Page {pg} — {table_count} table(s)")
            with col2:
                st.download_button(
                    "⬇️ Download",
                    data=data,
                    file_name=f"{basename}_page{pg}.xlsx",
                    key=f"dl_page_{pg}",
                    use_container_width=True,
                )

        with st.expander("Preview extracted tables"):
            for pg, dfs in preview_tables.items():
                st.markdown(f"**Page {pg}**")
                for i, df in enumerate(dfs, start=1):
                    st.caption(f"Table {i}")
                    st.dataframe(df, use_container_width=True)
    else:
        st.warning("No tables were detected anywhere in this PDF, so no files were generated.")
