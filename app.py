import streamlit as st
import pandas as pd
import os
import io
import re
import zipfile
import traceback
from fractions import Fraction

import pymupdf  # PyMuPDF

from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ============================================================
# Google OAuth Config
# ============================================================
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]


def _google_oauth_configured():
    return "google_oauth" in st.secrets and all(
        k in st.secrets["google_oauth"] for k in ("client_id", "client_secret", "redirect_uri")
    )


def _get_flow():
    cfg = st.secrets["google_oauth"]
    client_config = {
        "web": {
            "client_id": cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [cfg["redirect_uri"]],
        }
    }
    return Flow.from_client_config(
        client_config, scopes=GOOGLE_SCOPES, redirect_uri=cfg["redirect_uri"]
    )


def _handle_oauth_redirect():
    """If we just got redirected back from Google with a ?code=..., exchange
    it for credentials and store them in session state."""
    if "google_creds" in st.session_state:
        return
    params = st.query_params
    if "code" in params:
        try:
            flow = _get_flow()
            flow.fetch_token(code=params["code"])
            st.session_state.google_creds = flow.credentials
        except Exception as e:
            st.session_state.google_auth_error = f"{type(e).__name__}: {e}"
        finally:
            st.query_params.clear()
            st.rerun()


def upload_to_google_sheets(creds, title, page_tables):
    """Create one Google Sheet, one tab per (page, table), and fill it in.
    page_tables: {page_num: [df, df, ...]} (already unit-converted)
    Returns the shareable edit URL."""
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    spreadsheet = sheets_service.spreadsheets().create(
        body={"properties": {"title": title}}, fields="spreadsheetId"
    ).execute()
    spreadsheet_id = spreadsheet["spreadsheetId"]

    # Build the list of tab names first
    tab_names = []
    for pg, dfs in page_tables.items():
        for i in range(1, len(dfs) + 1):
            tab_names.append(f"P{pg}_T{i}"[:100])

    requests = []
    for idx, name in enumerate(tab_names):
        if idx == 0:
            requests.append({
                "updateSheetProperties": {
                    "properties": {"sheetId": 0, "title": name},
                    "fields": "title",
                }
            })
        else:
            requests.append({"addSheet": {"properties": {"title": name}}})

    if requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()

    # Write the data into each tab
    data_updates = []
    tab_iter = iter(tab_names)
    for pg, dfs in page_tables.items():
        for df in dfs:
            name = next(tab_iter)
            values = [list(map(str, df.columns))] + df.astype(str).values.tolist()
            data_updates.append({"range": f"'{name}'!A1", "values": values})

    if data_updates:
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"valueInputOption": "RAW", "data": data_updates},
        ).execute()

    # Make it viewable by anyone with the link (does not affect edit rights)
    try:
        drive_service.permissions().create(
            fileId=spreadsheet_id, body={"role": "reader", "type": "anyone"}
        ).execute()
    except HttpError:
        pass  # sharing permission failed; owner can still open/share manually

    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"


# ============================================================
# Session State Init
# ============================================================
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.page_files = None      # {page_num: xlsx_bytes}
    st.session_state.page_summary = None    # [(page_num, table_count)]
    st.session_state.basename = None
    st.session_state.preview_tables = None  # {page_num: [df, df, ...]}
    st.session_state.gsheet_link = None
    st.session_state.google_auth_error = None

st.set_page_config(page_title="PDF Table Extractor", layout="wide")

_handle_oauth_redirect()

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

def extract_tables_from_page(page):
    """Returns a list of pandas DataFrames, one per table found on the page."""
    dfs = []
    try:
        tab_finder = page.find_tables()
    except Exception:
        return dfs

    for tab in tab_finder.tables:
        try:
            data = tab.extract()
            if not data or len(data) < 1:
                continue
            header = data[0]
            rows = data[1:]
            clean_header = []
            for i, h in enumerate(header):
                h_str = str(h).strip() if h is not None and str(h).strip() != "" else f"Column_{i + 1}"
                clean_header.append(h_str)
            # Guard against duplicate column names (Excel/pandas dislike them)
            seen = {}
            for i, h in enumerate(clean_header):
                if h in seen:
                    seen[h] += 1
                    clean_header[i] = f"{h}_{seen[h]}"
                else:
                    seen[h] = 0
            df = pd.DataFrame(rows, columns=clean_header)
            dfs.append(df)
        except Exception:
            continue
    return dfs


def build_page_workbook(tables, target_unit, frac_denom=16):
    """Builds one xlsx (in-memory) for a single page, one sheet per table."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for i, df in enumerate(tables, start=1):
            converted = convert_dataframe(df, target_unit, frac_denom)
            sheet_name = f"Table_{i}"[:31]
            converted.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output.getvalue()


# ============================================================
# UI
# ============================================================

st.title("📐 PDF Table Extractor → XLSX / Google Sheets")
st.info(
    """
**How to use this tool:**
1. **Upload a PDF** below (e.g. drawings/specs with tables and dimensions).
2. Choose the **dimension unit** you want values converted to — Decimal or Architectural.
3. Choose where you want the results to go — **XLSX download**, **Google Sheets**, or **both**.
4. Click **Extract Tables**. Download the XLSX files and/or push the results straight to a Google Sheet in your own Google account.
"""
)

if st.session_state.google_auth_error:
    st.error(f"Google sign-in failed: {st.session_state.google_auth_error}")
    st.session_state.google_auth_error = None

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

st.subheader("3. Output Destination")
output_choice = st.radio(
    "Where should the extracted tables go?",
    ["XLSX files (download)", "Google Sheets (online)", "Both"],
    horizontal=True,
)
want_xlsx = output_choice in ("XLSX files (download)", "Both")
want_gsheet = output_choice in ("Google Sheets (online)", "Both")

if want_gsheet:
    if not _google_oauth_configured():
        st.warning(
            "Google Sheets export isn't configured yet. An admin needs to add a "
            "`[google_oauth]` section (client_id, client_secret, redirect_uri) to "
            "`.streamlit/secrets.toml` — see the setup guide below."
        )
        with st.expander("🔧 Google Sheets setup guide (for the app admin)"):
            st.markdown(
                """
1. In [Google Cloud Console](https://console.cloud.google.com/), create (or pick) a project.
2. Enable the **Google Sheets API** and **Google Drive API**.
3. Configure the OAuth consent screen (External is fine for testing; add yourself as a test user).
4. Create an **OAuth client ID** of type **Web application**.
5. Add an **Authorized redirect URI** matching where this app is hosted, e.g.
   `http://localhost:8501` for local dev, or `https://your-app-url` in production.
6. Add these values to `.streamlit/secrets.toml`:
```toml
[google_oauth]
client_id = "xxxxxxxx.apps.googleusercontent.com"
client_secret = "xxxxxxxx"
redirect_uri = "http://localhost:8501"
```
7. Restart the app.
"""
            )
    else:
        st.markdown("**Google account:**")
        if "google_creds" not in st.session_state:
            flow = _get_flow()
            auth_url, _ = flow.authorization_url(
                prompt="consent", access_type="offline", include_granted_scopes="true"
            )
            st.link_button("🔐 Sign in with Google", auth_url, use_container_width=False)
        else:
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.success("Signed in with Google ✅ — ready to create a Google Sheet.")
            with col_b:
                if st.button("Sign out"):
                    del st.session_state.google_creds
                    st.session_state.gsheet_link = None
                    st.rerun()

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
                page_summary = []
                preview_tables = {}

                for page_index in range(len(doc)):
                    page = doc.load_page(page_index)
                    tables = extract_tables_from_page(page)
                    page_num = page_index + 1
                    page_summary.append((page_num, len(tables)))
                    if tables:
                        converted_tables = [convert_dataframe(df, target_unit, frac_denom) for df in tables]
                        preview_tables[page_num] = converted_tables
                        if want_xlsx:
                            wb_bytes = build_page_workbook(tables, target_unit, frac_denom)
                            page_files[page_num] = wb_bytes

                doc.close()

            st.session_state.page_files = page_files
            st.session_state.page_summary = page_summary
            st.session_state.basename = basename
            st.session_state.preview_tables = preview_tables
            st.session_state.gsheet_link = None

        except Exception as e:
            st.error(f"Extraction failed: {type(e).__name__}: {e}")
            with st.expander("Show detailed error logs"):
                st.code(traceback.format_exc())

# ============================================================
# Results
# ============================================================

if st.session_state.preview_tables is not None:
    st.divider()
    st.subheader("5. Results")

    page_files = st.session_state.page_files or {}
    page_summary = st.session_state.page_summary
    basename = st.session_state.basename
    preview_tables = st.session_state.preview_tables

    total_pages = len(page_summary)
    pages_with_tables = len(preview_tables)
    total_tables = sum(c for _, c in page_summary)

    st.success(
        f"Scanned {total_pages} page(s). Found {total_tables} table(s) across "
        f"{pages_with_tables} page(s) with data."
    )

    with st.expander("Per-page table counts"):
        for pg, cnt in page_summary:
            st.write(f"• Page {pg}: {cnt} table(s)")

    no_table_pages = [pg for pg, cnt in page_summary if cnt == 0]
    if no_table_pages:
        st.warning(
            "No tables detected on page(s): " + ", ".join(str(p) for p in no_table_pages) +
            ". These pages were skipped."
        )

    if not preview_tables:
        st.warning("No tables were detected anywhere in this PDF, so no output was generated.")
    else:
        # ---------------- XLSX output ----------------
        if want_xlsx and page_files:
            st.markdown("### 📁 XLSX files")
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
            for pg, data in page_files.items():
                table_count = dict(page_summary)[pg]
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

        # ---------------- Google Sheets output ----------------
        if want_gsheet:
            st.markdown("### 🟢 Google Sheets")
            if not _google_oauth_configured():
                st.info("Configure Google OAuth (see setup guide above) to enable this.")
            elif "google_creds" not in st.session_state:
                st.info("Sign in with Google above, then click the button below.")
            else:
                if st.button("📤 Create Google Sheet with these tables", use_container_width=True):
                    with st.spinner("Creating your Google Sheet..."):
                        try:
                            link = upload_to_google_sheets(
                                st.session_state.google_creds,
                                f"{basename}_extracted_tables",
                                preview_tables,
                            )
                            st.session_state.gsheet_link = link
                        except HttpError as e:
                            st.error(f"Google Sheets API error: {e}")
                        except Exception as e:
                            st.error(f"Upload failed: {type(e).__name__}: {e}")
                            with st.expander("Show detailed error logs"):
                                st.code(traceback.format_exc())

                if st.session_state.gsheet_link:
                    st.success("Your Google Sheet is ready!")
                    st.link_button(
                        "📄 Open Google Sheet", st.session_state.gsheet_link, use_container_width=True
                    )

        # ---------------- Preview ----------------
        with st.expander("Preview extracted tables"):
            for pg, dfs in preview_tables.items():
                st.markdown(f"**Page {pg}**")
                for i, df in enumerate(dfs, start=1):
                    st.caption(f"Table {i}")
                    st.dataframe(df, use_container_width=True)
