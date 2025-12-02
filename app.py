import streamlit as st
import pandas as pd
import re
from io import BytesIO
import zipfile
import requests
from pathlib import Path
import tempfile
import os

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from docx.image.exceptions import UnrecognizedImageError
from docx2pdf import convert


# ---------- GOOGLE SHEET CSV LOADER ----------

def extract_sheet_id(sheet_input: str) -> str:
    """
    Accepts either:
    - Full Google Sheets URL
    - Or raw Sheet ID
    Returns the Sheet ID.
    Example URL:
    https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxx/edit#gid=0
    """
    url_pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(url_pattern, sheet_input)
    if match:
        return match.group(1)
    return sheet_input.strip()


def load_sheet_via_csv(sheet_input: str, gid: str | None = None) -> pd.DataFrame:
    """
    Fetch Google Sheet as CSV using public export URL.
    Sheet must be shared as 'Anyone with link can view'.
    """
    sheet_id = extract_sheet_id(sheet_input)
    base_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    if gid:
        base_url += f"&gid={gid}"
    df = pd.read_csv(base_url)
    return df


# ---------- SITE NAME PARSER ----------

def parse_site_name(raw: str):
    """
    Input:  "4-361-Candid Manesar"
    Output: zone='4', unit_code='361', site_name='Candid Manesar'
    """
    if not isinstance(raw, str):
        return "", "", ""

    parts = raw.split("-", 2)  # split into max 3 parts
    if len(parts) < 3:
        # Fallback: treat whole string as site_name
        return "", "", raw.strip()

    zone = parts[0].strip()
    unit_code = parts[1].strip()
    sitename = parts[2].strip()
    return zone, unit_code, sitename


# ---------- GOOGLE DRIVE IMAGE DOWNLOADER ----------

def extract_drive_file_id(url: str) -> str:
    """
    Supports common Drive URL formats:
    - https://drive.google.com/open?id=FILEID
    - https://drive.google.com/file/d/FILEID/view
    """
    patterns = [
        r"id=([A-Za-z0-9_-]+)",
        r"/d/([A-Za-z0-9_-]+)/",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return ""


def download_drive_image(url: str) -> BytesIO | None:
    """
    Download an image from a Google Drive share link.
    Returns BytesIO if content is a real image, otherwise None.
    """
    file_id = extract_drive_file_id(url)
    if not file_id:
        return None

    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    try:
        resp = requests.get(download_url, timeout=15)
        if resp.status_code == 200:
            content_type = resp.headers.get("Content-Type", "")
            # Only accept real images
            if content_type.startswith("image/"):
                return BytesIO(resp.content)
    except Exception:
        return None

    return None


# ---------- WORD DOCX GENERATION ----------

def render_docx_for_row(row: pd.Series, template_path: str) -> BytesIO:
    """
    Render template.docx for a single row and return BytesIO of the .docx file,
    embedding images at the bottom.
    """

    # Destructure site name
    raw_site = row.get("Site Name", "")
    zone, unit_code, site_name_clean = parse_site_name(raw_site)

    # Download & prepare images
    images_raw = str(row.get("Images", "") or "").strip()
    image_objs: list[BytesIO] = []
    if images_raw:
        urls = [u.strip() for u in images_raw.split(",") if u.strip()]
        for url in urls:
            img_bytes = download_drive_image(url)
            if img_bytes:
                image_objs.append(img_bytes)

    # Build context for template
    context = {
        "zone": zone,
        "unit_code": unit_code,
        "site_name": site_name_clean,
        "date": row.get("Date", ""),
        "time": row.get("Time", ""),
        "attendance_register": row.get("Documentation Check [Attendance Register]", ""),
        "handling_register": row.get("Documentation Check [Handling / Taking Over Register]", ""),
        "material_register": row.get("Documentation Check [Visitor Log Register]", ""),
        "grooming": row.get("Performance Check [Grooming]", ""),
        "alertness": row.get("Performance Check [Alertness]", ""),
        "post_discipline": row.get("Performance Check [Post Discipline]", ""),
        "overall_rating": row.get("Performance Check [Overall Rating]", ""),
        "observation": row.get("Observation", ""),
        "inspected_by": row.get("Inspected By", ""),
    }

    # Clean NaN / None
    for k, v in list(context.items()):
        if pd.isna(v):
            context[k] = ""
        elif v is None:
            context[k] = ""

    tpl = DocxTemplate(template_path)

    # Create InlineImage list (can be empty; template uses {% for img in images %})
    images_inline = []
    for img in image_objs:
        try:
            images_inline.append(InlineImage(tpl, img, width=Inches(2.5)))
        except UnrecognizedImageError:
            # Skip invalid / non-image content silently
            continue

    context["images"] = images_inline

    # Render the template
    tpl.render(context)

    # Save to buffer
    buf = BytesIO()
    tpl.save(buf)
    buf.seek(0)
    return buf


# ---------- STREAMLIT APP ----------

def main():
    st.set_page_config(
        page_title="SGV Night Check Report",
        layout="wide"
    )

    st.title("Night Check Report Generator (PDF)")

    # Sidebar: Google Sheet settings
    st.sidebar.header("Google Sheet Settings")
    sheet_input = st.sidebar.text_input("Sheet URL or Sheet ID")
    gid = st.sidebar.text_input("Worksheet gid (optional)")
    st.sidebar.write("Sheet must be shared as: **Anyone with link → Viewer**")

    if st.sidebar.button("Fetch Data"):
        if not sheet_input:
            st.error("Please enter Google Sheet URL or ID.")
        else:
            try:
                df = load_sheet_via_csv(sheet_input, gid if gid else None)

                if "Date" not in df.columns:
                    st.error("Column 'Date' not found in sheet.")
                else:
                    # Parse Date column once and store in session
                    df["Date_parsed"] = pd.to_datetime(df["Date"], errors="coerce")
                    st.session_state["df"] = df
                    st.success("✅ Data loaded successfully.")
            except Exception as e:
                st.error("Failed fetching sheet:")
                st.exception(e)

    # Main content: only show if df is loaded
    if "df" not in st.session_state:
        st.info("Use the sidebar to fetch data from Google Sheets to continue.")
        return

    df = st.session_state["df"]

    # Validate Date_parsed column
    if "Date_parsed" not in df.columns:
        st.error("Internal error: 'Date_parsed' column missing.")
        return

    valid_dates = sorted(df["Date_parsed"].dropna().dt.date.unique())
    if not valid_dates:
        st.error("No valid dates found in 'Date' column.")
        return

    selected_date = st.date_input(
        "Select Date",
        value=valid_dates[-1],
        min_value=valid_dates[0],
        max_value=valid_dates[-1]
    )

    df_date = df[df["Date_parsed"].dt.date == selected_date]

    st.subheader(f"Records for {selected_date}")
    if df_date.empty:
        st.warning("No records for this date.")
        return

    st.dataframe(df_date)

    template_path = "template.docx"
    if not Path(template_path).exists():
        st.error("❌ template.docx not found in project folder.")
        return

    if st.button("Generate Report"):
        try:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for idx, row in df_date.iterrows():

                    # 1. Generate DOCX in memory
                    docx_buf = render_docx_for_row(row, template_path)

                    # 2. Save temporary DOCX
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                        tmp_docx.write(docx_buf.getvalue())
                        tmp_docx_path = tmp_docx.name

                    # 3. Convert DOCX → PDF
                    tmp_pdf_path = tmp_docx_path.replace(".docx", ".pdf")
                    convert(tmp_docx_path, tmp_pdf_path)

                    # 4. Load PDF bytes
                    with open(tmp_pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()

                    # 5. Build filename
                    zone, unit_code, sitename = parse_site_name(row.get("Site Name", "Site"))
                    site_slug = (sitename or "Site").replace(" ", "_")
                    time_val = str(row.get("Time", "")).strip()
                    time_slug = time_val.replace(":", "").replace(" ", "_") if time_val else str(idx)

                    filename = f"{selected_date}_{site_slug}.pdf"

                    # 6. Add to ZIP
                    zipf.writestr(filename, pdf_bytes)

                    # 7. Cleanup temp files
                    try:
                        os.remove(tmp_docx_path)
                    except OSError:
                        pass
                    try:
                        os.remove(tmp_pdf_path)
                    except OSError:
                        pass

            zip_buffer.seek(0)
            st.download_button(
                "⬇️ Download ZIP (PDF)",
                data=zip_buffer,
                file_name=f"night_checks_{selected_date}.zip",
                mime="application/zip"
            )
        except Exception as e:
            st.error("Error generating PDF ZIP:")
            st.exception(e)


if __name__ == "__main__":
    main()
