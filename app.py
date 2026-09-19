import streamlit as st
import io
import os
import re
import csv
import gzip
import zipfile
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


# =========================================================
# PAGE SETUP
# =========================================================

st.set_page_config(
    page_title="GJM Bank Report Converter",
    page_icon="🏦",
    layout="centered"
)

st.markdown("""
<style>
.main-title {
    text-align: center;
    color: #1F4E78;
    font-size: 26px;
    font-weight: bold;
    margin-bottom: 5px;
}
.sub-title {
    text-align: center;
    color: #555555;
    font-size: 15px;
    margin-bottom: 20px;
}
.stButton button {
    width: 100%;
    min-height: 52px;
    font-size: 16px;
    font-weight: bold;
    border-radius: 8px;
}
.stDownloadButton button {
    width: 100%;
    min-height: 52px;
    background-color: #1F4E78;
    color: white;
    font-weight: bold;
    font-size: 16px;
    border-radius: 8px;
}
@media (max-width: 600px) {
    .main-title { font-size: 21px !important; line-height: 1.3; }
    .sub-title { font-size: 13px !important; line-height: 1.5; }
    .stButton button, .stDownloadButton button {
        min-height: 54px !important;
        font-size: 16px !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="main-title">🏦 GJM Bank Reports to Excel Converter</div>',
    unsafe_allow_html=True
)
st.markdown(
    '<div class="sub-title">📱 Mobile અને Desktop બંને માટે • ZIP / TXT / GZ → Excel</div>',
    unsafe_allow_html=True
)


# =========================================================
# REPORT CLEANING / PARSING
# =========================================================

IGNORE_PATTERNS = [
    "REPORT ID:", "PROC DATE:", "RUN DATE:", "BRANCH :", "BRANCH NO.",
    "PAGE NO", "PRODUCT TOTAL", "ACCOUNT TYPE TOTAL", "OVER DRAWN TOTAL",
    "NO OF ACCOUNTS", "SUB TOTAL", "GRAND TOTAL", "BALANCE FORWARD",
    "BHAVNAGAR DISTRICT", "TOTAL FOR PRODUCT", "TOTAL NO OF",
    "PRODUCT-WISE TOTAL", "====>", "GL CLASS CODE", "GL-CLASS-CODE", "AREA:"
]


def sanitize_text(val):
    if val is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(val)).strip()


def clean_dr_amount(val):
    val_str = sanitize_text(val)
    if not val_str:
        return ""

    if re.search(r"\bDr\b", val_str, re.IGNORECASE) or val_str.lower().endswith("dr"):
        num_clean = re.sub(r"(?i)\s*dr\s*", "", val_str).strip()
        return f"-{num_clean}" if not num_clean.startswith("-") else num_clean

    if re.search(r"\bCr\b", val_str, re.IGNORECASE):
        return re.sub(r"(?i)\s*cr\s*", "", val_str).strip()

    return val_str


def parse_deposits_balance_report(lines):
    pattern = re.compile(
        r"^\s*(\d{11}-\d|\d{8,16})\s+"
        r"(\S+(?:\s\S+)*)\s{2,}"
        r"(.+?)\s{2,}"
        r"([\d,]+\.\d{2})\s+"
        r"([\d,]+\.\d{2})\s+"
        r"([\d,]+\.\d{2}(?:\s*Dr)?)\s+"
        r"([\d,]+\.\d{2})\s*"
        r"(?:\s+(\d+[DMY]))?"
        r"\s+([\d,]+\.\d{2})"
        r"\s+([A-Z]+)"
        r"\s+([YN])\s*$",
        re.IGNORECASE
    )

    rows = []
    for line in lines:
        clean_line = sanitize_text(line)
        match = pattern.match(clean_line)
        if match:
            rows.append({
                "ACCOUNT NUMBER": match.group(1).strip(),
                "ACCOUNT TYPE (DESCRIPTION)": match.group(2).strip(),
                "CUSTOMER NAME": match.group(3).strip(),
                "AVAILABLE BALANCE": match.group(4).strip(),
                "UNCLEARED BALANCE": match.group(5).strip(),
                "CURRENT BALANCE": clean_dr_amount(match.group(6)),
                "LIMIT": match.group(7).strip(),
                "TERM": match.group(8).strip() if match.group(8) else "",
                "INT-RATE": match.group(9).strip(),
                "STATUS": match.group(10).strip(),
                "JOINT-HOLD-FLAG": match.group(11).strip()
            })

    return pd.DataFrame(rows) if rows else None


def parse_pipe_delimited(lines):
    header_cols = []
    data_rows = []

    for line in lines:
        clean_line = sanitize_text(line)
        if not clean_line:
            continue
        if re.match(r"^-{10,}", clean_line):
            continue
        if any(kw in clean_line.upper() for kw in IGNORE_PATTERNS):
            continue
        if "|" not in clean_line:
            continue

        parts = [clean_dr_amount(p) for p in clean_line.split("|")]
        if parts and parts[0] == "":
            parts = parts[1:]
        if parts and parts[-1] == "":
            parts = parts[:-1]

        if not header_cols and any(
            w in clean_line.upper()
            for w in ["NAME", "ACCOUNT", "AMOUNT", "LIMIT", "DATE", "PRODUCT"]
        ):
            header_cols = parts
        elif header_cols and len(parts) >= 2:
            if len(parts) < len(header_cols):
                parts.extend([""] * (len(header_cols) - len(parts)))
            data_rows.append(parts[:len(header_cols)])

    return pd.DataFrame(data_rows, columns=header_cols) if header_cols and data_rows else None


def parse_general_cbs_report(lines):
    header_index = -1

    for i, line in enumerate(lines):
        clean_line = sanitize_text(line)
        if re.match(r"^\s*-{15,}", clean_line) and i + 2 < len(lines):
            if re.match(r"^\s*-{15,}", sanitize_text(lines[i + 2])):
                header_index = i + 1
                break

    if header_index == -1:
        return None

    cols = [
        m.group(1).strip()
        for m in re.finditer(
            r"(\S+(?:\s(?!\s)\S+)*)",
            sanitize_text(lines[header_index])
        )
    ]
    if len(cols) < 2:
        return None

    rows = []
    for line in lines[header_index + 2:]:
        clean_line = sanitize_text(line)
        if not clean_line:
            continue
        if re.match(r"^-{10,}", clean_line):
            continue
        if any(kw in clean_line.upper() for kw in IGNORE_PATTERNS):
            continue
        if "ACCOUNT NUMBER" in clean_line.upper():
            continue

        parts = [
            p.strip()
            for p in re.split(r"\s{2,}", clean_line)
            if p.strip()
        ]

        if len(parts) > 1:
            rows.append({
                cols[k]: clean_dr_amount(parts[k]) if k < len(parts) else ""
                for k in range(len(cols))
            })

    return pd.DataFrame(rows) if rows else None


def parse_csv_generic(content):
    clean_content = sanitize_text(content)
    if not clean_content.strip():
        return None

    try:
        delimiter = csv.Sniffer().sniff(
            clean_content[:4096],
            delimiters=[",", "\t", "|", ";"]
        ).delimiter
    except Exception:
        delimiter = "," if "," in clean_content else "\t"

    try:
        df = pd.read_csv(
            io.StringIO(clean_content),
            sep=delimiter,
            engine="python",
            on_bad_lines="skip",
            encoding_errors="ignore"
        )
    except Exception:
        return None

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(clean_dr_amount)

    return df if not df.empty else None


def parse_report(content_string):
    lines = content_string.splitlines()
    upper_content = content_string.upper()

    df = None

    if "DEPOSITS BALANCE FILE" in upper_content or "AVAILABLE BALANCE" in upper_content:
        df = parse_deposits_balance_report(lines)

    if df is None or df.empty:
        pipe_count = sum(1 for line in lines[:50] if line.count("|") >= 3)
        if pipe_count >= 3:
            df = parse_pipe_delimited(lines)

    if df is None or df.empty:
        df = parse_general_cbs_report(lines)

    if df is None or df.empty:
        df = parse_csv_generic(content_string)

    return df


# =========================================================
# EXCEL CREATION
# =========================================================

def convert_to_excel(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Converted Report"

    clean_headers = [sanitize_text(c) for c in df.columns]
    ws.append(clean_headers)

    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )
    header_font = Font(
        name="Calibri", size=11, bold=True, color="FFFFFF"
    )
    data_font = Font(name="Calibri", size=10)

    border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for row in df.itertuples(index=False):
        ws.append([sanitize_text(v) for v in row])

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column
    ):
        for cell in row:
            cell.font = data_font
            cell.border = border
            value_string = str(cell.value or "").strip()

            if re.match(r"^-?[\d,]+(\.\d+)?$", value_string):
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

    for column in ws.columns:
        max_length = max(
            len(str(cell.value or "")) for cell in column
        )
        letter = get_column_letter(column[0].column)
        ws.column_dimensions[letter].width = min(
            max(max_length + 3, 12),
            50
        )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# FILE DECODING
# =========================================================

def decode_text(data):
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            pass
    return data.decode("latin1", errors="replace"), "latin1-replace"


def decompress_gzip(data):
    try:
        return gzip.decompress(data)
    except Exception as e:
        raise ValueError(f"GZIP decompression failed: {e}")


# =========================================================
# EXTRACT ZIP FILES
# =========================================================

def extract_input_files(uploaded_name, uploaded_bytes):
    """
    Returns:
        file_items = [(display_name, raw_text_bytes), ...]
        diagnostics = list of dictionaries
    """

    file_items = []
    diagnostics = []

    lower_name = uploaded_name.lower()

    # -------------------------------
    # ZIP
    # -------------------------------
    if lower_name.endswith(".zip"):

        with zipfile.ZipFile(io.BytesIO(uploaded_bytes), "r") as z:
            infos = z.infolist()

            for info in infos:

                if info.is_dir():
                    continue

                archive_name = info.filename
                base_name = os.path.basename(archive_name)

                if not base_name or base_name.startswith("."):
                    continue

                try:
                    raw = z.read(info)

                    if base_name.lower().endswith(".gz"):
                        raw = decompress_gzip(raw)
                        logical_name = os.path.splitext(base_name)[0]
                        compression = "GZIP → TXT"
                    else:
                        logical_name = base_name
                        compression = "Plain"

                    file_items.append((logical_name, raw))

                    diagnostics.append({
                        "File": archive_name,
                        "Size (KB)": round(len(raw) / 1024, 1),
                        "Type": compression,
                        "Status": "Extracted"
                    })

                except Exception as e:
                    diagnostics.append({
                        "File": archive_name,
                        "Size (KB)": 0,
                        "Type": "Unknown",
                        "Status": f"ERROR: {e}"
                    })

        return file_items, diagnostics

    # -------------------------------
    # Standalone GZ
    # -------------------------------
    if lower_name.endswith(".gz"):

        raw = decompress_gzip(uploaded_bytes)
        logical_name = os.path.splitext(uploaded_name)[0]

        file_items.append((logical_name, raw))

        diagnostics.append({
            "File": uploaded_name,
            "Size (KB)": round(len(raw) / 1024, 1),
            "Type": "GZIP → TXT",
            "Status": "Extracted"
        })

        return file_items, diagnostics

    # -------------------------------
    # Plain TXT / CSV / DAT / LOG
    # -------------------------------

    file_items.append((uploaded_name, uploaded_bytes))

    diagnostics.append({
        "File": uploaded_name,
        "Size (KB)": round(len(uploaded_bytes) / 1024, 1),
        "Type": "Plain",
        "Status": "Ready"
    })

    return file_items, diagnostics


# =========================================================
# USER UPLOAD
# =========================================================

uploaded_file = st.file_uploader(
    "📁 ZIP / TXT / GZ ફાઇલ પસંદ કરો",
    type=None,
    accept_multiple_files=False,
    key="bank_report_upload",
    help="તમારી Bank Report ZIP અથવા TXT/GZ file પસંદ કરો"
)


if uploaded_file is not None:

    uploaded_name = uploaded_file.name
    uploaded_bytes = uploaded_file.getvalue()

    st.info(
        f"📄 **{uploaded_name}**  •  "
        f"📦 {len(uploaded_bytes) / 1024 / 1024:.2f} MB"
    )

    try:

        file_items, diagnostics = extract_input_files(
            uploaded_name,
            uploaded_bytes
        )

    except zipfile.BadZipFile:

        st.error("❌ ZIP file valid નથી અથવા corrupt છે.")
        st.stop()

    except Exception as e:

        st.error(f"❌ File reading error: {e}")
        st.stop()

    # ---------------------------------------------
    # ZIP / GZ diagnostic
    # ---------------------------------------------

    if lower_name := uploaded_name.lower():
        if lower_name.endswith(".zip") or lower_name.endswith(".gz"):

            st.success(
                f"✅ {len(file_items)} report files successfully extracted/read."
            )

            with st.expander("📋 Files found inside upload"):

                if diagnostics:
                    st.dataframe(
                        pd.DataFrame(diagnostics),
                        use_container_width=True,
                        hide_index=True
                    )

    total_files = len(file_items)

    if total_files == 0:
        st.warning(
            "⚠️ Uploadની અંદર કોઈ process કરી શકાય તેવી file મળી નથી."
        )
        st.stop()

    # ---------------------------------------------
    # Convert
    # ---------------------------------------------

    if st.button(
        "🚀 CONVERT TO EXCEL",
        type="primary",
        use_container_width=True
    ):

        progress = st.progress(0)
        status = st.empty()

        converted_files = {}
        failed_files = []

        for index, (filename, raw_bytes) in enumerate(file_items):

            status.text(
                f"⏳ Processing: {filename} "
                f"({index + 1}/{total_files})"
            )

            try:

                content_string, encoding = decode_text(raw_bytes)

                df = parse_report(content_string)

                if df is None or df.empty:
                    failed_files.append(
                        f"{filename} — Data structure not recognized"
                    )
                else:

                    excel_data = convert_to_excel(df)

                    base_name = os.path.splitext(
                        os.path.basename(filename)
                    )[0]

                    output_name = f"{base_name}.xlsx"

                    converted_files[output_name] = excel_data

            except Exception as e:

                failed_files.append(
                    f"{filename} — {e}"
                )

            progress.progress(
                (index + 1) / total_files
            )

        status.empty()
        progress.empty()

        st.divider()

        st.success(
            f"✅ Conversion completed: "
            f"{len(converted_files)} / {total_files}"
        )

        if failed_files:

            with st.expander(
                f"⚠️ Failed / Not recognized ({len(failed_files)})"
            ):
                for item in failed_files:
                    st.write(f"• {item}")

        # -----------------------------------------
        # One Excel
        # -----------------------------------------

        if len(converted_files) == 1:

            output_name, excel_data = next(
                iter(converted_files.items())
            )

            st.download_button(
                label=f"📥 DOWNLOAD {output_name}",
                data=excel_data,
                file_name=output_name,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        # -----------------------------------------
        # Multiple Excel files
        # -----------------------------------------

        elif len(converted_files) > 1:

            result_zip = io.BytesIO()

            with zipfile.ZipFile(
                result_zip,
                "w",
                zipfile.ZIP_DEFLATED
            ) as z:

                for filename, excel_data in converted_files.items():
                    z.writestr(filename, excel_data)

            result_zip.seek(0)

            st.download_button(
                label=(
                    f"📦 DOWNLOAD ALL EXCEL FILES "
                    f"({len(converted_files)} FILES)"
                ),
                data=result_zip.getvalue(),
                file_name="Converted_Bank_Reports.zip",
                mime="application/zip",
                use_container_width=True
            )

        else:

            st.error(
                "❌ એક પણ report Excelમાં convert થઈ શક્યો નથી."
            )
            
