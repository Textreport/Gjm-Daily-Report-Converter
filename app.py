import streamlit as st
import io
import os
import re
import csv
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


# =========================================================
# MOBILE FRIENDLY CSS
# =========================================================

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
    min-height: 50px;
    font-size: 16px;
    font-weight: bold;
    border-radius: 8px;
}

.stDownloadButton button {
    width: 100%;
    min-height: 50px;
    background-color: #1F4E78;
    color: white;
    font-weight: bold;
    font-size: 16px;
    border-radius: 8px;
}

@media (max-width: 600px) {

    .main-title {
        font-size: 21px !important;
        line-height: 1.3;
    }

    .sub-title {
        font-size: 13px !important;
        line-height: 1.5;
    }

    .stButton button {
        min-height: 54px !important;
        font-size: 16px !important;
    }

    .stDownloadButton button {
        min-height: 54px !important;
        font-size: 16px !important;
    }

}

</style>
""", unsafe_allow_html=True)


# =========================================================
# TITLE
# =========================================================

st.markdown(
    '<div class="main-title">🏦 GJM Bank Reports to Excel Converter</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">'
    '📱 Mobile અને Desktop બંને માટે • ZIP / TXT → Excel'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# IGNORE PATTERNS
# =========================================================

IGNORE_PATTERNS = [
    "REPORT ID:",
    "PROC DATE:",
    "RUN DATE:",
    "BRANCH :",
    "BRANCH NO.",
    "PAGE NO",
    "PRODUCT TOTAL",
    "ACCOUNT TYPE TOTAL",
    "OVER DRAWN TOTAL",
    "NO OF ACCOUNTS",
    "SUB TOTAL",
    "GRAND TOTAL",
    "BALANCE FORWARD",
    "BHAVNAGAR DISTRICT",
    "TOTAL FOR PRODUCT",
    "TOTAL NO OF",
    "PRODUCT-WISE TOTAL",
    "====>",
    "GL CLASS CODE",
    "GL-CLASS-CODE",
    "AREA:"
]


# =========================================================
# TEXT CLEANING
# =========================================================

def sanitize_text(val):

    if val is None:
        return ""

    val_str = str(val)

    return ILLEGAL_CHARACTERS_RE.sub(
        "",
        val_str
    ).strip()


def clean_dr_amount(val):

    val_str = sanitize_text(val)

    if not val_str:
        return ""

    if (
        re.search(r'\bDr\b', val_str, re.IGNORECASE)
        or val_str.endswith("Dr")
        or val_str.endswith("dr")
    ):

        num_clean = re.sub(
            r'(?i)\s*dr\s*',
            '',
            val_str
        ).strip()

        if not num_clean.startswith("-"):
            return f"-{num_clean}"

        return num_clean

    elif re.search(
        r'\bCr\b',
        val_str,
        re.IGNORECASE
    ):

        return re.sub(
            r'(?i)\s*cr\s*',
            '',
            val_str
        ).strip()

    return val_str


# =========================================================
# DEPOSITS BALANCE REPORT
# =========================================================

def parse_deposits_balance_report(lines):

    pattern = re.compile(
        r'^\s*(\d{11}-\d|\d{8,16})\s+'
        r'(\S+(?:\s\S+)*)\s{2,}'
        r'(.+?)\s{2,}'
        r'([\d,]+\.\d{2})\s+'
        r'([\d,]+\.\d{2})\s+'
        r'([\d,]+\.\d{2}(?:\s*Dr)?)\s+'
        r'([\d,]+\.\d{2})\s*'
        r'(?:\s+(\d+[DMY]))?'
        r'\s+([\d,]+\.\d{2})'
        r'\s+([A-Z]+)'
        r'\s+([YN])\s*$',
        re.IGNORECASE
    )

    rows = []

    for line in lines:

        clean_line = sanitize_text(line)

        match = pattern.match(clean_line)

        if match:

            rows.append({

                "ACCOUNT NUMBER":
                    match.group(1).strip(),

                "ACCOUNT TYPE (DESCRIPTION)":
                    match.group(2).strip(),

                "CUSTOMER NAME":
                    match.group(3).strip(),

                "AVAILABLE BALANCE":
                    match.group(4).strip(),

                "UNCLEARED BALANCE":
                    match.group(5).strip(),

                "CURRENT BALANCE":
                    clean_dr_amount(
                        match.group(6)
                    ),

                "LIMIT":
                    match.group(7).strip(),

                "TERM":
                    match.group(8).strip()
                    if match.group(8)
                    else "",

                "INT-RATE":
                    match.group(9).strip(),

                "STATUS":
                    match.group(10).strip(),

                "JOINT-HOLD-FLAG":
                    match.group(11).strip()
            })

    if rows:
        return pd.DataFrame(rows)

    return None


# =========================================================
# PIPE DELIMITED REPORT
# =========================================================

def parse_pipe_delimited(lines):

    header_cols = []
    data_rows = []

    for line in lines:

        clean_line = sanitize_text(line)

        if not clean_line:
            continue

        if re.match(
            r'^-{10,}',
            clean_line
        ):
            continue

        if any(
            keyword in clean_line.upper()
            for keyword in IGNORE_PATTERNS
        ):
            continue

        if "|" not in clean_line:
            continue

        parts = [
            clean_dr_amount(part)
            for part in clean_line.split("|")
        ]

        if parts and parts[0] == "":
            parts = parts[1:]

        if parts and parts[-1] == "":
            parts = parts[:-1]

        if (
            not header_cols
            and any(
                word in clean_line.upper()
                for word in [
                    "NAME",
                    "ACCOUNT",
                    "AMOUNT",
                    "LIMIT",
                    "DATE",
                    "PRODUCT"
                ]
            )
        ):

            header_cols = parts

        elif header_cols and len(parts) >= 2:

            if len(parts) < len(header_cols):

                parts.extend(
                    [""] *
                    (len(header_cols) - len(parts))
                )

            data_rows.append(
                parts[:len(header_cols)]
            )

    if header_cols and data_rows:

        return pd.DataFrame(
            data_rows,
            columns=header_cols
        )

    return None


# =========================================================
# GENERAL CBS REPORT
# =========================================================

def parse_general_cbs_report(lines):

    header_index = -1

    for i, line in enumerate(lines):

        clean_line = sanitize_text(line)

        if re.match(
            r'^\s*-{15,}',
            clean_line
        ):

            if i + 2 < len(lines):

                if re.match(
                    r'^\s*-{15,}',
                    sanitize_text(lines[i + 2])
                ):

                    header_index = i + 1

                    break

    if header_index == -1:
        return None

    cols = [
        match.group(1).strip()
        for match in re.finditer(
            r'(\S+(?:\s(?!\s)\S+)*)',
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

        if re.match(
            r'^-{10,}',
            clean_line
        ):
            continue

        if any(
            keyword in clean_line.upper()
            for keyword in IGNORE_PATTERNS
        ):
            continue

        if "ACCOUNT NUMBER" in clean_line.upper():
            continue

        parts = [
            part.strip()
            for part in re.split(
                r'\s{2,}',
                clean_line
            )
            if part.strip() != ""
        ]

        if len(parts) > 1:

            rows.append({
                cols[k]:
                    clean_dr_amount(parts[k])
                    if k < len(parts)
                    else ""
                for k in range(len(cols))
            })

    if rows:
        return pd.DataFrame(rows)

    return None


# =========================================================
# CSV GENERIC
# =========================================================

def parse_csv_generic(content):

    clean_content = sanitize_text(content)

    try:

        delimiter = csv.Sniffer().sniff(
            clean_content[:4096],
            delimiters=[
                ",",
                "\t",
                "|",
                ";"
            ]
        ).delimiter

    except:

        if "," in clean_content:
            delimiter = ","
        else:
            delimiter = "\t"

    df = pd.read_csv(
        io.StringIO(clean_content),
        sep=delimiter,
        engine="python",
        on_bad_lines="skip",
        encoding_errors="ignore"
    )

    for column in df.columns:

        if df[column].dtype == object:

            df[column] = df[column].apply(
                clean_dr_amount
            )

    return df


# =========================================================
# EXCEL CREATION
# =========================================================

def convert_to_excel(df):

    output = io.BytesIO()

    wb = Workbook()

    ws = wb.active

    ws.title = "Converted Report"

    clean_headers = [
        sanitize_text(column)
        for column in df.columns
    ]

    ws.append(clean_headers)

    # Header style
    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    header_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="FFFFFF"
    )

    data_font = Font(
        name="Calibri",
        size=10
    )

    border = Border(
        left=Side(
            style="thin",
            color="D9D9D9"
        ),
        right=Side(
            style="thin",
            color="D9D9D9"
        ),
        top=Side(
            style="thin",
            color="D9D9D9"
        ),
        bottom=Side(
            style="thin",
            color="D9D9D9"
        )
    )

    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    # Data
    for row in df.itertuples(
        index=False
    ):

        clean_row = [
            sanitize_text(value)
            for value in row
        ]

        ws.append(clean_row)

    # Cell formatting
    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column
    ):

        for cell in row:

            cell.font = data_font

            cell.border = border

            value_string = str(
                cell.value or ""
            ).strip()

            if re.match(
                r'^-?[\d,]+(\.\d+)?$',
                value_string
            ):

                cell.alignment = Alignment(
                    horizontal="right"
                )

            else:

                cell.alignment = Alignment(
                    horizontal="left"
                )

    # Column widths
    for column in ws.columns:

        max_length = max(
            len(
                str(
                    cell.value or ""
                )
            )
            for cell in column
        )

        column_letter = get_column_letter(
            column[0].column
        )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max(max_length + 3, 12),
            50
        )

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = ws.dimensions

    wb.save(output)

    output.seek(0)

    return output.getvalue()


# =========================================================
# PROCESS ONE TXT FILE
# =========================================================

def process_text_file(
    filename,
    content_bytes
):

    try:

        # UTF-8 first
        try:

            content_string = (
                content_bytes.decode("utf-8")
            )

        except UnicodeDecodeError:

            # Fallback
            content_string = (
                content_bytes.decode(
                    "latin1",
                    errors="ignore"
                )
            )

        lines = content_string.splitlines()

        df = None

        # -------------------------------------------------
        # 1. Deposits Balance
        # -------------------------------------------------

        upper_content = (
            content_string.upper()
        )

        if (
            "DEPOSITS BALANCE FILE"
            in upper_content
            or
            "AVAILABLE BALANCE"
            in upper_content
        ):

            df = parse_deposits_balance_report(
                lines
            )

        # -------------------------------------------------
        # 2. Pipe report
        # -------------------------------------------------

        if df is None or df.empty:

            pipe_count = sum(
                1
                for line in lines[:50]
                if line.count("|") >= 3
            )

            if pipe_count >= 3:

                df = parse_pipe_delimited(
                    lines
                )

        # -------------------------------------------------
        # 3. General CBS
        # -------------------------------------------------

        if df is None or df.empty:

            df = parse_general_cbs_report(
                lines
            )

        # -------------------------------------------------
        # 4. Generic CSV
        # -------------------------------------------------

        if df is None or df.empty:

            df = parse_csv_generic(
                content_string
            )

        # -------------------------------------------------
        # Result
        # -------------------------------------------------

        if df is None or df.empty:

            return None, "Data not found"

        excel_data = convert_to_excel(df)

        base_name = os.path.splitext(
            os.path.basename(filename)
        )[0]

        output_name = (
            f"{base_name}.xlsx"
        )

        return (
            (
                output_name,
                excel_data
            ),
            None
        )

    except Exception as error:

        return (
            None,
            str(error)
        )


# =========================================================
# FILE UPLOAD
# =========================================================

uploaded_file = st.file_uploader(
    "📁 ZIP અથવા TXT ફાઇલ પસંદ કરો",
    type=None,
    accept_multiple_files=False,
    key="bank_report_upload",
    help="ZIP અથવા TXT ફાઇલ પસંદ કરો"
)


# =========================================================
# AFTER UPLOAD
# =========================================================

if uploaded_file is not None:

    uploaded_name = uploaded_file.name

    uploaded_bytes = uploaded_file.getvalue()

    lower_name = uploaded_name.lower()

    file_items = []

    # -----------------------------------------------------
    # ZIP
    # -----------------------------------------------------

    if lower_name.endswith(".zip"):

        try:

            zip_buffer = io.BytesIO(
                uploaded_bytes
            )

            with zipfile.ZipFile(
                zip_buffer,
                "r"
            ) as zip_file:

                for info in zip_file.infolist():

                    if info.is_dir():
                        continue

                    base_name = os.path.basename(
                        info.filename
                    )

                    if not base_name:
                        continue

                    if base_name.startswith("."):
                        continue

                    # Only text/data files
                    extension = os.path.splitext(
                        base_name
                    )[1].lower()

                    if extension not in [
                        ".txt",
                        ".csv",
                        ".dat",
                        ".log"
                    ]:
                        continue

                    content = zip_file.read(
                        info
                    )

                    file_items.append(
                        (
                            base_name,
                            content
                        )
                    )

            st.success(
                f"✅ ZIP uploaded successfully"
            )

            st.info(
                f"📁 ZIPમાં {len(file_items)} "
                f"processable files મળી."
            )

        except zipfile.BadZipFile:

            st.error(
                "❌ આ valid ZIP file નથી."
            )

            st.stop()

        except Exception as error:

            st.error(
                f"❌ ZIP ખોલવામાં ભૂલ: {error}"
            )

            st.stop()

    # -----------------------------------------------------
    # TXT / CSV / DAT / LOG
    # -----------------------------------------------------

    else:

        extension = os.path.splitext(
            lower_name
        )[1]

        if extension not in [
            ".txt",
            ".csv",
            ".dat",
            ".log"
        ]:

            st.error(
                "❌ કૃપા કરીને ZIP અથવા TXT file પસંદ કરો."
            )

            st.stop()

        file_items.append(
            (
                uploaded_name,
                uploaded_bytes
            )
        )

        st.success(
            f"✅ File uploaded: {uploaded_name}"
        )

    # -----------------------------------------------------
    # FILE COUNT
    # -----------------------------------------------------

    total_files = len(file_items)

    if total_files == 0:

        st.warning(
            "⚠️ ZIPમાં કોઈ TXT/data file મળી નથી."
        )

        st.stop()

    # -----------------------------------------------------
    # CONVERT BUTTON
    # -----------------------------------------------------

    if st.button(
        "🚀 CONVERT TO EXCEL",
        type="primary",
        use_container_width=True
    ):

        progress_bar = st.progress(0)

        status_text = st.empty()

        converted_files = {}

        failed_files = []

        for index, (
            filename,
            content_bytes
        ) in enumerate(file_items):

            status_text.text(
                f"⏳ Processing: {filename} "
                f"({index + 1}/{total_files})"
            )

            result, error = process_text_file(
                filename,
                content_bytes
            )

            if result is not None:

                output_name, excel_data = result

                converted_files[
                    output_name
                ] = excel_data

            else:

                failed_files.append(
                    f"{filename}: {error}"
                )

            progress_bar.progress(
                (index + 1) / total_files
            )

        status_text.empty()

        progress_bar.empty()

        # -------------------------------------------------
        # RESULT
        # -------------------------------------------------

        st.divider()

        st.success(
            f"✅ Conversion completed: "
            f"{len(converted_files)} / "
            f"{total_files}"
        )

        # Failed files
        if failed_files:

            with st.expander(
                "⚠️ Failed / Skipped Files"
            ):

                for item in failed_files:

                    st.write(
                        f"• {item}"
                    )

        # -------------------------------------------------
        # ONE EXCEL
        # -------------------------------------------------

        if len(converted_files) == 1:

            output_name, excel_data = (
                list(
                    converted_files.items()
                )[0]
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

        # -------------------------------------------------
        # MULTIPLE EXCEL
        # -------------------------------------------------

        elif len(converted_files) > 1:

            output_zip = io.BytesIO()

            with zipfile.ZipFile(
                output_zip,
                "w",
                zipfile.ZIP_DEFLATED
            ) as result_zip:

                for (
                    filename,
                    excel_data
                ) in converted_files.items():

                    result_zip.writestr(
                        filename,
                        excel_data
                    )

            output_zip.seek(0)

            st.download_button(
                label=(
                    "📦 DOWNLOAD ALL EXCEL FILES"
                ),
                data=output_zip.getvalue(),
                file_name=(
                    "Converted_Bank_Reports.zip"
                ),
                mime="application/zip",
                use_container_width=True
            )

        else:

            st.error(
                "❌ કોઈ file Excelમાં convert થઈ શકી નથી."
            )
