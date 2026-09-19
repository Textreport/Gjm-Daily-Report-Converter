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


st.set_page_config(
    page_title="GJM Bank Report Converter",
    page_icon="🏦",
    layout="centered",
)

st.markdown("""
<style>
.main-title{text-align:center;color:#1F4E78;font-size:26px;font-weight:700;margin-bottom:5px}
.sub-title{text-align:center;color:#555;font-size:15px;margin-bottom:20px}
.stButton button,.stDownloadButton button{width:100%;min-height:52px;border-radius:8px;font-weight:700;font-size:16px}
.stDownloadButton button{background:#1F4E78;color:white}
@media(max-width:600px){
 .main-title{font-size:21px!important;line-height:1.3}
 .sub-title{font-size:13px!important}
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🏦 GJM Bank Reports to Excel Converter</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">📱 Mobile + Desktop • ZIP / TXT / GZ → Excel</div>', unsafe_allow_html=True)


IGNORE_PATTERNS = [
    "REPORT ID:", "PROC DATE:", "RUN DATE:", "BRANCH :", "BRANCH NO.",
    "PAGE NO", "PRODUCT TOTAL", "ACCOUNT TYPE TOTAL", "OVER DRAWN TOTAL",
    "NO OF ACCOUNTS", "SUB TOTAL", "GRAND TOTAL", "BALANCE FORWARD",
    "BHAVNAGAR DISTRICT", "TOTAL FOR PRODUCT", "TOTAL NO OF",
    "PRODUCT-WISE TOTAL", "====>", "GL CLASS CODE", "GL-CLASS-CODE", "AREA:"
]


def clean_text(v):
    if v is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(v)).strip()


def clean_amount(v):
    s = clean_text(v)
    if not s:
        return ""
    if re.search(r"(?i)\bdr\b$", s):
        s = re.sub(r"(?i)\s*dr\s*$", "", s).strip()
        return "-" + s if not s.startswith("-") else s
    if re.search(r"(?i)\bcr\b$", s):
        return re.sub(r"(?i)\s*cr\s*$", "", s).strip()
    return s


def parse_fixed_by_header(lines, header_line, data_test, columns):
    """
    Uses the positions of column labels in the report header.
    This is much safer than splitting on arbitrary whitespace.
    """
    starts = []
    for col in columns:
        p = header_line.upper().find(col.upper())
        if p < 0:
            return None
        starts.append(p)

    rows = []
    for line in lines:
        if not data_test(line):
            continue
        values = []
        for i, start in enumerate(starts):
            end = starts[i + 1] if i + 1 < len(starts) else len(line)
            values.append(clean_text(line[start:end]))
        if any(values):
            rows.append(values)

    return pd.DataFrame(rows, columns=columns) if rows else None


def parse_npa(lines):
    pattern = re.compile(
        r"^\s*(\d+)\s+(\d{5})\s+(\S+)\s+(\d+)\s+(\d+)\s+"
        r"(.*?)\s+"
        r"(-?[\d,]+\.\d{2})\s+"
        r"(-?[\d,]+(?:\.\d+)?)\s+"
        r"(-?[\d,]+(?:\.\d+)?)\s+"
        r"(-?[\d,]+(?:\.\d+)?)\s+"
        r"(-?[\d,]+\.\d{2})\s+"
        r"(\d{2}-[A-Z]{3}-\d{2})\s+"
        r"(\d{2}-[A-Z]{3}-\d{2})\s+"
        r"(\S+)\s+(\S+)\s+(.*)$",
        re.I
    )
    rows = []
    for line in lines:
        m = pattern.match(line)
        if m:
            rows.append([
                m.group(1),m.group(2),m.group(3),m.group(4),m.group(5),
                clean_text(m.group(6)),clean_amount(m.group(7)),
                clean_amount(m.group(8)),clean_amount(m.group(9)),
                clean_amount(m.group(10)),clean_amount(m.group(11)),
                m.group(12),m.group(13),m.group(14),m.group(15),
                clean_text(m.group(16))
            ])
    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "SR_NO","BR_NO","SYS","ACCT_NO","CUST_NO","PROD_DESCRIPTION",
        "BAL_OUTSTAND","OVERDUE_INT","INCA","UIPY","IRR_AMT",
        "LST_ARR_D","NPA_DATE","NI","OI","NAME"
    ])



def parse_probable_npa(lines):
    pattern = re.compile(
        r"^\s*(\d+)\s+(\d{8,}-\d+)\s+(.*?)\s+"
        r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+"
        r"(\d+)\s+(\d{5})\s*(.*?)$",
        re.I
    )
    rows = []
    facility_words = re.compile(
        r"\b(LOAN|CC|OD|CASH|TERM|OVERDRAFT|AGRICULTURE|HOUSING|PERSONAL)\b",
        re.I
    )
    for line in lines:
        m = pattern.match(line)
        if not m:
            continue
        before_limit = m.group(3).strip()
        fm = facility_words.search(before_limit)
        if fm:
            account_name = before_limit[:fm.start()].strip()
            facility = before_limit[fm.start():].strip()
        else:
            account_name = before_limit
            facility = ""
        tail = m.group(8)
        # Fixed-width report leaves phone columns blank; split the remaining
        # text conservatively into office phone, home phone and remarks.
        tail_parts = re.split(r"\s{2,}", tail.strip()) if tail.strip() else []
        office = tail_parts[0] if len(tail_parts) > 0 and re.fullmatch(r"\d+", tail_parts[0]) else ""
        home = tail_parts[1] if len(tail_parts) > 1 and re.fullmatch(r"\d+", tail_parts[1]) else ""
        remarks = " ".join(tail_parts[2:] if office or home else tail_parts).strip()
        rows.append([
            m.group(1),m.group(2),account_name,facility,m.group(4),
            m.group(5),m.group(6),m.group(7),office,home,remarks
        ])
    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "SR.NO.","ACCOUNT NO.","ACCOUNT-NAME","FACILITY","LIMIT",
        "OUTSTANDING","RISK-GRADE","HOME-BRCH","OFFICE-PHONE",
        "HOME-PHONE","REMARKS"
    ])



def parse_glcc_detail(lines):
    pattern = re.compile(
        r"^\s*(\d+)\s+(\d{8,14})\s+(\d{8,14})\s+"
        r"(.*?)\s+"
        r"(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+"
        r"(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+"
        r"(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})\s+"
        r"(-?[\d,]+\.\d{2})\s*$"
    )
    rows = []
    for line in lines:
        m = pattern.match(line)
        if m:
            rows.append([
                m.group(1),m.group(2),m.group(3),clean_text(m.group(4)),
                *[clean_amount(m.group(i)) for i in range(5,12)]
            ])
    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "SLNO","CUSTOMER","ACCOUNT","NAME OF ACCOUNT","DR.BALANCE",
        "CR.BALANCE","INT.BALANCE","RATE","OD-DR-INT-BAL",
        "UNCLRED-BAL","COLL-AMT"
    ])



def parse_loans_balance(lines):
    """Parse Loans Balance File using its stable 2+ whitespace field layout."""
    rows = []
    for line in lines:
        s = line.strip()
        if not re.match(r"^\d{10,16}\s", s):
            continue
        parts = [clean_text(x) for x in re.split(r"\s{2,}", s) if x.strip()]
        # Expected fields:
        # ACCOUNT NO, ACCOUNT TYPE, CUSTOMER NAME, LIMIT, INT RATE,
        # THEO BALANCE, OUTSTANDING, IRREGULARITY, SANCTION DATE,
        # EMIS DUE, EMIS PAID, EMIS OVERDUE, NEW IRAC, OLD IRAC,
        # ADV PAID AMT, ARREAR COND, CURRENCY, ACCT-MTAIN-BRCH
        if len(parts) != 18:
            continue
        rows.append(parts)

    if not rows:
        return None

    return pd.DataFrame(rows, columns=[
        "ACCOUNT NO",
        "ACCOUNT TYPE (DESCRIPTION)",
        "CUSTOMER NAME",
        "LIMIT",
        "INT RATE",
        "THEO. BALANCE",
        "OUTSTANDING",
        "IRREGULARITY",
        "SANCTION DATE",
        "EMIS DUE",
        "EMIS PAID",
        "EMIS OVERDUE",
        "NEW IRAC",
        "OLD IRAC",
        "ADV PAID AMT",
        "ARREAR COND",
        "CURRENCY",
        "ACCT-MTAIN-BRCH",
    ])


def parse_daily_productwise(lines):
    # This report is pipe-delimited. Ignore visual/group headers and totals.
    rows = []
    active_product = ""
    for line in lines:
        s = line.rstrip("\r\n")
        if "PROD CODE :" in s.upper():
            m = re.search(r"PROD CODE\s*:\s*([0-9 ]+)\s+PROD DESC\s*:\s*(.*)", s, re.I)
            if m:
                active_product = m.group(1).strip() + " - " + m.group(2).strip()
            continue

        if "|" not in s:
            continue
        if "S.NO." in s.upper() or "ACCOUNT NO'S" in s.upper():
            continue

        parts = [clean_amount(x) for x in s.split("|")]
        # Expected: SNO + ACCOUNT + 8 amount columns + optional empty final cell.
        if len(parts) < 10:
            continue

        sno = parts[0].strip()
        account = parts[1].strip()
        if not re.fullmatch(r"\d+", sno or "") or not account:
            continue

        amounts = parts[2:10]
        while len(amounts) < 8:
            amounts.append("")

        rows.append([active_product, sno, account] + amounts)

    if not rows:
        return None

    cols = [
        "PRODUCT","S.NO.","ACCOUNT NO.",
        "CASH CREDIT","CASH DEBIT",
        "CLEARING CREDIT","CLEARING DEBIT",
        "TRANSFER CREDIT","TRANSFER DEBIT",
        "PRODUCT TOTAL CREDIT","PRODUCT TOTAL DEBIT"
    ]
    return pd.DataFrame(rows, columns=cols)


def parse_pipe(lines):
    header = None
    for line in lines:
        if "|" in line and any(w in line.upper() for w in ["ACCOUNT","NAME","AMOUNT","DATE","PRODUCT"]):
            parts = [clean_text(x) for x in line.split("|")]
            if len(parts) >= 2:
                header = parts
                break
    if not header:
        return None

    rows = []
    for line in lines:
        if "|" not in line or line.strip().startswith("-"):
            continue
        parts = [clean_amount(x) for x in line.split("|")]
        if len(parts) < 2 or parts == header:
            continue
        if any(k in line.upper() for k in IGNORE_PATTERNS):
            continue
        if len(parts) < len(header):
            parts += [""] * (len(header) - len(parts))
        rows.append(parts[:len(header)])

    return pd.DataFrame(rows, columns=header) if rows else None


def parse_general(lines):
    # Fallback for reports whose columns are genuinely separated by 2+ spaces.
    for i, line in enumerate(lines):
        if re.match(r"^\s*-{15,}", line) and i + 2 < len(lines):
            if re.match(r"^\s*-{15,}", lines[i + 2]):
                header = lines[i + 1]
                parts = [x.strip() for x in re.split(r"\s{2,}", header.strip()) if x.strip()]
                if len(parts) < 2:
                    continue

                rows = []
                for data in lines[i + 3:]:
                    s = clean_text(data)
                    if not s or re.match(r"^-{10,}", s):
                        continue
                    if any(k in s.upper() for k in IGNORE_PATTERNS):
                        continue
                    p = [x.strip() for x in re.split(r"\s{2,}", s) if x.strip()]
                    if len(p) >= 2:
                        if len(p) < len(parts):
                            p += [""] * (len(parts) - len(p))
                        rows.append(p[:len(parts)])
                if rows:
                    return pd.DataFrame(rows, columns=parts)
    return None


def parse_report(text):
    lines = text.splitlines()
    u = text.upper()

    # Specific parsers first.
    if "LOANS BALANCE FILE" in u and "ACCOUNT TYPE (DESCRIPTION)" in u:
        df = parse_loans_balance(lines)
        if df is not None and not df.empty:
            return df, "LOANS_BALANCE"

    if "NPA / OVERDUE STATEMENT REPORT" in u and "BAL_OUTSTAND" in u:
        df = parse_npa(lines)
        if df is not None and not df.empty:
            return df, "NPA_STMT"

    if "NPA CLASSIFICATION REPORT" in u and "ACCOUNT-NAME" in u:
        df = parse_probable_npa(lines)
        if df is not None and not df.empty:
            return df, "PROBABLE_NPA"

    if "BALANCE IN DEPOSITS ACCOUNTS - GL-CLASS-CODE WISE - DETAIL" in u:
        df = parse_glcc_detail(lines)
        if df is not None and not df.empty:
            return df, "GLCC_WISE_DETAIL"

    if "CASH-CLG-TRANSFER TRAN CONTROL REPORT" in u:
        df = parse_daily_productwise(lines)
        if df is not None and not df.empty:
            return df, "DAILY_PRODUCTWISE"

    # Existing-style specialized deposit report.
    dep_pattern = re.compile(
        r"^\s*(\d{11}-\d|\d{8,16})\s+"
        r"(\S+(?:\s\S+)*)\s{2,}(.+?)\s{2,}"
        r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+"
        r"([\d,]+\.\d{2}(?:\s*Dr)?)\s+([\d,]+\.\d{2})\s*"
        r"(?:\s+(\d+[DMY]))?\s+([\d,]+\.\d{2})\s+([A-Z]+)\s+([YN])\s*$",
        re.I
    )
    if "AVAILABLE BALANCE" in u:
        out = []
        for line in lines:
            m = dep_pattern.match(clean_text(line))
            if m:
                out.append({
                    "ACCOUNT NUMBER":m.group(1),
                    "ACCOUNT TYPE (DESCRIPTION)":m.group(2),
                    "CUSTOMER NAME":m.group(3),
                    "AVAILABLE BALANCE":m.group(4),
                    "UNCLEARED BALANCE":m.group(5),
                    "CURRENT BALANCE":clean_amount(m.group(6)),
                    "LIMIT":m.group(7),
                    "TERM":m.group(8) or "",
                    "INT-RATE":m.group(9),
                    "STATUS":m.group(10),
                    "JOINT-HOLD-FLAG":m.group(11),
                })
        if out:
            return pd.DataFrame(out), "DEPOSITS_BALANCE"

    df = parse_pipe(lines)
    if df is not None and not df.empty:
        return df, "PIPE"

    df = parse_general(lines)
    if df is not None and not df.empty:
        return df, "GENERAL"

    # Last-resort CSV/tab parser.
    try:
        sample = clean_text(text)
        delim = csv.Sniffer().sniff(sample[:4096], delimiters=[",","\t",";"]).delimiter
        df = pd.read_csv(io.StringIO(sample), sep=delim, engine="python", on_bad_lines="skip")
        if not df.empty:
            return df, "CSV"
    except Exception:
        pass

    # Final safety net: never lose a report just because its layout is new.
    # It is exported as one clean LINE column, preserving the original text.
    raw_lines = [clean_text(x) for x in lines if clean_text(x)]
    if raw_lines:
        return pd.DataFrame({"LINE": raw_lines}), "RAW_TEXT"
    return None, "EMPTY"


def decode_bytes(data):
    for enc in ("utf-8-sig","utf-8","cp1252","latin1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("latin1", errors="replace")


def extract_files(uploaded_name, uploaded_bytes):
    items = []
    info = []

    def add(name, raw, source):
        # Only text-like files from archives are sent to the report parser.
        lower = name.lower()
        if lower.endswith(".gz"):
            raw = gzip.decompress(raw)
            name = os.path.splitext(name)[0]
        if lower.endswith((".txt",".csv",".dat",".log",".gz")) or source == "direct":
            items.append((os.path.basename(name), raw))
            info.append((name, len(raw), "OK"))

    if uploaded_name.lower().endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(uploaded_bytes), "r") as z:
            for member in z.infolist():
                if member.is_dir():
                    continue
                base = os.path.basename(member.filename)
                if not base or base.startswith("."):
                    continue
                try:
                    add(base, z.read(member), "zip")
                except Exception as e:
                    info.append((member.filename, 0, f"ERROR: {e}"))
    elif uploaded_name.lower().endswith(".gz"):
        add(uploaded_name, uploaded_bytes, "direct")
    else:
        add(uploaded_name, uploaded_bytes, "direct")

    return items, info


def make_excel(df):
    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Converted Report"

    headers = [clean_text(x) for x in df.columns]
    ws.append(headers)

    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=10)
    side = Side(style="thin", color="D9D9D9")
    border = Border(left=side,right=side,top=side,bottom=side)

    for c in ws[1]:
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in df.itertuples(index=False):
        ws.append([clean_text(v) for v in row])

    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = body_font
            c.border = border
            s = str(c.value or "").strip()
            c.alignment = Alignment(
                horizontal="right" if re.fullmatch(r"-?[\d,]+(?:\.\d+)?", s) else "left",
                vertical="center"
            )

    for col in ws.columns:
        n = min(max(max(len(str(c.value or "")) for c in col) + 3, 12), 55)
        ws.column_dimensions[get_column_letter(col[0].column)].width = n

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(out)
    return out.getvalue()


def _save_uploaded_files():
    """Copy browser-uploaded bytes into Session State immediately.
    This prevents the files from disappearing on the next Streamlit rerun
    (for example when the user presses CONVERT TO EXCEL).
    """
    selected = st.session_state.get("upload_widget", [])
    saved = []
    for f in selected:
        try:
            saved.append({
                "name": f.name,
                "data": f.getvalue(),
                "mime": getattr(f, "type", "") or "application/octet-stream",
            })
        except Exception:
            pass
    st.session_state["saved_uploads"] = saved


if "saved_uploads" not in st.session_state:
    st.session_state["saved_uploads"] = []

st.file_uploader(
    "📁 ZIP / TXT / GZ / CSV ફાઇલ પસંદ કરો",
    # No extension/MIME filter: Android file pickers can report TXT files
    # with different MIME types. Streamlit therefore accepts all files here;
    # the server-side code decides what it can read.
    type=None,
    accept_multiple_files=True,
    key="upload_widget",
    on_change=_save_uploaded_files,
    max_upload_size=500,
    help=(
        "Mobileમાં એક અથવા ઘણી files પસંદ કરી શકો છો. "
        "એકવારમાં અનેક files પસંદ ન થાય તો Upload/Browse ફરી દબાવીને "
        "વધુ files પણ ઉમેરો. ZIPમાં રહેલી TXT/TXT.GZ files પણ વાંચાશે."
    ),
)

saved_uploads = st.session_state.get("saved_uploads", [])

# A clear button is deliberately separate from the uploader. It is not
# attached to the uploader's widget state, so it safely clears our saved copy.
if saved_uploads:
    st.info(f"📦 {len(saved_uploads)} file(s) સુરક્ષિત રીતે upload થઈ ગઈ છે.")

    if st.button("🗑️ Clear uploaded files", use_container_width=True):
        st.session_state["saved_uploads"] = []
        st.rerun()

    file_items = []
    file_info = []
    upload_errors = []

    for saved in saved_uploads:
        try:
            raw_upload = saved["data"]
            filename = saved["name"]
            st.caption(
                f"📄 {filename} • {len(raw_upload)/1024/1024:.2f} MB"
            )

            items, info = extract_files(filename, raw_upload)
            file_items.extend(items)
            file_info.extend(info)
        except zipfile.BadZipFile:
            upload_errors.append(
                f"{saved['name']} — ZIP file corrupt/invalid"
            )
        except Exception as e:
            upload_errors.append(f"{saved['name']} — {e}")

    if upload_errors:
        with st.expander(
            f"⚠️ Upload errors ({len(upload_errors)})",
            expanded=True
        ):
            for x in upload_errors:
                st.write("• " + x)

    if file_items:
        st.success(
            f"✅ {len(file_items)} report files ready for conversion."
        )

        with st.expander("📋 Files ready for conversion", expanded=False):
            st.dataframe(
                pd.DataFrame(
                    file_info,
                    columns=["FILE", "SIZE BYTES", "STATUS"]
                ),
                use_container_width=True,
                hide_index=True
            )

        if st.button(
            "🚀 CONVERT TO EXCEL",
            type="primary",
            use_container_width=True
        ):
            progress = st.progress(0)
            status = st.empty()
            converted = {}
            failed = []
            parser_counts = {}

            for i, (filename, raw) in enumerate(file_items):
                status.text(
                    f"⏳ Processing {filename} "
                    f"({i+1}/{len(file_items)})"
                )
                try:
                    text = decode_bytes(raw)
                    df, parser = parse_report(text)
                    if df is None or df.empty:
                        failed.append(
                            f"{filename} — format not recognized"
                        )
                    else:
                        parser_counts[parser] = (
                            parser_counts.get(parser, 0) + 1
                        )
                        base = os.path.splitext(
                            os.path.basename(filename)
                        )[0]
                        output_name = base + ".xlsx"

                        if output_name in converted:
                            n = 2
                            while f"{base}_{n}.xlsx" in converted:
                                n += 1
                            output_name = f"{base}_{n}.xlsx"

                        converted[output_name] = make_excel(df)
                except Exception as e:
                    failed.append(f"{filename} — {e}")

                progress.progress((i + 1) / len(file_items))

            status.empty()
            progress.empty()

            st.success(
                f"✅ Conversion completed: "
                f"{len(converted)} / {len(file_items)}"
            )

            with st.expander("🔎 Parser used"):
                st.write(parser_counts)

            if failed:
                with st.expander(
                    f"⚠️ Failed / not recognized ({len(failed)})"
                ):
                    for x in failed:
                        st.write("• " + x)

            if len(converted) == 1:
                name, data = next(iter(converted.items()))
                st.download_button(
                    f"📥 DOWNLOAD {name}",
                    data=data,
                    file_name=name,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )
            elif len(converted) > 1:
                buf = io.BytesIO()
                with zipfile.ZipFile(
                    buf,
                    "w",
                    zipfile.ZIP_DEFLATED
                ) as z:
                    for name, data in converted.items():
                        z.writestr(name, data)
                buf.seek(0)

                st.download_button(
                    f"📦 DOWNLOAD ALL EXCEL FILES "
                    f"({len(converted)} FILES)",
                    data=buf.getvalue(),
                    file_name="Converted_Bank_Reports.zip",
                    mime="application/zip",
                    use_container_width=True
                )
    else:
        st.warning(
            "⚠️ Uploadમાંથી કોઈ readable TXT/ZIP/GZ file મળેલી નથી."
        )
