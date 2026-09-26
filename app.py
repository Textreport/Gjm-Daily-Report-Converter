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


# ------------------------------------------------------------
# MOBILE-FIRST UPLOAD BLOCK
# Keep this block immediately after the title so the uploader is
# rendered before any report parsing/conversion logic runs.
# ------------------------------------------------------------
if "saved_uploads" not in st.session_state:
    st.session_state["saved_uploads"] = []
if "upload_message" not in st.session_state:
    st.session_state["upload_message"] = ""
if "cleared_upload_signature" not in st.session_state:
    st.session_state["cleared_upload_signature"] = ""

def _clear_uploaded_files():
    """Clear saved uploads without modifying the file_uploader widget state."""
    current = st.session_state.get("mobile_upload")
    if current is not None:
        try:
            data = current.getvalue()
            st.session_state["cleared_upload_signature"] = f"{current.name}:{len(data)}"
        except Exception:
            st.session_state["cleared_upload_signature"] = ""
    st.session_state["saved_uploads"] = []
    st.session_state["upload_message"] = ""


st.markdown("### 📤 Report Upload")
st.caption("Mobile + Desktop • ZIP / TXT / GZ / CSV • એક વખતે એક file પસંદ કરો")

# Deliberately use the simplest Streamlit uploader API for maximum
# Android/iOS/browser compatibility. No form and no callback are required.
uploaded_file = st.file_uploader(
    "📁 SELECT ZIP / TXT / GZ / CSV FILE",
    type=["zip", "txt", "gz", "csv"],
    accept_multiple_files=False,
    key="mobile_upload",
    help="ZIPમાં 100+ TXT reports હોય તો ZIP upload કરો.",
)

if uploaded_file is not None:
    try:
        data = uploaded_file.getvalue()
        if data:
            current_name = os.path.basename(uploaded_file.name)
            source_signature = f"{current_name}:{len(data)}"
            # After CLEAR ALL, Streamlit may still display the old selected file.
            # Do not re-save that same selection until the user chooses another file.
            if source_signature == st.session_state.get("cleared_upload_signature", ""):
                saved_uploads = st.session_state.get("saved_uploads", [])
                existing_names = {x.get("name", "") for x in saved_uploads}
            else:
                existing_names = {x.get("name", "") for x in st.session_state["saved_uploads"]}
            save_name = current_name
            if save_name in existing_names:
                stem, ext = os.path.splitext(current_name)
                n = 2
                while f"{stem}_{n}{ext}" in existing_names:
                    n += 1
                save_name = f"{stem}_{n}{ext}"
            # Replace same widget selection in-place; duplicate additions are
            # handled by the renamed save_name above.
            if source_signature != st.session_state.get("cleared_upload_signature", ""):
                if not any(x.get("source_key") == source_signature for x in st.session_state["saved_uploads"]):
                    st.session_state["saved_uploads"].append({
                        "name": save_name,
                        "data": data,
                        "mime": getattr(uploaded_file, "type", "") or "application/octet-stream",
                        "source_key": source_signature,
                    })
                st.success(f"✅ {save_name} સુરક્ષિત રીતે ઉમેરાઈ ગઈ છે.")
        else:
            st.warning("⚠️ પસંદ કરેલી file ખાલી છે.")
    except Exception as e:
        st.error(f"❌ Upload error: {e}")

saved_uploads = st.session_state.get("saved_uploads", [])

if saved_uploads:
    st.info(f"📦 {len(saved_uploads)} file(s) તૈયાર છે.")


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


def _find_header_positions(header_line, labels):
    """Find ordered column starts while tolerating spaces/punctuation."""
    starts = []
    cursor = 0
    upper = header_line.upper()
    for label in labels:
        parts = re.findall(r"[A-Z0-9]+", label.upper())
        if not parts:
            starts.append(-1)
            continue
        pattern = r"[\s\-_/:.\&()]*".join(re.escape(x) for x in parts)
        m = re.search(pattern, upper[cursor:])
        if not m:
            starts.append(-1)
            continue
        starts.append(cursor + m.start())
        cursor += m.end()
    return starts


def parse_npa_bank_fixed(lines):
    """Parse the NPA account list format shown in the supplied screenshot."""
    header_idx = None
    header = ""
    for i, line in enumerate(lines):
        u = line.upper()
        if "ACCOUNT-NUMBER" in u and "CUSTOMER-NAME" in u and "NPA-DATE" in u:
            header_idx, header = i, line
            break
    if header_idx is None:
        return None

    labels = [
        "SR NO", "ACCOUNT-NUMBER", "CUSTOMER-NAME", "INCA", "UIPY",
        "OLD", "NEW", "NPA-DATE", "OUTSTANDING", "ARR-COND",
        "SYS", "SYS", "INT-AMT", "PRODUCT"
    ]
    starts = _find_header_positions(header, labels)
    if sum(x >= 0 for x in starts) < 10:
        return None

    rows = []
    for line in lines[header_idx + 1:]:
        if not re.match(r"^\s*\d+\s+\d", line):
            continue
        vals = []
        for i, start in enumerate(starts):
            if start < 0:
                vals.append("")
                continue
            end = starts[i + 1] if i + 1 < len(starts) and starts[i + 1] >= 0 else len(line)
            vals.append(clean_text(line[start:end]))
        if vals:
            rows.append(vals)

    if not rows or len(rows[0]) != 14:
        return None

    return pd.DataFrame(rows, columns=[
        "SR NO", "ACCOUNT NUMBER", "CUSTOMER NAME", "INCA", "UIPY",
        "OLD", "NEW", "NPA DATE", "OUTSTANDING", "ARR-COND",
        "SYS 1", "SYS 2", "INT-AMT", "PRODUCT"
    ])


def parse_safe_deposit(lines):
    """Parse LIST OF SAFE DEPOSIT VAULT (LOCKERS) report."""
    u = "\n".join(lines).upper()
    if "SAFE DEPOSIT VAULT" not in u and "LOCKERS" not in u:
        return None

    rows = []
    for line in lines:
        if "|" not in line:
            continue
        parts = [clean_text(x) for x in line.strip().strip("|").split("|")]
        if len(parts) < 5:
            continue
        if not re.fullmatch(r"\d+", parts[0] or ""):
            continue
        while len(parts) < 6:
            parts.append("")
        rows.append(parts[:6])

    if not rows:
        return None

    return pd.DataFrame(rows, columns=[
        "SR NO", "CABINET ID", "LOCKER ID", "KEY STATUS",
        "LOCKER TYPE", "CLOSURE DATE"
    ])


def parse_least_transaction(lines):
    """Parse IDs WITH LEAST TRANSACTION VOLUME report."""
    u = "\n".join(lines).upper()
    if "LEAST TRANSACTION VOLUME" not in u or "MEMO-HITS" not in u:
        return None

    rows = []
    for line in lines:
        s = line.strip()
        m = re.match(r"^(\d+)\s+(\d{5,12})\s+(-?\d+)\s+(-?\d+)(?:\s+(.*))?$", s)
        if not m:
            continue
        rows.append([
            m.group(1), m.group(2), m.group(3), m.group(4),
            clean_text(m.group(5) or "")
        ])

    if not rows:
        return None

    return pd.DataFrame(rows, columns=[
        "SL NO", "ID", "DEBIT", "CREDIT", "MEMO-HITS"
    ])


def parse_gl_daybook(lines):
    """Parse BGL Voucher Verification / GL Day-Book transaction rows."""
    u = "\n".join(lines).upper()
    if "BGL VOUCHER VERIFICATION" not in u and "GL DAY-BOOK" not in u:
        return None

    header_idx = None
    header = ""
    for i, line in enumerate(lines):
        uu = line.upper()
        if "ACCOUNT NUMBER" in uu and "ACCOUNT NAME" in uu and "VALUE DATE" in uu and "AMOUNT" in uu:
            header_idx, header = i, line
            break
    if header_idx is None:
        return None

    labels = [
        "ACCOUNT NUMBER", "ACCOUNT NAME", "HOME", "VALUE DATE",
        "TXN TYPE", "CHEQUE NO.", "AMOUNT", "USER ID",
        "CHK1 ID", "CHK2 ID", "SUP ID"
    ]
    starts = _find_header_positions(header, labels)

    if sum(x >= 0 for x in starts) < 7:
        labels = [
            "ACCOUNT NUMBER", "ACCOUNT NAME", "VALUE DATE",
            "TXN TYPE", "CHEQUE NO.", "AMOUNT", "USER ID",
            "CHK1 ID", "CHK2 ID", "SUP ID"
        ]
        starts = _find_header_positions(header, labels)

    if sum(x >= 0 for x in starts) < 6:
        return None

    rows = []
    for line in lines[header_idx + 1:]:
        s = line.rstrip()
        if not re.match(r"^\s*\d{8,18}\b", s):
            continue
        if "ACCOUNT TOTAL" in s.upper():
            continue
        vals = []
        for i, start in enumerate(starts):
            end = starts[i + 1] if i + 1 < len(starts) and starts[i + 1] >= 0 else len(s)
            vals.append(clean_text(s[start:end]))
        if any(vals):
            rows.append(vals)

    if not rows:
        return None

    cols = (
        ["ACCOUNT NUMBER", "ACCOUNT NAME", "HOME", "VALUE DATE",
         "TXN TYPE", "CHEQUE NO.", "AMOUNT", "USER ID",
         "CHK1 ID", "CHK2 ID", "SUP I
