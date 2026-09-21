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

def _clear_uploaded_files():
    """Clear persisted uploads and reset the uploader widget state."""
    st.session_state["saved_uploads"] = []
    st.session_state["upload_message"] = ""
    # Reset the uploader on the next rerun so the just-cleared file is not re-added.
    st.session_state["mobile_upload_v3"] = None


st.markdown("### 📤 Report Upload")
st.caption("Mobile + Desktop • ZIP / TXT / GZ / CSV • એક વખતે એક file પસંદ કરો")

# Deliberately use the simplest Streamlit uploader API for maximum
# Android/iOS/browser compatibility. No form and no callback are required.
uploaded_file = st.file_uploader(
    "📁 SELECT ZIP / TXT / GZ / CSV FILE",
    type=["zip", "txt", "gz", "csv"],
    accept_multiple_files=False,
    key="mobile_upload_v3",
    help="ZIPમાં 100+ TXT reports હોય તો ZIP upload કરો.",
)

if uploaded_file is not None:
    try:
        data = uploaded_file.getvalue()
        if data:
            current_name = os.path.basename(uploaded_file.name)
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
            if not any(x.get("source_key") == f"{current_name}:{len(data)}" for x in st.session_state["saved_uploads"]):
                st.session_state["saved_uploads"].append({
                    "name": save_name,
                    "data": data,
                    "mime": getattr(uploaded_file, "type", "") or "application/octet-stream",
                    "source_key": f"{current_name}:{len(data)}",
                })
            st.success(f"✅ {save_name} સુરક્ષિત રીતે ઉમેરાઈ ગઈ છે.")
        else:
            st.warning("⚠️ પસંદ કરેલી file ખાલી છે.")
    except Exception as e:
        st.error(f"❌ Upload error: {e}")

saved_uploads = st.session_state.get("saved_uploads", [])

if saved_uploads:
    st.info(f"📦 {len(saved_uploads)} file(s) તૈયાર છે.")
    if st.button("🗑️ CLEAR ALL UPLOADS", use_container_width=True):
        st.session_state["saved_uploads"] = []
        st.rerun()



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
         "CHK1 ID", "CHK2 ID", "SUP ID"]
        if len(labels) == 11 else
        ["ACCOUNT NUMBER", "ACCOUNT NAME", "VALUE DATE", "TXN TYPE",
         "CHEQUE NO.", "AMOUNT", "USER ID", "CHK1 ID", "CHK2 ID", "SUP ID"]
    )
    return pd.DataFrame(rows, columns=cols)




# ============================================================
# UNIVERSAL BANK REPORT PARSERS
# These parsers are intentionally conservative:
# - known layouts are parsed into real columns
# - unknown layouts are never discarded; Raw Source is exported
# ============================================================

def _rows_after_header(lines, header_idx, labels, row_pattern=r"^\s*\d+\s+"):
    header = lines[header_idx]
    starts = _find_header_positions(header, labels)
    if sum(x >= 0 for x in starts) < max(3, len(labels) - 1):
        return []
    rows = []
    for line in lines[header_idx + 1:]:
        if not re.match(row_pattern, line):
            continue
        vals = []
        for j, start in enumerate(starts):
            if start < 0:
                vals.append("")
                continue
            end = starts[j + 1] if j + 1 < len(starts) and starts[j + 1] >= 0 else len(line)
            vals.append(clean_text(line[start:end]))
        if any(vals):
            rows.append(vals)
    return rows


def parse_account_open_close(lines):
    """Exact parser for LIST OF ACCOUNTS OPENED / CLOSED."""
    u = "\n".join(lines).upper()
    is_open = "LIST OF ACCOUNTS OPENED" in u
    is_closed = "LIST OF ACCOUNTS CLOSED" in u
    if not (is_open or is_closed):
        return None

    rows = []
    for line in lines:
        # Fixed-width bank line: SR + 12-digit account + 12-digit customer
        m = re.match(
            r"^\s*(\d+)\s+"
            r"(\d{10,18})\s+"
            r"(\d{10,20})\s+"
            r"(.*?)\s+"
            r"(\d{6,10})\s+"
            r"(.+?)\s+"
            r"(\d{2}/\d{2}/\d{4})\s+"
            r"([A-Z0-9]+)\s*$",
            line,
            re.I,
        )
        if not m:
            continue
        rows.append([
            m.group(1), m.group(2), m.group(3), clean_text(m.group(4)),
            m.group(5), clean_text(m.group(6)), m.group(7), m.group(8)
        ])

    if not rows:
        return None

    cols = [
        "SR NO", "ACCOUNT NUMBER", "CUSTOMER NO", "ACCOUNT NAME",
        "PRODUCT", "PRODUCT DESCRIPTION",
        "OPENED DATE" if is_open else "CLOSED DATE", "SYS"
    ]
    # Opened reports contain INT-RATE and BALANCE between product description
    # and date; parse those separately because their layout differs.
    if is_open:
        rows2 = []
        for line in lines:
            m = re.match(
                r"^\s*(\d+)\s+(\d{10,18})\s+(\d{10,20})\s+"
                r"(.*?)\s+(\d{6,10})\s+(.+?)\s+"
                r"([\d,]+\.\d+)\s+([\d,]+\.\d+)\s+"
                r"(\d{2}/\d{2}/\d{4})\s+([A-Z0-9]+)\s*$",
                line, re.I
            )
            if m:
                rows2.append([
                    m.group(1), m.group(2), m.group(3), clean_text(m.group(4)),
                    m.group(5), clean_text(m.group(6)), m.group(7),
                    m.group(8), m.group(9), m.group(10)
                ])
        if rows2:
            return pd.DataFrame(rows2, columns=[
                "SR NO", "ACCOUNT NUMBER", "CUSTOMER NO", "ACCOUNT NAME",
                "PRODUCT", "PRODUCT DESCRIPTION", "INT-RATE",
                "BALANCE", "OPENED DATE", "SYS"
            ])
    return pd.DataFrame(rows, columns=cols)


def parse_cash_transactions_universal(lines):
    """Cash report with repeated teller blocks; keeps teller number."""
    u = "\n".join(lines).upper()
    if "CASH TRANSACTIONS COMMITTED AT HOST" not in u:
        return None

    rows = []
    teller = ""
    for line in lines:
        mt = re.search(
            r"TELLER NO\s*\(MAKER OF THE TRANSACTION\)\s*:\s*(\d+)",
            line, re.I
        )
        if mt:
            teller = mt.group(1)
            continue

        m = re.match(
            r"^\s*(\d+)\s+(\d{8,18})\s+(.+?)\s+"
            r"(\d{4,8})\s+(\d{5,12})\s+"
            r"(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s*$",
            line
        )
        if m:
            rows.append([
                teller, m.group(1), m.group(2).strip(), m.group(3),
                m.group(4), m.group(5), clean_amount(m.group(6)),
                clean_amount(m.group(7))
            ])

    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "TELLER NO", "SL NO", "A/C NO", "CUSTOMER NAME",
        "TRANSACTION", "JOURNAL NO", "PAYMENTS", "RECEIPTS"
    ])


def parse_projected_tds_universal(lines):
    """Projected TDS report; carries customer number/name onto each account row."""
    u = "\n".join(lines).upper()
    if "PROJECTED TDS CUSTOMER-WISE" not in u or "TDS PROJ" not in u:
        return None

    rows = []
    customer_no = ""
    customer_name = ""
    for line in lines:
        mc = re.search(
            r"Customer No\.\s*([0-9]+).*?Name\s*:\s*(.*)$",
            line, re.I
        )
        if mc:
            customer_no = mc.group(1).strip()
            customer_name = clean_text(mc.group(2))
            continue

        m = re.match(
            r"^\s*(\d+)\s+(\d{8,18})\s+(.+?)\s+"
            r"(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s+"
            r"(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s+"
            r"(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)\s+"
            r"(-?[\d,]+\.\d+)\s+(-?[\d,]+\.\d+)"
            r"(?:\s+(.*))?$",
            line
        )
        if not m:
            continue
        rows.append([
            m.group(1), customer_no, customer_name, m.group(2),
            clean_text(m.group(3)),
            clean_amount(m.group(4)), clean_amount(m.group(5)),
            clean_amount(m.group(6)), clean_amount(m.group(7)),
            clean_amount(m.group(8)), clean_amount(m.group(9)),
            clean_amount(m.group(10)), clean_amount(m.group(11)),
            clean_text(m.group(12) or "")
        ])

    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "SR NO", "CUSTOMER NO", "CUSTOMER NAME", "ACCOUNT NO",
        "TYPE OF ACCOUNT", "INT PAID", "INT PROJ", "TDS DED", "TDS PROJ",
        "SRCHG DED", "SRCHG PROJ", "CESS DED", "CESS PROJ", "15H SUB DT"
    ])


def parse_gl_accounts_detail_universal(lines):
    """GL account detail: one row per GL account, preserving DR/CR blanks."""
    u = "\n".join(lines).upper()
    if "BALANCE IN GL" not in u or "LEDGER NAME" not in u:
        return None

    header_idx = None
    header = ""
    for i, line in enumerate(lines):
        if "SLNO" in line.upper() and "LEDGER NAME" in line.upper() and "DR.BALANCE" in line.upper():
            header_idx, header = i, line
            break
    if header_idx is None:
        return None

    labels = ["SLNO", "ACCOUNT NO", "LEDGER NAME", "CURRENCY", "DR.BALANCE", "CR.BALANCE"]
    starts = _find_header_positions(header, labels)
    if not starts or sum(x >= 0 for x in starts) < 6:
        return None

    rows = []
    current_gl_class = ""
    for line in lines[header_idx + 1:]:
        mg = re.search(r"GL CLASS CODE\s+([A-Z0-9]+)", line, re.I)
        if mg:
            current_gl_class = mg.group(1).strip()
            continue
        if not re.match(r"^\s*\d+\s+\d{8,18}\s+", line):
            continue

        # Name can contain tokens such as "FF" and "A/C", so identify
        # currency explicitly instead of assuming the token after name.
        prefix = re.match(
            r"^\s*(\d+)\s+(\d{8,18})\s+(.+?)\s+(INR|USD|EUR|GBP)\s+",
            line, re.I
        )
        if not prefix:
            continue

        sl, account, ledger, currency = prefix.groups()
        dr = clean_text(line[starts[4]:starts[5]]) if starts[4] >= 0 else ""
        cr = clean_text(line[starts[5]:]) if starts[5] >= 0 else ""
        rows.append([
            current_gl_class, sl, account, clean_text(ledger),
            currency.upper(), clean_amount(dr), clean_amount(cr)
        ])

    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "GL CLASS CODE", "SL NO", "ACCOUNT NO", "LEDGER NAME",
        "CURRENCY", "DR BALANCE", "CR BALANCE"
    ])


def parse_loan_detail_universal(lines):
    """Loan GL-class detail with customer/account and interest fields."""
    u = "\n".join(lines).upper()
    if "LOAN ACCOUNTS" not in u or "NAME OF ACCOUNT" not in u:
        return None

    header_idx = None
    header = ""
    for i, line in enumerate(lines):
        if "SLNO" in line.upper() and "NAME OF ACCOUNT" in line.upper():
            header_idx, header = i, line
            break
    if header_idx is None:
        return None

    labels = [
        "SLNO", "CUSTOMER", "ACCOUNT", "NAME OF ACCOUNT",
        "DR.BALANCE", "CR.BALANCE", "INT.BALANCE&RATE",
        "INT-YTD-CY", "INT-YTD-PY", "UNPD-INT"
    ]
    starts = _find_header_positions(header, labels)
    if not starts or sum(x >= 0 for x in starts) < 8:
        return None

    rows = []
    current_gl_class = ""
    for line in lines[header_idx + 1:]:
        mg = re.search(r"GL CLASS CODE\s+([A-Z0-9]+)", line, re.I)
        if mg:
            current_gl_class = mg.group(1).strip()
            continue
        if not re.match(r"^\s*\d+\s+\d{10,20}\s+\d{10,20}\s+", line):
            continue

        # The first three fields begin slightly before their header labels in
        # this bank's fixed-width printout. Regex is safer for those IDs.
        prefix = re.match(
            r"^\s*(\d+)\s+(\d{10,20})\s+(\d{10,20})\s+(.*?)\s*$",
            line[:starts[4]] if starts[4] > 0 else line
        )
        if not prefix:
            continue

        sl, customer, account, name = prefix.groups()
        vals = [
            sl, customer, account, clean_text(name),
            clean_text(line[starts[4]:starts[5]]),
            clean_text(line[starts[5]:starts[6]]),
            clean_text(line[starts[6]:starts[7]]),
            clean_text(line[starts[7]:starts[8]]),
            clean_text(line[starts[8]:starts[9]]),
            clean_text(line[starts[9]:]),
        ]
        vals[4:] = [clean_amount(x) for x in vals[4:]]
        rows.append([current_gl_class] + vals)

    if not rows:
        return None
    return pd.DataFrame(rows, columns=[
        "GL CLASS CODE", "SL NO", "CUSTOMER", "ACCOUNT", "NAME OF ACCOUNT",
        "DR BALANCE", "CR BALANCE", "INT BALANCE & RATE",
        "INT YTD CY", "INT YTD PY", "UNPD INT"
    ])


def parse_loan_summary_universal(lines):
    """Loan GL-class summary."""
    u = "\n".join(lines).upper()
    if "GL-CLASS-CODE WISE - SUMMARY" not in u or "TOTAL INTEREST" not in u:
        return None
    rows=[]
    for line in lines:
        m=re.match(
            r"^\s*([A-Z0-9]{10,30})\s+(\d+)\s+(.+?)\s+"
            r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$",
            line,re.I
        )
        if m:
            rows.append([
                m.group(1),m.group(2),clean_text(m.group(3)),
                m.group(4),m.group(5)
            ])
    if not rows:return None
    return pd.DataFrame(rows,columns=[
        "GL CLASS CODE","ACT TOTAL","NAME","TOTAL AMOUNT","TOTAL INTEREST"
    ])


def parse_ccod_defaulters_universal(lines):
    """CC/OD defaulters use ':' as field separator in data rows."""
    u="\n".join(lines).upper()
    if "CC/OD DEFAULTERS LIST" not in u or "DATE OF NPA" not in u:
        return None
    rows=[]
    for line in lines:
        if ":" not in line:
            continue
        # Data starts with serial number followed by ':'.
        if not re.match(r"^\s*\d+\s*:",line):
            continue
        parts=[clean_text(x) for x in line.split(":")]
        # There are 17 data fields; preserve empty/trailing fields.
        if len(parts)<17:
            parts += [""]*(17-len(parts))
        if len(parts)>17:
            parts=parts[:16]+[":".join(parts[16:])]
        parts[16] = parts[16].rstrip()
        if parts[16].endswith(":"):
            parts[16] = parts[16][:-1].rstrip()
        rows.append(parts[:17])
    if not rows:return None
    return pd.DataFrame(rows,columns=[
        "S-NO","CUSTOMER NAME","CUSTOMER ADDRESS","CIF NUMBER",
        "ACCOUNT NUMBER","PRODUCT DESC.","APPROVED AMOUNT","DRAWING POWER",
        "OUTSTANDING AMOUNT","PRINCIPAL","INTEREST","INTEREST ACCURAL",
        "PENAL INTEREST","PENAL INT ACCURAL","CHARGES","NPA ST","DATE OF NPA"
    ])


def parse_term_deposit_advices_universal(lines):
    """Term deposit maturity advice: one clean row per account advice."""
    u="\n".join(lines).upper()
    if "TERM DEPOSIT MATURING ADVICES" not in u or "ACCOUNT-STATUS" not in u:
        return None
    rows=[]
    customer_no=customer_name=""
    address=[]
    branch=""
    date=""
    i=0
    while i<len(lines):
        line=lines[i]
        mb=re.search(r"BRANCH\s*:\s*(\d+)\s+(.+?)\s*$",line,re.I)
        if mb: branch=mb.group(1)+" "+clean_text(mb.group(2))
        md=re.search(r"\bDATE\s*:\s*(\d{2}/\d{2}/\d{4})",line,re.I)
        if md: date=md.group(1)
        mc=re.match(r"^\s*(Mrs?\.?|Ms\.?)\s+(.+?)\s+\((\d+)\)\s*$",line,re.I)
        if mc:
            customer_name=clean_text(mc.group(2))
            customer_no=mc.group(3)
        if "ACCOUNT NO." in line.upper() and "ACCOUNT-STATUS" in line.upper():
            if i+2<len(lines):
                for j in range(i+1,min(i+4,len(lines))):
                    s=lines[j].strip()
                    if re.match(r"^\d{10,18}\s+\d+\s+INR\s+",s):
                        m=re.match(
                            r"^(\d{10,18})\s+(\d+)\s+(\S+)\s+"
                            r"([\d.]+)\s+([\d.]+)\s+"
                            r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s*$",s
                        )
                        if m:
                            rows.append([
                                branch, date, customer_no, customer_name,
                                m.group(1),m.group(2),m.group(3),
                                m.group(4),m.group(5),m.group(6),
                                clean_text(m.group(7))
                            ])
                        break
        i+=1
    if not rows:return None
    return pd.DataFrame(rows,columns=[
        "BRANCH","REPORT DATE","CUSTOMER NO","CUSTOMER NAME",
        "ACCOUNT NO","RECEIPT NO","CURRENCY","TERM VALUE",
        "MAT VALUE","MAT DATE","ACCOUNT STATUS"
    ])


def parse_generic_fixed_width(lines):
    """Conservative universal fixed-width fallback.

    Finds a likely header line with 3+ labelled columns and then extracts rows
    using header positions. It does not claim unknown reports are understood;
    Raw Source is always retained in the workbook.
    """
    candidates=[]
    keywords=("ACCOUNT","CUSTOMER","NAME","DATE","AMOUNT","BALANCE",
              "PRODUCT","STATUS","CODE","NUMBER","DEBIT","CREDIT","RATE")
    for i,line in enumerate(lines):
        u=line.upper()
        score=sum(1 for k in keywords if k in u)
        if score>=3 and len(line)>=40 and not set(line.strip()) <= set("-=_"):
            if not re.match(r"^\s*(REPORT ID|RUN DATE|PROC DATE|BRANCH|AREA|PAGE)",u):
                candidates.append((score,i,line))
    candidates.sort(reverse=True)
    for _,i,header in candidates[:10]:
        # Build columns from 2+ whitespace-separated header labels.
        parts=list(re.finditer(r"\S+(?:\s+\S+)*?(?=\s{2,}|\s*$)",header))
        # Prefer recognisable contiguous labels.
        labels=[m.group(0).strip() for m in parts]
        if len(labels)<3 or len(labels)>25:
            continue
        starts=[m.start() for m in parts]
        rows=[]
        for line in lines[i+1:]:
            if not re.match(r"^\s*(?:\d+|[A-Z0-9]{5,})\s+",line):
                continue
            if any(x in line.upper() for x in
                   ["TOTAL FOR","GRAND TOTAL","NO OF ACCOUNT","REPORT ID:",
                    "GL CLASS CODE"]):
                continue
            vals=[]
            for j,s in enumerate(starts):
                e=starts[j+1] if j+1<len(starts) else len(line)
                vals.append(clean_text(line[s:e]))
            if sum(bool(x) for x in vals)>=2:
                rows.append(vals)
        if rows:
            # make unique headers
            out=[]
            seen={}
            for lab in labels:
                lab=clean_text(lab) or "COLUMN"
                seen[lab]=seen.get(lab,0)+1
                out.append(lab if seen[lab]==1 else f"{lab}_{seen[lab]}")
            return pd.DataFrame(rows,columns=out)
    return None


def parse_universal_report(text):
    """Universal dispatcher. Specific parsers run first, then safe fallbacks."""
    lines=text.splitlines()
    # Exact/known layouts first.
    parsers=[
        (parse_account_open_close,"ACCOUNT_OPEN_CLOSE"),
        (parse_term_deposit_advices_universal,"TERM_DEPOSIT_ADVICES"),
        (parse_ccod_defaulters_universal,"CCOD_DEFAULTERS"),
        (parse_cash_transactions_universal,"CASH_TRANSACTIONS"),
        (parse_projected_tds_universal,"PROJECTED_TDS"),
        (parse_loan_summary_universal,"LOAN_GL_SUMMARY"),
        (parse_loan_detail_universal,"LOAN_GL_DETAIL"),
        (parse_gl_accounts_detail_universal,"GL_ACCOUNTS_DETAIL"),
        (parse_npa_bank_fixed,"NPA_BANK_FIXED"),
        (parse_safe_deposit,"SAFE_DEPOSIT"),
        (parse_least_transaction,"LEAST_TRANSACTION"),
        (parse_gl_daybook,"GL_DAYBOOK"),
        (parse_loans_balance,"LOANS_BALANCE"),
        (parse_npa,"NPA_STMT"),
        (parse_probable_npa,"PROBABLE_NPA"),
        (parse_glcc_detail,"GLCC_WISE_DETAIL"),
        (parse_daily_productwise,"DAILY_PRODUCTWISE"),
    ]
    for fn,name in parsers:
        try:
            df=fn(lines)
        except Exception:
            df=None
        if df is not None and not df.empty:
            return df,name
    # Existing generic pipe parser.
    try:
        df=parse_pipe(lines)
        if df is not None and not df.empty:
            return df,"PIPE_GENERIC"
    except Exception:
        pass
    # Colon/semicolon tables.
    try:
        # A simple colon-table detector for future reports.
        for line in lines:
            if line.count(":")>=4 and re.match(r"^\s*\d+\s*:",line):
                fields=[clean_text(x) for x in line.split(":")]
                if len(fields)>=5:
                    n=len(fields)
                    rows=[]
                    for ln in lines:
                        if re.match(r"^\s*\d+\s*:",ln):
                            p=[clean_text(x) for x in ln.split(":")]
                            if len(p)<n:p+=[""]*(n-len(p))
                            rows.append(p[:n])
                    if rows:
                        return pd.DataFrame(rows,columns=[f"FIELD {i+1}" for i in range(n)]),"COLON_GENERIC"
    except Exception:
        pass
    try:
        df=parse_general(lines)
        if df is not None and not df.empty:
            return df,"FIXED_GENERIC"
    except Exception:
        pass
    try:
        df=parse_generic_fixed_width(lines)
        if df is not None and not df.empty:
            return df,"AUTO_FIXED_GENERIC"
    except Exception:
        pass
    # CSV/tab as a last structured attempt.
    try:
        sample="\n".join(lines)
        sniff=csv.Sniffer().sniff(sample[:10000],delimiters=[",","\t",";"])
        df=pd.read_csv(io.StringIO(sample),sep=sniff.delimiter,engine="python",on_bad_lines="skip",dtype=str)
        if not df.empty:
            return df,"CSV_GENERIC"
    except Exception:
        pass
    raw_lines = [clean_text(x) for x in lines]
    if raw_lines:
        return pd.DataFrame({"LINE": raw_lines}), "RAW_ONLY"
    return None,"EMPTY"


def _safe_sheet_name(name):
    name=re.sub(r'[\[\]\:\*\?\/\\]','_',str(name))
    return (name[:31] or "Sheet")


def _excel_cell_value(v):
    if v is None:
        return ""
    return clean_text(v)


def make_excel_universal(df, raw_text, filename, parser_name):
    """Create a loss-safe workbook: Clean Data + Report Info + Raw Source."""
    out=io.BytesIO()
    wb=Workbook()
    ws=wb.active
    ws.title="Clean Data"

    # Always write headers as text and values as text to protect account IDs,
    # branch codes, receipt numbers and leading zeros.
    headers=[clean_text(x) or f"COLUMN {i+1}" for i,x in enumerate(df.columns)]
    ws.append(headers)
    for row in df.itertuples(index=False):
        ws.append([_excel_cell_value(v) for v in row])

    fill=PatternFill("solid",fgColor="1F4E78")
    font=Font(name="Calibri",size=11,bold=True,color="FFFFFF")
    body=Font(name="Calibri",size=10)
    side=Side(style="thin",color="D9D9D9")
    border=Border(left=side,right=side,top=side,bottom=side)
    for c in ws[1]:
        c.fill=fill;c.font=font
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font=body;c.border=border
            s=str(c.value or "")
            c.alignment=Alignment(
                horizontal="right" if re.fullmatch(r"-?[\d,]+(?:\.\d+)?",s) else "left",
                vertical="center"
            )
            c.number_format="@"

    for col in ws.columns:
        width=min(max(max(len(str(c.value or "")) for c in col)+3,12),55)
        ws.column_dimensions[get_column_letter(col[0].column)].width=width
    ws.freeze_panes="A2"
    ws.auto_filter.ref=ws.dimensions

    info=wb.create_sheet("Report Info")
    info_rows=[
        ("SOURCE FILE",filename),("PARSER",parser_name),
        ("ROWS EXTRACTED",str(len(df))),("COLUMNS",str(len(df.columns))),
        ("NOTE","Clean Data contains structured rows. Raw Source preserves every original line.")
    ]
    for r in info_rows: info.append(list(r))
    info["A1"].font=font; info["B1"].font=font
    info.column_dimensions["A"].width=24;info.column_dimensions["B"].width=90

    raw=wb.create_sheet("Raw Source")
    raw.append(["LINE NO","SOURCE TEXT"])
    for i,line in enumerate(raw_text.splitlines(),1):
        raw.append([str(i),clean_text(line)])
    for c in raw[1]:
        c.fill=fill;c.font=font
    raw.column_dimensions["A"].width=12
    raw.column_dimensions["B"].width=180
    raw.freeze_panes="A2"
    raw.auto_filter.ref=raw.dimensions

    wb.save(out)
    return out.getvalue()


def make_conversion_summary(results):
    wb=Workbook()
    ws=wb.active
    ws.title="Summary"
    ws.append(["FILE","PARSER","ROWS","COLUMNS","STATUS","NOTE"])
    for r in results:
        ws.append(list(r))
    for c in ws[1]:
        c.font=Font(bold=True,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor="1F4E78")
        c.alignment=Alignment(horizontal="center")
    for col in ws.columns:
        ws.column_dimensions[get_column_letter(col[0].column)].width=min(
            max(max(len(str(c.value or "")) for c in col)+3,12),50
        )
    ws.freeze_panes="A2"
    out=io.BytesIO();wb.save(out);return out.getvalue()


def parse_report(text):
    return parse_universal_report(text)

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


saved_uploads = st.session_state.get("saved_uploads", [])

if saved_uploads:
    st.success(f"📦 {len(saved_uploads)} file(s) સુરક્ષિત રીતે તૈયાર છે.")

    col1, col2 = st.columns(2)
    with col1:
        st.button(
            "🗑️ CLEAR ALL",
            use_container_width=True,
            on_click=_clear_uploaded_files,
        )
    with col2:
        st.caption("વધુ file ઉમેરવા ઉપરનું SELECT FILE ફરી વાપરો.")

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
            conversion_results = []

            for i, (filename, raw) in enumerate(file_items):
                status.text(
                    f"⏳ Processing {filename} "
                    f"({i+1}/{len(file_items)})"
                )
                try:
                    text = decode_bytes(raw)
                    df, parser = parse_report(text)
                    if df is None or df.empty:
                        failed.append(f"{filename} — empty report")
                        conversion_results.append(
                            (filename, parser, 0, 0, "EMPTY", "No text rows found.")
                        )
                    else:
                        parser_counts[parser] = (
                            parser_counts.get(parser, 0) + 1
                        )
                        note = (
                            "Raw fallback used; every source line is preserved."
                            if parser == "RAW_ONLY"
                            else "Structured data extracted; Raw Source preserved."
                        )
                        conversion_results.append(
                            (filename, parser, len(df), len(df.columns), "OK", note)
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

                        converted[output_name] = make_excel_universal(
                            df, text, filename, parser
                        )
                except Exception as e:
                    failed.append(f"{filename} — {e}")
                    conversion_results.append(
                        (filename, "ERROR", 0, 0, "ERROR", str(e))
                    )

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
                ) as zout:
                    for name, data in converted.items():
                        zout.writestr(name, data)
                    if conversion_results:
                        zout.writestr(
                            "Conversion_Summary.xlsx",
                            make_conversion_summary(conversion_results)
                        )
                st.download_button(
                    "📦 DOWNLOAD ALL EXCEL FILES (ZIP)",
                    data=buf.getvalue(),
                    file_name="Converted_Bank_Reports.zip",
                    mime="application/zip",
                    use_container_width=True
                )

