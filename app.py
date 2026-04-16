# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  PwC Brand | Source Sans Pro
#  Features:
#  - Browse existing quotations by category/vendor/service
#  - Upload new quotation → auto-extract price
#  - Compare against historical quotations
#  - Price score + similarity score
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import openpyxl
import os, re, io, zipfile, difflib

try:
    import requests
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ════════════════════════════════════════════════════════════
# PAGE CONFIG
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="IT Procurement Intelligence",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# CSS
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

html, body, [class*="css"], div, p, span, td, th,
label, button, .stMarkdown {
    font-family: 'Source Sans Pro','Helvetica Neue',
                 Arial, sans-serif !important;
}
.main .block-container {
    background-color : #F3F3F3 !important;
    padding-top      : 1.5rem;
    max-width        : 100% !important;
    padding-left     : 2rem !important;
    padding-right    : 2rem !important;
}
#MainMenu {visibility:hidden;}
footer    {visibility:hidden;}
header    {visibility:hidden;}
[data-testid="collapsedControl"] { display:none !important; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color : #2D2D2D !important;
    border-right     : 3px solid #D04A02;
    min-width        : 300px !important;
    max-width        : 300px !important;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {
    color: #F0F0F0 !important;
    font-family: 'Source Sans Pro', sans-serif !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] {
    background-color:#FFFFFF !important;
    border-radius:2px !important;
    border:1px solid #999 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    color:#2D2D2D !important;
}
section[data-testid="stSidebar"] div[data-baseweb="input"] {
    background-color:#FFFFFF !important;
    border-radius:2px !important;
}
section[data-testid="stSidebar"] div[data-baseweb="input"] input {
    color:#2D2D2D !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color:#D04A02 !important;
    border-radius:2px !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span {
    color:white !important;
}

/* KPI */
.kpi-box {
    border-radius:4px; padding:18px 10px;
    text-align:center; color:white;
    border-left:5px solid rgba(255,255,255,0.25);
}
.kpi-value {
    font-size:2.1em; font-weight:700;
    margin:0; line-height:1.1; letter-spacing:-0.5px;
}
.kpi-label {
    font-size:0.78em; font-weight:700; opacity:0.9;
    margin-top:5px; letter-spacing:0.8px; text-transform:uppercase;
}

/* Tabs */
button[data-baseweb="tab"] {
    font-weight:600 !important; font-size:0.92em !important;
    color:#7D7D7D !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color:#D04A02 !important;
    border-bottom:3px solid #D04A02 !important;
}

/* Expander */
div[data-testid="stExpander"] details summary {
    list-style:none !important;
}
div[data-testid="stExpander"] details summary::-webkit-details-marker {
    display:none !important;
}
div[data-testid="stExpander"] details summary::marker {
    display:none !important; content:"" !important;
}
div[data-testid="stExpander"] details summary p {
    font-weight:700; font-size:0.95em; color:#2D2D2D !important;
}
div[data-testid="stExpander"] details {
    border:1px solid #ddd; border-radius:4px; margin-bottom:10px;
}

/* Score cards */
.score-card {
    border-radius:4px; padding:16px 20px;
    margin-bottom:12px; border-left:5px solid #D04A02;
}
.score-high   { background:#FFF3F0; border-color:#D04A02; }
.score-medium { background:#FFF8E1; border-color:#FFB600; }
.score-low    { background:#F0FFF4; border-color:#22992E; }
.score-num {
    font-size:2.4em; font-weight:800;
    line-height:1; letter-spacing:-1px;
}
.score-label {
    font-size:0.80em; font-weight:700;
    letter-spacing:0.5px; text-transform:uppercase;
    margin-top:4px; opacity:0.75;
}

/* Comparison table */
.comp-table {
    width:100%; border-collapse:collapse; table-layout:fixed;
    font-size:0.84em; border:1px solid #e0e0e0;
}
.comp-table thead tr { background:#2D2D2D; }
.comp-table thead th {
    padding:10px 12px; text-align:left;
    font-weight:700; font-size:0.82em;
    letter-spacing:0.4px; text-transform:uppercase;
    color:white !important; border:none; word-break:break-word;
}
.comp-table tbody tr:nth-child(even) { background:#F3F3F3; }
.comp-table tbody tr:hover           { background:#FCE8DC; }
.comp-table tbody td {
    padding:9px 12px; border-bottom:1px solid #e8e8e8;
    vertical-align:middle; word-break:break-word;
    font-size:0.83em; color:#2D2D2D;
}
.comp-table th:nth-child(1),
.comp-table td:nth-child(1) { width:15%; }
.comp-table th:nth-child(2),
.comp-table td:nth-child(2) { width:12%; }
.comp-table th:nth-child(3),
.comp-table td:nth-child(3) { width:22%; }
.comp-table th:nth-child(4),
.comp-table td:nth-child(4) { width:12%; }
.comp-table th:nth-child(5),
.comp-table td:nth-child(5) { width:12%; }
.comp-table th:nth-child(6),
.comp-table td:nth-child(6) { width:14%; }
.comp-table th:nth-child(7),
.comp-table td:nth-child(7) { width:13%; }

/* Vendor badge */
.vendor-badge {
    display:inline-block; padding:3px 8px;
    border-radius:2px; color:white;
    font-size:0.78em; font-weight:700;
    white-space:nowrap; overflow:hidden;
    text-overflow:ellipsis; max-width:100%;
    box-sizing:border-box;
}

/* Upload zone */
.upload-header {
    background:#2D2D2D; color:white;
    padding:16px 20px; border-radius:4px;
    border-left:6px solid #D04A02;
    margin-bottom:16px;
}

/* Price badge */
.price-higher { color:#E0301E; font-weight:700; }
.price-lower  { color:#22992E; font-weight:700; }
.price-equal  { color:#FFB600; font-weight:700; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# PwC COLOURS
# ════════════════════════════════════════════════════════════
COLORS = [
    "#D04A02","#295477","#299D8F",
    "#FFB600","#22992E","#E0301E",
    "#EB8C00","#6E2585","#8C8C8C","#004F9F",
]

def get_color(i):
    return COLORS[i % len(COLORS)]

CFONT = dict(
    family="Source Sans Pro, Helvetica Neue, Arial",
    size=11, color="#2D2D2D")
CBG = "#F3F3F3"


# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION ENGINE
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"""
    (?:USD|EUR|GBP|JPY|SGD|MYR|THB|AUD|CAD|INR)\s?
    \d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?
    |
    (?:[\$\€\£\¥]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?
    |
    \d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?
    """,
    re.VERBOSE | re.IGNORECASE,
)

TOTAL_KW = [
    "grand total","total amount","total price","amount due",
    "net total","total cost","total value","quote total",
    "subtotal","estimated total","total",
]


def _parse_num(s):
    try:
        return float(re.sub(r"[^\d.]", "", str(s)) or "0")
    except Exception:
        return 0.0


def _fmt(val):
    try:
        return "${:,.2f}".format(float(
            re.sub(r"[^\d.]", "", str(val)) or "0"))
    except Exception:
        return str(val)


def _best_price(text):
    tl = text.lower()
    for kw in TOTAL_KW:
        idx = tl.find(kw)
        if idx == -1:
            continue
        snippet = text[max(0, idx - 20): idx + 300]
        hits    = PRICE_RE.findall(snippet)
        valid   = [h.strip() for h in hits if _parse_num(h) >= 50]
        if valid:
            return max(valid, key=_parse_num)
    # fallback: largest number in doc
    all_hits = PRICE_RE.findall(text)
    valid    = [h.strip() for h in all_hits if _parse_num(h) >= 100]
    if valid:
        return max(valid, key=_parse_num)
    return ""


def _extract_text_from_bytes(content, ext):
    text = ""
    ext  = ext.lower().strip(".")
    try:
        if ext == "pdf":
            if not PDF_OK:
                return "pdfplumber_not_installed"
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"

        elif ext in ("xlsx", "xls"):
            wb = openpyxl.load_workbook(
                io.BytesIO(content), data_only=True, read_only=True)
            rows_text = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    row_str = "  ".join(
                        str(c) for c in row if c is not None)
                    if row_str.strip():
                        rows_text.append(row_str)
            text = "\n".join(rows_text)
            wb.close()

        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml  = z.read("word/document.xml").decode(
                        "utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>", " ", xml)
                    text = re.sub(r"\s{2,}", "\n", text)

        elif ext == "pptx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                for name in z.namelist():
                    if name.startswith("ppt/slides/slide"):
                        xml = z.read(name).decode("utf-8", errors="ignore")
                        text += re.sub(r"<[^>]+>", " ", xml) + "\n"

    except Exception as e:
        return "error:{}".format(e)
    return text


def extract_price_from_bytes(content, ext):
    text  = _extract_text_from_bytes(content, ext)
    price = _best_price(text)
    return {
        "price"    : price,
        "price_num": _parse_num(price) if price else 0.0,
        "text"     : text[:5000],
    }


def extract_price_from_url(url):
    if not REQUESTS_OK:
        return {"price": "", "price_num": 0.0, "text": ""}
    try:
        r   = requests.get(url, timeout=20)
        ext = url.split("?")[0].rsplit(".", 1)[-1].lower()
        return extract_price_from_bytes(r.content, ext)
    except Exception as e:
        return {"price": "", "price_num": 0.0,
                "text": "error:{}".format(e)}


# ════════════════════════════════════════════════════════════
# SIMILARITY ENGINE
# ════════════════════════════════════════════════════════════
def similarity_score(text_a, text_b):
    """
    Returns 0-100 similarity score between two document texts.
    Uses SequenceMatcher on cleaned token sets.
    """
    def clean(t):
        t = re.sub(r"[^\w\s]", " ", t.lower())
        tokens = [w for w in t.split() if len(w) > 2]
        return " ".join(sorted(set(tokens)))

    a = clean(text_a[:3000])
    b = clean(text_b[:3000])
    if not a or not b:
        return 0
    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    return round(ratio * 100, 1)


def service_similarity(services_a, services_b):
    """
    Returns 0-100 score based on overlapping services.
    """
    if not services_a or not services_b:
        return 0.0
    sa = set([s.lower().strip() for s in services_a])
    sb = set([s.lower().strip() for s in services_b])
    if not sa or not sb:
        return 0.0
    intersection = sa.intersection(sb)
    union        = sa.union(sb)
    return round(len(intersection) / len(union) * 100, 1)


def price_score(new_price, hist_prices):
    """
    Returns score 0-100 and label.
    100 = new price is lowest
    0   = new price is highest
    """
    if not hist_prices or new_price <= 0:
        return None, "No comparison data"
    valid = [p for p in hist_prices if p > 0]
    if not valid:
        return None, "No comparison data"
    mn = min(valid)
    mx = max(valid)
    avg = sum(valid) / len(valid)
    if mx == mn:
        return 50, "Same as historical average"
    score = round((1 - (new_price - mn) / (mx - mn)) * 100, 1)
    score = max(0, min(100, score))
    pct_vs_avg = round((new_price - avg) / avg * 100, 1)
    if new_price < avg:
        label = "{}% BELOW historical average — COMPETITIVE".format(
            abs(pct_vs_avg))
    elif new_price > avg:
        label = "{}% ABOVE historical average — REVIEW NEEDED".format(
            abs(pct_vs_avg))
    else:
        label = "Matches historical average"
    return score, label


# ════════════════════════════════════════════════════════════
# HYPERLINK EXTRACTION
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path):
    link_map = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        fn_col = hr = None
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and
                        str(cell.value).strip().lower() == "file name"):
                    fn_col = cell.column
                    hr     = cell.row
                    break
            if fn_col:
                break
        if fn_col:
            for row in ws.iter_rows(
                    min_row=hr + 1,
                    min_col=fn_col, max_col=fn_col):
                cell = row[0]
                if cell.value and cell.hyperlink:
                    link_map[str(cell.value).strip()] = \
                        str(cell.hyperlink.target).strip()
        wb.close()
    except Exception as e:
        st.warning("Hyperlink warning: {}".format(e))
    return link_map


# ════════════════════════════════════════════════════════════
# DATA LOADING
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    FILE_PATH = "Master Catalog.xlsx"
    if not os.path.exists(FILE_PATH):
        st.error("File not found: {}".format(FILE_PATH))
        return None, None

    raw = pd.read_excel(FILE_PATH, engine="openpyxl", header=None)
    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if (any("category" in v for v in vals) and
                any("file" in v for v in vals)):
            header_row = i
            break
    if header_row is None:
        st.error("Could not detect header row.")
        return None, None

    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category":
            col_map["Category"]     = c
        elif "vendor" in cl or "type" in cl:
            col_map["Vendor"]       = c
        elif cl == "file name":
            col_map["File Name"]    = c
        elif cl == "file link":
            col_map["File Link"]    = c
        elif cl == "file url":
            col_map["File URL"]     = c
        elif "comment" in cl:
            col_map["Comments"]     = c
        elif "quoted" in cl or "price" in cl:
            col_map["Quoted Price"] = c

    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)

    keep = ["Category", "Vendor", "File Name", "Comments"]
    for e in ["File Link", "File URL", "Quoted Price"]:
        if e in df.columns:
            keep.append(e)
    df = df[[c for c in keep if c in df.columns]].copy()

    df = df[
        ~(df["Category"].astype(str).str.strip().isin(["", "nan"]) &
          df["Vendor"].astype(str).str.strip().isin(["", "nan"]))
    ].copy()

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df.reset_index(drop=True, inplace=True)

    hmap = extract_hyperlinks(FILE_PATH)
    df["Hyperlink"] = df["File Name"].map(hmap).fillna("")
    for fb in ["File Link", "File URL"]:
        if fb in df.columns:
            df["Hyperlink"] = df.apply(
                lambda r: r["Hyperlink"]
                if r["Hyperlink"] not in ["", "nan"]
                else r[fb], axis=1)

    def parse_svc(v):
        if not v or str(v).strip() in ["", "nan"]:
            return ["(unspecified)"]
        parts = [s.strip() for s in str(v).split("\n") if s.strip()]
        return parts or ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_svc)

    df_exp = df.explode("Services List").copy()
    df_exp.rename(columns={"Services List": "Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].str.strip()
    df_exp = df_exp[
        ~df_exp["Service"].isin(["", "(unspecified)", "nan"])
    ].reset_index(drop=True)

    return df, df_exp


# ════════════════════════════════════════════════════════════
# LOAD
# ════════════════════════════════════════════════════════════
df_master, df_exploded = load_data()
if df_master is None or df_exploded is None:
    st.stop()

vendor_color_map = {
    v: get_color(i)
    for i, v in enumerate(sorted(df_master["Vendor"].unique()))
}


# ════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════
def sb_label(txt):
    st.markdown(
        "<p style='color:#F0F0F0;font-weight:700;font-size:0.85em;"
        "margin:12px 0 4px;letter-spacing:0.5px;"
        "text-transform:uppercase'>{}</p>".format(txt),
        unsafe_allow_html=True)


def section_title(txt, caption=""):
    st.markdown(
        "<div style='font-size:0.78em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:4px'>{}</div>".format(txt),
        unsafe_allow_html=True)
    if caption:
        st.caption(caption)


def vendor_pill(v, color):
    return (
        "<span class='vendor-badge' style='background:{}'>"
        "{}</span>".format(color, v)
    )


def score_color(score):
    if score is None:
        return "#8C8C8C"
    if score >= 70:
        return "#22992E"
    if score >= 40:
        return "#FFB600"
    return "#E0301E"


def score_label_css(score):
    if score is None:
        return "score-medium"
    if score >= 70:
        return "score-low"
    if score >= 40:
        return "score-medium"
    return "score-high"


# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        "<div style='text-align:center;padding:20px 0 14px'>"
        "<div style='font-size:2em'>📋</div>"
        "<div style='font-size:1.05em;font-weight:700;color:white;"
        "margin:5px 0 2px;letter-spacing:0.5px'>IT Procurement</div>"
        "<div style='font-size:0.72em;color:#aaa;letter-spacing:1px;"
        "text-transform:uppercase'>Intelligence Dashboard</div>"
        "</div>"
        "<hr style='border-color:#D04A02;border-width:2px;"
        "margin:0 0 16px'>",
        unsafe_allow_html=True)

    sb_label("📂 Category")
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["", "nan"]
    ])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed")

    sb_label("🏢 Vendor")
    vpool = (df_master if selected_cat == "All"
             else df_master[df_master["Category"] == selected_cat])
    all_vendors = ["All"] + sorted([
        v for v in vpool["Vendor"].unique()
        if str(v).strip() not in ["", "nan"]
    ])
    selected_vendor = st.selectbox(
        "Vendor", all_vendors, label_visibility="collapsed")

    st.markdown(
        "<hr style='border-color:#555;margin:14px 0'>",
        unsafe_allow_html=True)

    d_filt = df_exploded.copy()
    if selected_cat    != "All":
        d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All":
        d_filt = d_filt[d_filt["Vendor"]   == selected_vendor]

    sb_label("🔍 Search Services")
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle…",
        label_visibility="collapsed")

    avail = sorted([s for s in d_filt["Service"].unique()
                    if str(s).strip() not in ["", "nan"]])
    if svc_search:
        avail = [s for s in avail if svc_search.lower() in s.lower()]

    sb_label("🛠 Services ({} available)".format(len(avail)))
    selected_svcs = st.multiselect(
        "Services", options=avail, default=[],
        label_visibility="collapsed")

    st.markdown(
        "<hr style='border-color:#555;margin:14px 0'>",
        unsafe_allow_html=True)
    st.markdown(
        "<p style='color:#888;font-size:0.78em;margin:2px 0'>"
        "📄 {} quotes &nbsp;|&nbsp; "
        "🛠 {} services &nbsp;|&nbsp; "
        "🏢 {} vendors</p>".format(
            len(df_master),
            df_exploded["Service"].nunique(),
            df_master["Vendor"].nunique()),
        unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div style='background:#2D2D2D;color:white;"
    "padding:20px 28px;border-radius:4px;"
    "border-left:6px solid #D04A02;margin-bottom:22px'>"
    "<div style='font-size:0.72em;font-weight:700;"
    "letter-spacing:2px;text-transform:uppercase;"
    "color:#D04A02;margin-bottom:5px'>IT Procurement Analytics</div>"
    "<h1 style='margin:0;font-size:1.45em;font-weight:700;"
    "color:white'>Procurement Intelligence Dashboard</h1>"
    "<p style='margin:6px 0 0;opacity:0.6;font-size:0.86em;'>"
    "Browse quotations · Upload new files · "
    "Compare prices · Similarity scoring"
    "</p></div>",
    unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════
k1, k2, k3, k4 = st.columns(4)


def kpi(col, val, lbl, bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div>"
        "</div>".format(bg, val, lbl),
        unsafe_allow_html=True)


kpi(k1, d_filt["File Name"].nunique(),  "Total Quotes",    "#D04A02")
kpi(k2, d_filt["Service"].nunique(),    "Unique Services", "#295477")
kpi(k3, d_filt["Vendor"].nunique(),     "Vendors",         "#299D8F")
kpi(k4, d_filt["Category"].nunique(),   "Categories",      "#2D2D2D")

st.markdown("<br>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# MAIN TABS
# ════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Analytics",
    "📋 Browse Quotations",
    "📤 Upload & Compare",
    "📄 Data Table",
])


# ════════════════════════════════════════════════════════════
# TAB 1 — ANALYTICS
# ════════════════════════════════════════════════════════════
with tab1:

    col_l, col_r = st.columns(2, gap="large")

    with col_l:
        section_title(
            "SERVICE OVERLAP ANALYSIS",
            "Orange = same service quoted by multiple vendors.")
        shared = (
            d_filt.groupby("Service")["Vendor"].nunique()
            .sort_values(ascending=False).head(20).reset_index()
        )
        shared.columns = ["Service", "Vendor Count"]
        shared["Color"] = shared["Vendor Count"].apply(
            lambda x: "#D04A02" if x > 1 else "#C0C0C0")
        fig1 = go.Figure(go.Bar(
            x=shared["Vendor Count"],
            y=shared["Service"].str[:44],
            orientation="h",
            marker_color=shared["Color"],
            marker_line_width=0,
            text=shared["Vendor Count"],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig1.update_layout(
            height=480,
            plot_bgcolor=CBG, paper_bgcolor=CBG,
            margin=dict(l=5, r=40, t=20, b=10),
            font=CFONT,
            xaxis=dict(title="Vendors", showgrid=True,
                       gridcolor="#E0E0E0", zeroline=False),
            yaxis=dict(autorange="reversed", tickfont=dict(size=9.5)),
            bargap=0.35)
        st.plotly_chart(fig1, use_container_width=True)

    with col_r:
        section_title(
            "VENDOR SERVICE COVERAGE",
            "Higher = broader vendor capability.")
        spv = (
            d_filt.groupby("Vendor")["Service"].nunique()
            .sort_values(ascending=False).reset_index()
        )
        spv.columns = ["Vendor", "Count"]
        spv["Color"] = [
            vendor_color_map.get(v, "#8C8C8C") for v in spv["Vendor"]]
        fig2 = go.Figure(go.Bar(
            x=spv["Vendor"], y=spv["Count"],
            marker_color=spv["Color"],
            marker_line_width=0,
            text=spv["Count"], textposition="outside",
            textfont=dict(size=10),
        ))
        fig2.update_layout(
            height=480,
            plot_bgcolor=CBG, paper_bgcolor=CBG,
            margin=dict(l=5, r=10, t=20, b=10),
            font=CFONT,
            yaxis=dict(title="Unique Services", showgrid=True,
                       gridcolor="#E0E0E0", zeroline=False),
            xaxis=dict(tickangle=-35, tickfont=dict(size=9.5)),
            bargap=0.35)
        st.plotly_chart(fig2, use_container_width=True)

    section_title(
        "PROCUREMENT CATEGORY DISTRIBUTION",
        "Share of quote files across categories.")
    cat_c = (
        d_filt.drop_duplicates(subset=["Category", "File Name"])
        .groupby("Category").size().reset_index()
    )
    cat_c.columns = ["Category", "Count"]
    if not cat_c.empty:
        fig3 = px.pie(
            cat_c, names="Category", values="Count",
            hole=0.50, color_discrete_sequence=COLORS)
        fig3.update_traces(
            textposition="outside", textinfo="label+percent",
            textfont_size=11, pull=[0.03] * len(cat_c))
        fig3.update_layout(
            height=400, margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor=CBG, font=CFONT,
            legend=dict(orientation="v", x=1.02, y=0.5,
                        font=dict(size=10)))
        st.plotly_chart(fig3, use_container_width=True)


# ════════════════════════════════════════════════════════════
# TAB 2 — BROWSE QUOTATIONS
# ════════════════════════════════════════════════════════════
with tab2:

    if not selected_svcs:
        st.info(
            "👈 Select one or more **services** from the sidebar "
            "to browse matching quotations.")
    else:
        d_sel = d_filt[d_filt["Service"].isin(selected_svcs)].copy()

        if d_sel.empty:
            st.warning("No results found under current filters.")
        else:
            vsmap = defaultdict(set)
            for _, r in d_sel.iterrows():
                vsmap[r["Vendor"]].add(r["Service"])

            vendors_all  = sorted([v for v, s in vsmap.items()
                                    if set(selected_svcs).issubset(s)])
            vendors_some = sorted([v for v, s in vsmap.items()
                                    if not set(selected_svcs).issubset(s)])

            if len(selected_svcs) > 1:
                if vendors_all:
                    names = " · ".join(
                        ["**{}**".format(v) for v in vendors_all])
                    st.success(
                        "✅ {} vendor(s) offer ALL {} services: {}".format(
                            len(vendors_all), len(selected_svcs), names))
                else:
                    st.warning(
                        "No single vendor covers all {} services.".format(
                            len(selected_svcs)))
                if vendors_some:
                    with st.expander(
                            "Vendors with partial coverage",
                            expanded=False):
                        for v in vendors_some:
                            cov   = vsmap[v].intersection(
                                set(selected_svcs))
                            color = vendor_color_map.get(v, "#8C8C8C")
                            st.markdown(
                                "{} covers **{}/{}**: _{}_ ".format(
                                    vendor_pill(v, color),
                                    len(cov), len(selected_svcs),
                                    ", ".join(sorted(cov))),
                                unsafe_allow_html=True)

            section_title("QUOTATION FILES — PER SERVICE")
            has_price = "Quoted Price" in d_sel.columns

            for svc in selected_svcs:
                d_svc = (
                    d_sel[d_sel["Service"] == svc]
                    .drop_duplicates(subset=["Vendor", "File Name"])
                    .sort_values("Vendor")
                )
                vc      = d_svc["Vendor"].nunique()
                s_tag   = (
                    "SHARED" if vc > 1 else "SINGLE VENDOR")

                with st.expander(
                    "{}  —  {} vendor(s) · {} file(s) · {}".format(
                        svc, vc, len(d_svc), s_tag),
                    expanded=True,
                ):
                    pills = " ".join([
                        vendor_pill(v,
                                    vendor_color_map.get(v, "#8C8C8C"))
                        for v in sorted(d_svc["Vendor"].unique())
                    ])
                    st.markdown(
                        "<div style='margin-bottom:12px'>"
                        "<b style='font-size:0.87em'>"
                        "Vendors:</b>&nbsp;&nbsp;{}</div>".format(pills),
                        unsafe_allow_html=True)

                    rows = [
                        "<table class='comp-table'><thead><tr>"
                        "<th>Vendor</th><th>Category</th>"
                        "<th>File Name</th>"
                    ]
                    if has_price:
                        rows.append("<th>Quoted Price</th>")
                    rows.append(
                        "<th>Extracted Price</th>"
                        "<th>Price Score</th>"
                        "<th>File Link</th>"
                        "</tr></thead><tbody>"
                    )

                    # Collect prices for scoring
                    hist_prices = []
                    for _, row in d_svc.iterrows():
                        qp = _parse_num(
                            str(row.get("Quoted Price", "")).strip())
                        if qp > 0:
                            hist_prices.append(qp)
                        ck = "px_{}".format(
                            str(row.get("File Name", "")).strip())
                        cached = st.session_state.get(ck)
                        if cached and cached.get("price_num", 0) > 0:
                            hist_prices.append(cached["price_num"])

                    for i, (_, row) in enumerate(d_svc.iterrows()):
                        bg    = "#ffffff" if i % 2 == 0 else "#F3F3F3"
                        color = vendor_color_map.get(
                            row["Vendor"], "#8C8C8C")
                        fname = str(row.get("File Name", "")).strip()

                        url = str(row.get("Hyperlink", "")).strip()
                        if not url or url == "nan":
                            url = str(row.get("File Link","")).strip()
                        if not url or url == "nan":
                            url = str(row.get("File URL", "")).strip()
                        if url == "nan":
                            url = ""

                        v_cell = vendor_pill(row["Vendor"], color)
                        fn_cell = (
                            "<span style='font-family:monospace;"
                            "font-size:0.79em;word-break:break-all'>"
                            "{}</span>".format(fname))

                        l_cell = (
                            "<a href='{}' target='_blank' "
                            "style='color:#D04A02;font-weight:600;"
                            "font-size:0.82em;text-decoration:none'>"
                            "Open</a>".format(url)
                            if url and url.startswith("http")
                            else "<span style='color:#bbb'>—</span>"
                        )

                        qp_str = str(
                            row.get("Quoted Price", "")).strip()
                        qp_num = _parse_num(qp_str)
                        qp_cell = (
                            "<span style='color:#22992E;font-weight:700;"
                            "font-family:monospace'>{}</span>".format(
                                _fmt(qp_str))
                            if qp_num > 0
                            else "<span style='color:#bbb'>—</span>"
                        )

                        ck     = "px_{}".format(fname)
                        cached = st.session_state.get(ck)
                        if cached and cached.get("price_num", 0) > 0:
                            ep_num  = cached["price_num"]
                            ep_cell = (
                                "<span style='color:#295477;"
                                "font-weight:700;"
                                "font-family:monospace'>{}</span>".format(
                                    _fmt(cached["price"])))
                            ref_price = ep_num
                        elif qp_num > 0:
                            ep_cell   = (
                                "<span style='color:#bbb;"
                                "font-size:0.80em'>—</span>")
                            ref_price = qp_num
                        else:
                            ep_cell   = (
                                "<span style='color:#bbb;"
                                "font-size:0.80em'>—</span>")
                            ref_price = 0

                        # Price score vs others in same service
                        others = [
                            p for p in hist_prices if p != ref_price
                        ]
                        if ref_price > 0 and others:
                            ps, _ = price_score(ref_price, others)
                            sc_color = score_color(ps)
                            ps_cell  = (
                                "<span style='color:{};font-weight:700'>"
                                "{}/100</span>".format(sc_color, ps)
                                if ps is not None
                                else "<span style='color:#bbb'>—</span>"
                            )
                        else:
                            ps_cell = (
                                "<span style='color:#bbb'>—</span>")

                        rows.append(
                            "<tr style='background:{}'>"
                            "<td>{}</td>"
                            "<td style='color:#555'>{}</td>"
                            "<td>{}</td>".format(
                                bg, v_cell,
                                row["Category"], fn_cell))
                        if has_price:
                            rows.append(
                                "<td>{}</td>".format(qp_cell))
                        rows.append(
                            "<td>{}</td>"
                            "<td>{}</td>"
                            "<td>{}</td>"
                            "</tr>".format(
                                ep_cell, ps_cell, l_cell))

                    rows.append("</tbody></table>")
                    st.markdown("".join(rows), unsafe_allow_html=True)

                    # Extract prices button
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(
                        "Extract Prices from Files — {}".format(
                            svc[:40]),
                        key="ep_{}".format(svc[:35]),
                        type="primary",
                    ):
                        prog = st.progress(0)
                        n    = len(d_svc)
                        for ki, (_, row) in enumerate(
                                d_svc.iterrows()):
                            fname = str(
                                row.get("File Name","")).strip()
                            url   = str(
                                row.get("Hyperlink","")).strip()
                            if not url or url == "nan":
                                url = str(
                                    row.get("File Link","")).strip()
                            ck = "px_{}".format(fname)
                            if (url and url.startswith("http") and
                                    st.session_state.get(ck) is None):
                                res = extract_price_from_url(url)
                                st.session_state[ck] = res
                            prog.progress((ki + 1) / n)
                        prog.empty()
                        st.rerun()

                    # Mini bar chart — prices
                    price_data = []
                    for _, row in d_svc.iterrows():
                        fname  = str(row.get("File Name","")).strip()
                        ck     = "px_{}".format(fname)
                        cached = st.session_state.get(ck)
                        qp     = _parse_num(
                            str(row.get("Quoted Price","")).strip())
                        ep     = (cached["price_num"]
                                  if cached else 0.0)
                        pval   = ep if ep > 0 else qp
                        if pval > 0:
                            price_data.append({
                                "Vendor" : row["Vendor"],
                                "File"   : fname[:30],
                                "Price"  : pval,
                                "Color"  : vendor_color_map.get(
                                    row["Vendor"], "#8C8C8C"),
                            })

                    if len(price_data) >= 2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        section_title(
                            "PRICE COMPARISON — {}".format(svc[:50]))
                        pdf = pd.DataFrame(price_data)
                        mf  = go.Figure(go.Bar(
                            x=pdf["File"],
                            y=pdf["Price"],
                            marker_color=pdf["Color"],
                            marker_line_width=0,
                            text=pdf["Price"].apply(
                                lambda x: _fmt(x)),
                            textposition="outside",
                        ))
                        mf.update_layout(
                            height=280,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(l=5, r=10, t=20, b=10),
                            font=CFONT,
                            yaxis=dict(
                                title="Price (USD)",
                                showgrid=True,
                                gridcolor="#E0E0E0",
                                zeroline=False),
                            xaxis=dict(tickangle=-20),
                            bargap=0.4)
                        st.plotly_chart(mf, use_container_width=True)


# ════════════════════════════════════════════════════════════
# TAB 3 — UPLOAD & COMPARE
# ════════════════════════════════════════════════════════════
with tab3:

    st.markdown(
        "<div class='upload-header'>"
        "<div style='font-size:0.72em;font-weight:700;"
        "letter-spacing:2px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:4px'>New Quotation Analysis</div>"
        "<div style='font-size:1.1em;font-weight:700'>Upload a new "
        "quotation file to compare against historical data</div>"
        "<div style='font-size:0.85em;opacity:0.65;margin-top:4px'>"
        "Supports PDF, XLSX, XLS, DOCX · "
        "Auto-extracts price · "
        "Scores vs historical quotations · "
        "Similarity matching"
        "</div></div>",
        unsafe_allow_html=True)

    # ── Step 1: Upload ────────────────────────────────────────
    section_title("STEP 1 — UPLOAD NEW QUOTATION FILE")
    uploaded = st.file_uploader(
        "Upload",
        type=["pdf", "xlsx", "xls", "docx"],
        label_visibility="collapsed",
        help="Upload a quotation file in PDF, Excel or Word format",
    )

    if uploaded is not None:
        content  = uploaded.read()
        ext      = uploaded.name.rsplit(".", 1)[-1].lower()
        fname_up = uploaded.name

        st.success("File uploaded: **{}** ({} KB)".format(
            fname_up, round(len(content) / 1024, 1)))

        # ── Step 2: Extract price ─────────────────────────────
        section_title("STEP 2 — EXTRACTED PRICE")

        with st.spinner("Extracting price from file…"):
            result   = extract_price_from_bytes(content, ext)
            new_price = result["price_num"]
            new_text  = result["text"]

        if new_price > 0:
            st.markdown(
                "<div class='score-card score-low'>"
                "<div style='font-size:0.75em;font-weight:700;"
                "letter-spacing:1px;text-transform:uppercase;"
                "color:#22992E'>Extracted Price</div>"
                "<div class='score-num' style='color:#22992E'>"
                "{}</div>"
                "<div class='score-label'>from {}</div>"
                "</div>".format(_fmt(new_price), fname_up),
                unsafe_allow_html=True)
        else:
            st.warning(
                "Could not extract a price automatically. "
                "Please enter it manually below.")
            manual = st.number_input(
                "Enter price manually (USD)",
                min_value=0.0, step=100.0, value=0.0)
            if manual > 0:
                new_price = manual

        # ── Step 3: Filter to compare ─────────────────────────
        section_title(
            "STEP 3 — SELECT SERVICES IN THIS QUOTATION",
            "Select what services this quotation covers "
            "to find matching historical quotes.")

        all_svcs_list = sorted([
            s for s in df_exploded["Service"].unique()
            if str(s).strip() not in ["", "nan"]
        ])
        svc_search_up = st.text_input(
            "Search services",
            placeholder="Type to filter…",
            key="svc_search_upload")
        filtered_svcs = (
            [s for s in all_svcs_list
             if svc_search_up.lower() in s.lower()]
            if svc_search_up else all_svcs_list
        )
        new_services = st.multiselect(
            "Services in this quotation",
            options=filtered_svcs,
            key="new_svcs",
            help="Select services this file covers")

        cat_filter_up = st.selectbox(
            "Filter historical quotes by category",
            options=["All"] + sorted([
                c for c in df_master["Category"].unique()
                if str(c).strip() not in ["", "nan"]]),
            key="cat_up")

        # ── Step 4: Find matches ──────────────────────────────
        section_title("STEP 4 — HISTORICAL COMPARISON")

        if not new_services and new_price <= 0:
            st.info(
                "Select services above and/or ensure a price "
                "was extracted to see comparison results.")
        else:
            # Build candidate historical rows
            if new_services:
                mask = df_exploded["Service"].isin(new_services)
                candidates = df_exploded[mask].copy()
            else:
                candidates = df_exploded.copy()

            if cat_filter_up != "All":
                candidates = candidates[
                    candidates["Category"] == cat_filter_up
                ]

            # Unique files from candidates
            cand_files = (
                candidates
                .drop_duplicates(subset=["File Name", "Vendor"])
                [["File Name","Vendor","Category",
                  "Hyperlink","Quoted Price"]]
                .copy()
            )

            if cand_files.empty:
                st.warning(
                    "No historical quotes found matching "
                    "the selected services.")
            else:
                # Collect historical prices
                hist_prices_list = []
                for _, r in cand_files.iterrows():
                    qp = _parse_num(
                        str(r.get("Quoted Price","")).strip())
                    if qp > 0:
                        hist_prices_list.append(qp)
                    ck = "px_{}".format(
                        str(r.get("File Name","")).strip())
                    cached = st.session_state.get(ck)
                    if cached and cached.get("price_num",0) > 0:
                        hist_prices_list.append(
                            cached["price_num"])

                # ── Price score ───────────────────────────────
                if new_price > 0 and hist_prices_list:
                    ps, ps_label = price_score(
                        new_price, hist_prices_list)
                    avg_hist = (sum(hist_prices_list) /
                                len(hist_prices_list))
                    mn_hist  = min(hist_prices_list)
                    mx_hist  = max(hist_prices_list)

                    sc1, sc2, sc3, sc4 = st.columns(4)

                    sc1.markdown(
                        "<div class='score-card {}'>"
                        "<div style='font-size:0.72em;font-weight:700;"
                        "letter-spacing:1px;text-transform:uppercase;"
                        "color:{}'>Price Score</div>"
                        "<div class='score-num' style='color:{}'>"
                        "{}/100</div>"
                        "<div class='score-label'>"
                        "vs {} historical quotes</div>"
                        "</div>".format(
                            score_label_css(ps),
                            score_color(ps),
                            score_color(ps),
                            ps if ps is not None else "N/A",
                            len(hist_prices_list)),
                        unsafe_allow_html=True)

                    sc2.markdown(
                        "<div class='score-card score-medium'>"
                        "<div style='font-size:0.72em;font-weight:700;"
                        "letter-spacing:1px;text-transform:uppercase;"
                        "color:#FFB600'>Your Price</div>"
                        "<div class='score-num' style='color:#D04A02'>"
                        "{}</div>"
                        "<div class='score-label'>uploaded file</div>"
                        "</div>".format(_fmt(new_price)),
                        unsafe_allow_html=True)

                    sc3.markdown(
                        "<div class='score-card score-medium'>"
                        "<div style='font-size:0.72em;font-weight:700;"
                        "letter-spacing:1px;text-transform:uppercase;"
                        "color:#FFB600'>Historical Avg</div>"
                        "<div class='score-num' style='color:#295477'>"
                        "{}</div>"
                        "<div class='score-label'>"
                        "min {} · max {}</div>"
                        "</div>".format(
                            _fmt(avg_hist),
                            _fmt(mn_hist),
                            _fmt(mx_hist)),
                        unsafe_allow_html=True)

                    sc4.markdown(
                        "<div class='score-card {}'>"
                        "<div style='font-size:0.72em;font-weight:700;"
                        "letter-spacing:1px;text-transform:uppercase;"
                        "color:{}'>Verdict</div>"
                        "<div style='font-size:0.95em;font-weight:700;"
                        "color:{};margin-top:6px'>{}</div>"
                        "</div>".format(
                            score_label_css(ps),
                            score_color(ps),
                            score_color(ps),
                            ps_label),
                        unsafe_allow_html=True)

                    st.markdown("<br>", unsafe_allow_html=True)

                    # Price comparison chart
                    section_title("PRICE POSITIONING CHART")
                    chart_data = []
                    for _, r in cand_files.iterrows():
                        fname  = str(r.get("File Name","")).strip()
                        ck     = "px_{}".format(fname)
                        cached = st.session_state.get(ck)
                        qp     = _parse_num(
                            str(r.get("Quoted Price","")).strip())
                        ep     = (cached["price_num"]
                                  if cached else 0.0)
                        pval   = ep if ep > 0 else qp
                        if pval > 0:
                            chart_data.append({
                                "Label"  : "{} / {}".format(
                                    r["Vendor"], fname[:20]),
                                "Price"  : pval,
                                "Type"   : "Historical",
                                "Color"  : vendor_color_map.get(
                                    r["Vendor"], "#8C8C8C"),
                            })

                    # Add new file
                    chart_data.append({
                        "Label" : "★ {}".format(fname_up[:25]),
                        "Price" : new_price,
                        "Type"  : "New Upload",
                        "Color" : "#D04A02",
                    })

                    cdf = pd.DataFrame(chart_data)
                    cdf = cdf.sort_values("Price")

                    cf = go.Figure()
                    cf.add_trace(go.Bar(
                        x=cdf[cdf["Type"] == "Historical"]["Label"],
                        y=cdf[cdf["Type"] == "Historical"]["Price"],
                        marker_color=cdf[
                            cdf["Type"] == "Historical"]["Color"],
                        marker_line_width=0,
                        name="Historical",
                        text=cdf[cdf["Type"] == "Historical"][
                            "Price"].apply(_fmt),
                        textposition="outside",
                    ))
                    cf.add_trace(go.Bar(
                        x=cdf[cdf["Type"] == "New Upload"]["Label"],
                        y=cdf[cdf["Type"] == "New Upload"]["Price"],
                        marker_color="#D04A02",
                        marker_line_width=0,
                        name="Your Upload",
                        text=cdf[cdf["Type"] == "New Upload"][
                            "Price"].apply(_fmt),
                        textposition="outside",
                    ))
                    # Average line
                    cf.add_hline(
                        y=avg_hist,
                        line_dash="dash",
                        line_color="#FFB600",
                        line_width=2,
                        annotation_text="Avg: {}".format(
                            _fmt(avg_hist)),
                        annotation_position="top right",
                    )
                    cf.update_layout(
                        height=380,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5, r=10, t=20, b=10),
                        font=CFONT,
                        barmode="group",
                        yaxis=dict(
                            title="Price",
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-25),
                        legend=dict(
                            orientation="h",
                            x=0, y=1.05),
                        bargap=0.25)
                    st.plotly_chart(cf, use_container_width=True)

                # ── Similarity scoring ────────────────────────
                if new_text and len(new_text) > 100:
                    section_title(
                        "DOCUMENT SIMILARITY ANALYSIS",
                        "How similar is the new file to historical "
                        "quotations based on content & services.")

                    sim_rows = []
                    for _, r in cand_files.iterrows():
                        fname  = str(r.get("File Name","")).strip()
                        url    = str(r.get("Hyperlink","")).strip()
                        if not url or url == "nan":
                            url = str(
                                r.get("File Link","")).strip()

                        # Service similarity
                        hist_svcs = list(
                            df_exploded[
                                df_exploded["File Name"] == fname
                            ]["Service"].unique()
                        )
                        svc_sim = service_similarity(
                            new_services, hist_svcs)

                        # Text similarity (if cached)
                        ck     = "px_{}".format(fname)
                        cached = st.session_state.get(ck)
                        txt_sim = (
                            similarity_score(
                                new_text, cached["text"])
                            if cached and cached.get("text")
                            else None
                        )

                        # Combined score
                        if txt_sim is not None:
                            combined = round(
                                svc_sim * 0.6 + txt_sim * 0.4, 1)
                        else:
                            combined = svc_sim

                        sim_rows.append({
                            "Vendor"    : r["Vendor"],
                            "Category"  : r["Category"],
                            "File Name" : fname,
                            "Services Overlap %": svc_sim,
                            "Text Similarity %":
                                txt_sim if txt_sim else "—",
                            "Combined Score": combined,
                            "URL"       : url,
                        })

                    if sim_rows:
                        sim_df = pd.DataFrame(sim_rows)
                        sim_df = sim_df.sort_values(
                            "Combined Score", ascending=False)

                        sim_rows_html = [
                            "<table class='comp-table'>"
                            "<thead><tr>"
                            "<th>Vendor</th>"
                            "<th>Category</th>"
                            "<th>File Name</th>"
                            "<th>Services Match %</th>"
                            "<th>Text Similarity %</th>"
                            "<th>Combined Score</th>"
                            "<th>Link</th>"
                            "</tr></thead><tbody>"
                        ]

                        for i, r in sim_df.iterrows():
                            bg    = (
                                "#ffffff" if i % 2 == 0
                                else "#F3F3F3")
                            color = vendor_color_map.get(
                                r["Vendor"], "#8C8C8C")
                            cs    = r["Combined Score"]
                            sc    = score_color(cs)

                            url_cell = (
                                "<a href='{}' target='_blank' "
                                "style='color:#D04A02;"
                                "font-weight:600;"
                                "text-decoration:none'>"
                                "Open</a>".format(r["URL"])
                                if r["URL"] and
                                r["URL"].startswith("http")
                                else "—"
                            )
                            sim_rows_html.append(
                                "<tr style='background:{}'>"
                                "<td>{}</td>"
                                "<td style='color:#555'>{}</td>"
                                "<td style='font-family:monospace;"
                                "font-size:0.79em;word-break:"
                                "break-all'>{}</td>"
                                "<td style='text-align:center;"
                                "font-weight:700;color:#295477'>"
                                "{}</td>"
                                "<td style='text-align:center;"
                                "color:#299D8F;font-weight:700'>"
                                "{}</td>"
                                "<td style='text-align:center'>"
                                "<span style='font-weight:800;"
                                "font-size:1.1em;color:{}'>"
                                "{}/100</span></td>"
                                "<td>{}</td>"
                                "</tr>".format(
                                    bg,
                                    vendor_pill(r["Vendor"], color),
                                    r["Category"],
                                    r["File Name"],
                                    r["Services Overlap %"],
                                    r["Text Similarity %"],
                                    sc, cs,
                                    url_cell)
                            )
                        sim_rows_html.append(
                            "</tbody></table>")
                        st.markdown(
                            "".join(sim_rows_html),
                            unsafe_allow_html=True)

                        # Top match callout
                        top = sim_df.iloc[0]
                        tc  = score_color(top["Combined Score"])
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(
                            "<div class='score-card {}' "
                            "style='border-color:{}'>"
                            "<div style='font-size:0.72em;"
                            "font-weight:700;letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:{}'>Best Match</div>"
                            "<div style='font-size:1.1em;"
                            "font-weight:700;color:{};margin:6px 0'>"
                            "{} — {}</div>"
                            "<div style='font-size:0.85em;"
                            "color:#555'>"
                            "Combined similarity score: "
                            "<b style='color:{}'>"
                            "{}/100</b></div>"
                            "</div>".format(
                                score_label_css(
                                    top["Combined Score"]),
                                tc, tc, tc,
                                top["Vendor"],
                                top["File Name"][:50],
                                tc,
                                top["Combined Score"]),
                            unsafe_allow_html=True)

                        # Extract texts for similarity button
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button(
                            "Extract Prices & Enable Text Similarity "
                            "for All Matched Files",
                            type="primary",
                            key="extract_sim",
                        ):
                            prog2 = st.progress(0)
                            n2    = len(cand_files)
                            for ki2, (_, r2) in enumerate(
                                    cand_files.iterrows()):
                                fn2 = str(
                                    r2.get("File Name","")).strip()
                                u2  = str(
                                    r2.get("Hyperlink","")).strip()
                                if not u2 or u2 == "nan":
                                    u2 = str(
                                        r2.get("File Link",
                                               "")).strip()
                                ck2 = "px_{}".format(fn2)
                                if (u2 and u2.startswith("http") and
                                        st.session_state.get(
                                            ck2) is None):
                                    res2 = extract_price_from_url(u2)
                                    st.session_state[ck2] = res2
                                prog2.progress((ki2 + 1) / n2)
                            prog2.empty()
                            st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 4 — DATA TABLE
# ════════════════════════════════════════════════════════════
with tab4:
    dm = df_master.copy()
    if selected_cat    != "All":
        dm = dm[dm["Category"] == selected_cat]
    if selected_vendor != "All":
        dm = dm[dm["Vendor"]   == selected_vendor]
    st.dataframe(
        dm.drop(
            columns=["Services List", "Hyperlink"],
            errors="ignore"),
        use_container_width=True,
        height=500)
