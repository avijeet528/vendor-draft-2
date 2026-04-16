# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  Refined: No overlaps, dummy prices, simple service select,
#  demo upload file, full scoring pipeline
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import os, re, io, zipfile, difflib, random, json

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

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
# CSS — all overlap fixes applied
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
    background-color: #F3F3F3 !important;
    padding-top: 1.2rem;
    max-width: 100% !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
}
#MainMenu {visibility:hidden;}
footer    {visibility:hidden;}
header    {visibility:hidden;}
[data-testid="collapsedControl"] { display:none !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background-color: #2D2D2D !important;
    border-right: 3px solid #D04A02;
    min-width: 300px !important;
    max-width: 300px !important;
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

/* ── KPI boxes — fix overlap ── */
.kpi-box {
    border-radius: 4px;
    padding: 16px 12px;
    text-align: center;
    color: white;
    border-left: 5px solid rgba(255,255,255,0.25);
    margin-bottom: 8px;
    overflow: hidden;
    min-height: 85px;
    display: flex;
    flex-direction: column;
    justify-content: center;
}
.kpi-value {
    font-size: 1.8em;
    font-weight: 700;
    margin: 0;
    line-height: 1.2;
    letter-spacing: -0.5px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.kpi-label {
    font-size: 0.72em;
    font-weight: 700;
    opacity: 0.9;
    margin-top: 4px;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

/* ── Tabs ── */
button[data-baseweb="tab"] {
    font-weight: 600 !important;
    font-size: 0.90em !important;
    color: #7D7D7D !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #D04A02 !important;
    border-bottom: 3px solid #D04A02 !important;
}

/* ── Expander — remove ALL arrow overlap ── */
div[data-testid="stExpander"] details > summary {
    list-style: none !important;
    padding-left: 14px !important;
    cursor: pointer;
}
div[data-testid="stExpander"] details > summary::before,
div[data-testid="stExpander"] details > summary::after,
div[data-testid="stExpander"] details > summary::-webkit-details-marker,
div[data-testid="stExpander"] details > summary::marker {
    display: none !important;
    content: "" !important;
    width: 0 !important;
    height: 0 !important;
    font-size: 0 !important;
}
div[data-testid="stExpander"] details summary svg {
    display: none !important;
}
div[data-testid="stExpander"] details summary p {
    font-weight: 700;
    font-size: 0.90em;
    color: #2D2D2D !important;
    padding-left: 0 !important;
    margin-left: 0 !important;
    line-height: 1.4;
}
div[data-testid="stExpander"] details {
    border: 1px solid #ddd;
    border-radius: 4px;
    margin-bottom: 8px;
    padding: 2px 0;
}

/* ── Comparison table ── */
.comp-table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    font-size: 0.82em;
    border: 1px solid #e0e0e0;
}
.comp-table thead tr { background: #2D2D2D; }
.comp-table thead th {
    padding: 10px 8px;
    text-align: left;
    font-weight: 700;
    font-size: 0.78em;
    letter-spacing: 0.4px;
    text-transform: uppercase;
    color: white !important;
    border: none;
    word-break: break-word;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.comp-table tbody tr:nth-child(even) { background: #F9F9F9; }
.comp-table tbody tr:hover { background: #FCE8DC; }
.comp-table tbody td {
    padding: 8px 8px;
    border-bottom: 1px solid #e8e8e8;
    vertical-align: middle;
    word-break: break-word;
    font-size: 0.82em;
    color: #2D2D2D;
}

/* ── Vendor badge ── */
.vendor-badge {
    display: inline-block;
    padding: 3px 8px;
    border-radius: 2px;
    color: white;
    font-size: 0.76em;
    font-weight: 700;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 100%;
    box-sizing: border-box;
}

/* ── Score cards — no overlap ── */
.score-card {
    border-radius: 4px;
    padding: 14px 14px;
    margin-bottom: 8px;
    border-left: 5px solid #D04A02;
    overflow: hidden;
}
.score-high   { background: #FFF3F0; border-color: #E0301E; }
.score-medium { background: #FFF8E1; border-color: #FFB600; }
.score-low    { background: #F0FFF4; border-color: #22992E; }
.score-num {
    font-size: 1.9em;
    font-weight: 800;
    line-height: 1.1;
    letter-spacing: -1px;
    white-space: nowrap;
}
.score-label {
    font-size: 0.70em;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 4px;
}

/* ── AI box ── */
.ai-box {
    background: #F8F0FF;
    border-left: 5px solid #6E2585;
    border-radius: 4px;
    padding: 14px 18px;
    margin: 10px 0;
}
.ai-box-title {
    font-size: 0.72em;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #6E2585;
    margin-bottom: 6px;
}

/* ── Insight card ── */
.insight-card {
    background: #FFFFFF;
    border-radius: 4px;
    padding: 14px 16px;
    margin-bottom: 10px;
    border: 1px solid #e0e0e0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}

/* ── Section header ── */
.section-hdr {
    font-size: 0.76em;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #D04A02;
    margin-bottom: 6px;
    margin-top: 4px;
}

/* ── Dark header block ── */
.dark-header {
    background: #2D2D2D;
    color: white;
    padding: 18px 24px;
    border-radius: 4px;
    border-left: 6px solid #D04A02;
    margin-bottom: 20px;
    overflow: hidden;
}
.dark-header h1 {
    margin: 0;
    font-size: 1.35em;
    font-weight: 700;
    color: white;
    line-height: 1.3;
}
.dark-header .subtitle {
    margin: 5px 0 0;
    opacity: 0.6;
    font-size: 0.83em;
    line-height: 1.4;
}
.dark-header .tag {
    font-size: 0.70em;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #D04A02;
    margin-bottom: 4px;
}

/* ── Catalog step ── */
.catalog-step {
    background: white;
    border-left: 4px solid #D04A02;
    border-radius: 4px;
    padding: 14px 18px;
    margin-bottom: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.catalog-step-num {
    font-size: 0.70em;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: #D04A02;
    margin-bottom: 3px;
}

/* ── Verdict badges ── */
.verdict-good {
    background: #E8F5E9;
    color: #22992E;
    padding: 3px 10px;
    border-radius: 2px;
    font-weight: 700;
    font-size: 0.80em;
    display: inline-block;
}
.verdict-avg {
    background: #FFF8E1;
    color: #E6A000;
    padding: 3px 10px;
    border-radius: 2px;
    font-weight: 700;
    font-size: 0.80em;
    display: inline-block;
}
.verdict-bad {
    background: #FFEBEE;
    color: #E0301E;
    padding: 3px 10px;
    border-radius: 2px;
    font-weight: 700;
    font-size: 0.80em;
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# PwC COLOURS
# ════════════════════════════════════════════════════════════
COLORS = [
    "#D04A02", "#295477", "#299D8F",
    "#FFB600", "#22992E", "#E0301E",
    "#EB8C00", "#6E2585", "#8C8C8C", "#004F9F",
]


def get_color(i):
    return COLORS[i % len(COLORS)]


CFONT = dict(family="Source Sans Pro, Helvetica Neue, Arial", size=11, color="#2D2D2D")
CBG = "#F3F3F3"


# ════════════════════════════════════════════════════════════
# PRICE HELPERS
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"(?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}(?:[,]\d{3})*(?:\.\d{1,2})?"
    r"|(?:[\$\€\£]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?"
    r"|\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?",
    re.IGNORECASE,
)
TOTAL_KW = [
    "grand total", "total amount", "total price", "amount due",
    "net total", "total cost", "total value", "quote total",
    "subtotal", "estimated total", "total",
]


def _parse_num(s):
    try:
        return float(re.sub(r"[^\d.]", "", str(s)) or "0")
    except Exception:
        return 0.0


def _fmt(val):
    try:
        v = float(re.sub(r"[^\d.]", "", str(val)) or "0")
        if v <= 0:
            return "—"
        return "${:,.2f}".format(v)
    except Exception:
        return str(val)


def _best_price(text):
    tl = text.lower()
    for kw in TOTAL_KW:
        idx = tl.find(kw)
        if idx == -1:
            continue
        snippet = text[max(0, idx - 20): idx + 300]
        hits = PRICE_RE.findall(snippet)
        valid = [h.strip() for h in hits if _parse_num(h) >= 50]
        if valid:
            return max(valid, key=_parse_num)
    all_hits = PRICE_RE.findall(text)
    valid = [h.strip() for h in all_hits if _parse_num(h) >= 100]
    if valid:
        return max(valid, key=_parse_num)
    return ""


def _text_from_bytes(content, ext):
    text = ""
    ext = ext.lower().strip(".")
    try:
        if ext == "pdf" and PDF_OK:
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
        elif ext in ("xlsx", "xls") and OPENPYXL_OK:
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rs = "  ".join(str(c) for c in row if c is not None)
                    if rs.strip():
                        text += rs + "\n"
            wb.close()
        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>", " ", xml)
                    text = re.sub(r"\s{2,}", "\n", text)
    except Exception:
        pass
    return text


def extract_price_from_bytes(content, ext):
    text = _text_from_bytes(content, ext)
    price = _best_price(text)
    return {
        "price": price,
        "price_num": _parse_num(price) if price else 0.0,
        "text": text[:5000],
    }


# ════════════════════════════════════════════════════════════
# SCORING
# ════════════════════════════════════════════════════════════
def similarity_score(text_a, text_b):
    def clean(t):
        t = re.sub(r"[^\w\s]", " ", t.lower())
        tokens = [w for w in t.split() if len(w) > 2]
        return " ".join(sorted(set(tokens)))
    a, b = clean(text_a[:3000]), clean(text_b[:3000])
    if not a or not b:
        return 0
    return round(difflib.SequenceMatcher(None, a, b).ratio() * 100, 1)


def service_similarity(sa, sb):
    if not sa or not sb:
        return 0.0
    a = set(s.lower().strip() for s in sa)
    b = set(s.lower().strip() for s in sb)
    if not a or not b:
        return 0.0
    return round(len(a & b) / len(a | b) * 100, 1)


def price_score(new_price, hist_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0:
        return None, "No data"
    mn, mx = min(valid), max(valid)
    avg = sum(valid) / len(valid)
    if mx == mn:
        return 50, "Same as average"
    score = round((1 - (new_price - mn) / (mx - mn)) * 100, 1)
    score = max(0, min(100, score))
    pct = round((new_price - avg) / avg * 100, 1)
    if new_price < avg:
        label = "{}% BELOW avg — COMPETITIVE".format(abs(pct))
    elif new_price > avg:
        label = "{}% ABOVE avg — REVIEW".format(abs(pct))
    else:
        label = "Matches average"
    return score, label


def score_color(s):
    if s is None:
        return "#8C8C8C"
    if s >= 70:
        return "#22992E"
    if s >= 40:
        return "#FFB600"
    return "#E0301E"


def score_css(s):
    if s is None:
        return "score-medium"
    if s >= 70:
        return "score-low"
    if s >= 40:
        return "score-medium"
    return "score-high"


def verdict_html(s):
    if s is None:
        return "<span style='color:#bbb'>—</span>"
    if s >= 70:
        return "<span class='verdict-good'>✓ Competitive</span>"
    if s >= 40:
        return "<span class='verdict-avg'>~ Average</span>"
    return "<span class='verdict-bad'>✗ Expensive</span>"


# ════════════════════════════════════════════════════════════
# AI INSIGHTS
# ════════════════════════════════════════════════════════════
def ai_vendor_insight(vendor, services, prices):
    n_svc = len(services)
    n_prc = len([p for p in prices if p > 0])
    lines = ["**{}** offers **{}** service(s).".format(vendor, n_svc)]
    if n_prc > 0:
        avg_p = sum(p for p in prices if p > 0) / n_prc
        lines.append("Average quoted price: **{}**.".format(_fmt(avg_p)))
    if n_svc >= 5:
        lines.append("Broad coverage — suitable for consolidated procurement.")
    elif n_svc >= 2:
        lines.append("Moderate coverage — consider combining vendors.")
    else:
        lines.append("Limited coverage — specialist supplier.")
    return " ".join(lines)


def ai_price_insight(new_price, hist_prices, vendor_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0:
        return "Insufficient data for price analysis."
    avg = sum(valid) / len(valid)
    mn, mx = min(valid), max(valid)
    pct = round((new_price - avg) / avg * 100, 1)
    lines = []
    if new_price <= mn:
        lines.append("This quote is the **lowest price** seen — excellent value.")
    elif new_price >= mx:
        lines.append("This quote is **above all historical prices** — negotiate strongly.")
    elif pct > 15:
        lines.append("Quote is **{}% above** average. Request revised quote.".format(abs(pct)))
    elif pct < -15:
        lines.append("Quote is **{}% below** average — very competitive.".format(abs(pct)))
    else:
        lines.append("Quote is **within normal range** ({}% vs avg).".format(pct))
    if vendor_prices:
        best_v = min(vendor_prices, key=vendor_prices.get)
        lines.append("**{}** has historically offered the lowest prices.".format(best_v))
    return " ".join(lines)


def ai_service_summary(services_by_vendor):
    if not services_by_vendor:
        return "No vendor data available."
    best = max(services_by_vendor, key=lambda v: len(services_by_vendor[v]))
    n_best = len(services_by_vendor[best])
    total = len(set(s for svcs in services_by_vendor.values() for s in svcs))
    lines = ["**{}** covers the most services ({} of {} total).".format(best, n_best, total)]
    shared = [s for s in set(s for svcs in services_by_vendor.values() for s in svcs)
              if sum(1 for svcs in services_by_vendor.values() if s in svcs) > 1]
    if shared:
        lines.append("**{}** service(s) offered by multiple vendors — ideal for benchmarking.".format(len(shared)))
    return " ".join(lines)


def ai_analyze_catalog(df, df_exp):
    insights = {}
    n_vendors = df["Vendor"].nunique()
    n_files = df["File Name"].nunique()
    n_cats = df["Category"].nunique()
    n_services = df_exp["Service"].nunique()
    insights["overview"] = (
        "Catalog contains **{} vendors**, **{} quote files**, "
        "**{} categories** and **{} unique services**.".format(n_vendors, n_files, n_cats, n_services))
    spv = df_exp.groupby("Vendor")["Service"].nunique().sort_values(ascending=False)
    if not spv.empty:
        insights["top_vendor"] = "**{}** leads with **{}** unique services.".format(spv.index[0], spv.iloc[0])
    shared = df_exp.groupby("Service")["Vendor"].nunique().sort_values(ascending=False)
    hot = shared[shared > 1]
    if not hot.empty:
        insights["competitive"] = "**{}** is the most competitive service with **{}** vendors quoting.".format(
            hot.index[0], hot.iloc[0])
    cat_counts = df.drop_duplicates(subset=["Category", "File Name"]).groupby("Category").size().sort_values(ascending=False)
    if not cat_counts.empty:
        pct = round(cat_counts.iloc[0] / cat_counts.sum() * 100, 1)
        insights["category"] = "**{}** dominates with **{}%** of quote files.".format(cat_counts.index[0], pct)
    if "Quoted Price" in df.columns:
        prices = df["Quoted Price"].apply(_parse_num)
        prices = prices[prices > 0]
        if not prices.empty:
            insights["pricing"] = "Prices range **{}** to **{}**, average **{}**.".format(
                _fmt(prices.min()), _fmt(prices.max()), _fmt(prices.mean()))
    recs = []
    if not hot.empty and len(hot) >= 2:
        recs.append("Run competitive bids on {} multi-vendor services.".format(len(hot)))
    if n_vendors >= 3:
        recs.append("Consider vendor consolidation — {} vendors may create overhead.".format(n_vendors))
    insights["recommendations"] = recs
    return insights


# ════════════════════════════════════════════════════════════
# DUMMY DATA GENERATOR
# ════════════════════════════════════════════════════════════
def generate_dummy_catalog():
    """Generates a realistic dummy master catalog with prices."""
    random.seed(42)
    categories = {
        "Network Infrastructure": {
            "vendors": ["Cisco Systems", "Juniper Networks", "Arista Networks", "HPE Networking"],
            "services": [
                "Core Switch Deployment", "Firewall Configuration", "SD-WAN Implementation",
                "Network Monitoring Setup", "Wireless Access Points", "Load Balancer Setup",
                "VPN Gateway Installation", "Network Security Audit",
            ],
        },
        "Cloud Services": {
            "vendors": ["AWS", "Microsoft Azure", "Google Cloud", "IBM Cloud"],
            "services": [
                "Cloud Migration Assessment", "AWS EC2 Hosting", "Azure AD Integration",
                "Cloud Backup Solution", "Kubernetes Cluster Setup", "Serverless Architecture",
                "Cloud Cost Optimization", "Multi-Cloud Strategy",
            ],
        },
        "Cybersecurity": {
            "vendors": ["CrowdStrike", "Palo Alto Networks", "Fortinet", "Splunk"],
            "services": [
                "Endpoint Detection & Response", "SIEM Implementation", "Penetration Testing",
                "Zero Trust Architecture", "SOC Setup", "DLP Implementation",
                "Threat Intelligence Platform", "Incident Response Planning",
            ],
        },
        "Software Licensing": {
            "vendors": ["Microsoft", "Oracle", "SAP", "Salesforce"],
            "services": [
                "Microsoft 365 Enterprise", "Oracle DB Licensing", "SAP S/4HANA License",
                "Salesforce CRM License", "Power BI Pro Licenses", "ServiceNow ITSM",
                "Jira & Confluence Suite", "Adobe Creative Cloud",
            ],
        },
        "Managed Services": {
            "vendors": ["Accenture", "Infosys", "TCS", "Wipro"],
            "services": [
                "24/7 IT Helpdesk", "Infrastructure Monitoring", "Patch Management",
                "Backup & Disaster Recovery", "Database Administration",
                "Application Support", "IT Asset Management",
            ],
        },
    }

    rows = []
    for cat, info in categories.items():
        for vendor in info["vendors"]:
            # Each vendor quotes on 2-5 services from this category
            n_svc = random.randint(2, min(5, len(info["services"])))
            chosen_services = random.sample(info["services"], n_svc)
            services_text = "\n".join(chosen_services)

            base_price = random.randint(8, 95) * 1000
            variation = random.uniform(0.7, 1.4)
            price = round(base_price * variation, 2)

            file_name = "{} - {} Quote.pdf".format(vendor, cat.split()[0])

            rows.append({
                "Category": cat,
                "Vendor": vendor,
                "File Name": file_name,
                "Comments": services_text,
                "Quoted Price": price,
                "File Link": "",
            })

    df = pd.DataFrame(rows)
    return df


def generate_demo_quotation_xlsx():
    """Generates a demo quotation file for upload & score testing."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"

    # Header
    ws["A1"] = "QUOTATION"
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=16)
    ws["A3"] = "Vendor:"
    ws["B3"] = "NovaTech Solutions"
    ws["A4"] = "Date:"
    ws["B4"] = "2025-01-15"
    ws["A5"] = "Quote Ref:"
    ws["B5"] = "QT-2025-0042"
    ws["A6"] = "Valid Until:"
    ws["B6"] = "2025-03-15"

    # Services
    ws["A8"] = "Service Description"
    ws["B8"] = "Unit"
    ws["C8"] = "Qty"
    ws["D8"] = "Unit Price"
    ws["E8"] = "Amount"
    for c in ["A8", "B8", "C8", "D8", "E8"]:
        ws[c].font = openpyxl.styles.Font(bold=True)

    items = [
        ("Core Switch Deployment", "Project", 1, 45000),
        ("Firewall Configuration", "Project", 1, 18500),
        ("Network Monitoring Setup", "Annual", 1, 12000),
        ("SD-WAN Implementation", "Project", 1, 32000),
    ]

    row = 9
    subtotal = 0
    for desc, unit, qty, up in items:
        ws.cell(row=row, column=1, value=desc)
        ws.cell(row=row, column=2, value=unit)
        ws.cell(row=row, column=3, value=qty)
        ws.cell(row=row, column=4, value=up)
        ws.cell(row=row, column=5, value=qty * up)
        subtotal += qty * up
        row += 1

    row += 1
    ws.cell(row=row, column=4, value="Subtotal")
    ws.cell(row=row, column=5, value=subtotal)
    ws.cell(row=row, column=4).font = openpyxl.styles.Font(bold=True)

    row += 1
    ws.cell(row=row, column=4, value="Tax (8%)")
    ws.cell(row=row, column=5, value=round(subtotal * 0.08, 2))

    row += 1
    total = round(subtotal * 1.08, 2)
    ws.cell(row=row, column=4, value="Grand Total")
    ws.cell(row=row, column=5, value=total)
    ws.cell(row=row, column=4).font = openpyxl.styles.Font(bold=True, size=12)
    ws.cell(row=row, column=5).font = openpyxl.styles.Font(bold=True, size=12)

    # Adjust column widths
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════
# PROCESS CATALOG
# ════════════════════════════════════════════════════════════
def ai_detect_columns(df_raw):
    col_map = {}
    for c in df_raw.columns:
        cl = str(c).lower().strip()
        if any(k in cl for k in ["category", "type", "domain"]):
            if "Category" not in col_map:
                col_map["Category"] = c
        elif any(k in cl for k in ["vendor", "supplier", "company", "provider", "partner"]):
            if "Vendor" not in col_map:
                col_map["Vendor"] = c
        elif any(k in cl for k in ["file name", "filename", "document", "file", "attachment"]):
            if "File Name" not in col_map:
                col_map["File Name"] = c
        elif any(k in cl for k in ["link", "url", "hyperlink", "path"]):
            if "File Link" not in col_map:
                col_map["File Link"] = c
        elif any(k in cl for k in ["comment", "service", "description", "scope", "remark", "note"]):
            if "Comments" not in col_map:
                col_map["Comments"] = c
        elif any(k in cl for k in ["price", "cost", "amount", "value", "quote", "rate"]):
            if "Quoted Price" not in col_map:
                col_map["Quoted Price"] = c
    return col_map


def process_catalog_df(df_raw):
    """Process a raw dataframe into standard format."""
    df = df_raw.copy()
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    col_map = ai_detect_columns(df)
    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)

    for req in ["Category", "Vendor", "File Name"]:
        if req not in df.columns:
            df[req] = ""
    if "Comments" not in df.columns:
        df["Comments"] = ""
    if "Quoted Price" not in df.columns:
        df["Quoted Price"] = ""

    keep = ["Category", "Vendor", "File Name", "Comments", "Quoted Price"]
    for e in ["File Link", "Hyperlink"]:
        if e in df.columns:
            keep.append(e)
    df = df[[c for c in keep if c in df.columns]].copy()

    df = df[~(df["Category"].astype(str).str.strip().isin(["", "nan"]) &
              df["Vendor"].astype(str).str.strip().isin(["", "nan"]))].copy()

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    if "Hyperlink" not in df.columns:
        df["Hyperlink"] = ""

    df.reset_index(drop=True, inplace=True)

    def parse_svc(v):
        if not v or str(v).strip() in ["", "nan"]:
            return ["(unspecified)"]
        parts = [s.strip() for s in str(v).split("\n") if s.strip()]
        return parts or ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_svc)

    df_exp = df.explode("Services List").copy()
    df_exp.rename(columns={"Services List": "Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].str.strip()
    df_exp = df_exp[~df_exp["Service"].isin(["", "(unspecified)", "nan"])].reset_index(drop=True)

    return df, df_exp


def process_uploaded_catalog(file_bytes, filename):
    try:
        ext = filename.rsplit(".", 1)[-1].lower()
        if ext in ("xlsx", "xls"):
            raw = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", header=None)
        elif ext == "csv":
            raw = pd.read_csv(io.BytesIO(file_bytes), header=None)
        else:
            return None, None, "Unsupported file type."

        # Detect header
        header_row = 0
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            joined = " ".join(vals)
            if any(k in joined for k in ["vendor", "supplier", "company"]):
                header_row = i
                break

        if ext in ("xlsx", "xls"):
            df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", header=header_row)
        else:
            df = pd.read_csv(io.BytesIO(file_bytes), header=header_row)

        df_master, df_exp = process_catalog_df(df)
        return df_master, df_exp, None

    except Exception as e:
        return None, None, "Error: {}".format(e)


# ════════════════════════════════════════════════════════════
# LOAD DATA — generate dummy if no file
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    FILE_PATH = "Master Catalog.xlsx"

    if os.path.exists(FILE_PATH):
        try:
            raw = pd.read_excel(FILE_PATH, engine="openpyxl", header=None)
            header_row = 0
            for i, row in raw.iterrows():
                vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if any("category" in v for v in vals) and any("file" in v for v in vals):
                    header_row = i
                    break
            df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)
            return process_catalog_df(df)
        except Exception:
            pass

    # Generate dummy catalog
    df_raw = generate_dummy_catalog()
    return process_catalog_df(df_raw)


# ════════════════════════════════════════════════════════════
# DETERMINE DATA SOURCE
# ════════════════════════════════════════════════════════════
if ("uploaded_catalog_df" in st.session_state and
        st.session_state["uploaded_catalog_df"] is not None):
    df_master = st.session_state["uploaded_catalog_df"]
    df_exploded = st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE = "uploaded"
else:
    df_master, df_exploded = load_data()
    DATA_SOURCE = "default"

NO_DATA = df_master is None or df_exploded is None

if not NO_DATA:
    vendor_color_map = {v: get_color(i) for i, v in enumerate(sorted(df_master["Vendor"].unique()))}
else:
    vendor_color_map = {}


# ════════════════════════════════════════════════════════════
# UI HELPERS
# ════════════════════════════════════════════════════════════
def sb_label(txt):
    st.markdown(
        "<p style='color:#F0F0F0;font-weight:700;font-size:0.82em;"
        "margin:10px 0 4px;letter-spacing:0.5px;text-transform:uppercase'>{}</p>".format(txt),
        unsafe_allow_html=True)


def section_title(txt, caption=""):
    st.markdown("<div class='section-hdr'>{}</div>".format(txt), unsafe_allow_html=True)
    if caption:
        st.caption(caption)


def vendor_pill(v, color):
    return "<span class='vendor-badge' style='background:{}'>{}</span>".format(color, v)


def ai_box(content):
    st.markdown(
        "<div class='ai-box'><div class='ai-box-title'>🤖 AI Insight</div>{}</div>".format(content),
        unsafe_allow_html=True)


def kpi(col, val, lbl, bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div></div>".format(bg, val, lbl),
        unsafe_allow_html=True)


def resolve_url(row):
    for col in ["Hyperlink", "File Link", "File URL"]:
        url = str(row.get(col, "")).strip()
        if url and url != "nan" and url.startswith("http"):
            return url
    return ""


# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
selected_svcs = []
selected_cat = "All"
selected_vendor = "All"
d_filt = pd.DataFrame()

if not NO_DATA:
    with st.sidebar:
        st.markdown(
            "<div style='text-align:center;padding:18px 0 12px'>"
            "<div style='font-size:1.8em'>📋</div>"
            "<div style='font-size:1.0em;font-weight:700;color:white;margin:4px 0 2px'>IT Procurement</div>"
            "<div style='font-size:0.68em;color:#aaa;letter-spacing:1px;text-transform:uppercase'>"
            "Intelligence Dashboard</div></div>"
            "<hr style='border-color:#D04A02;border-width:2px;margin:0 0 12px'>",
            unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown(
                "<div style='background:#D04A02;color:white;padding:5px 10px;border-radius:2px;"
                "font-size:0.73em;font-weight:700;text-align:center;margin-bottom:8px'>"
                "📤 UPLOADED CATALOG</div>", unsafe_allow_html=True)

        sb_label("📂 Category")
        all_cats = ["All"] + sorted([c for c in df_master["Category"].unique() if str(c).strip() not in ["", "nan"]])
        selected_cat = st.selectbox("Category", all_cats, label_visibility="collapsed")

        sb_label("🏢 Vendor")
        vpool = df_master if selected_cat == "All" else df_master[df_master["Category"] == selected_cat]
        all_vendors = ["All"] + sorted([v for v in vpool["Vendor"].unique() if str(v).strip() not in ["", "nan"]])
        selected_vendor = st.selectbox("Vendor", all_vendors, label_visibility="collapsed")

        st.markdown("<hr style='border-color:#555;margin:10px 0'>", unsafe_allow_html=True)

        d_filt = df_exploded.copy()
        if selected_cat != "All":
            d_filt = d_filt[d_filt["Category"] == selected_cat]
        if selected_vendor != "All":
            d_filt = d_filt[d_filt["Vendor"] == selected_vendor]

        sb_label("🛠 Select Services")
        all_svcs = sorted([s for s in d_filt["Service"].unique() if str(s).strip() not in ["", "nan"]])

        svc_search = st.text_input("Search", placeholder="e.g. Firewall, Cloud…", label_visibility="collapsed")
        if svc_search:
            all_svcs = [s for s in all_svcs if svc_search.lower() in s.lower()]

        selected_svcs = st.multiselect(
            "Services", options=all_svcs, default=[],
            label_visibility="collapsed", key="main_svc_select")

        st.markdown("<hr style='border-color:#555;margin:10px 0'>", unsafe_allow_html=True)
        st.markdown(
            "<p style='color:#888;font-size:0.76em;margin:2px 0'>"
            "📄 {} quotes · 🛠 {} services · 🏢 {} vendors</p>".format(
                df_master["File Name"].nunique(), df_exploded["Service"].nunique(), df_master["Vendor"].nunique()),
            unsafe_allow_html=True)

        if selected_svcs:
            st.markdown(
                "<p style='color:#D04A02;font-size:0.78em;font-weight:700;margin:4px 0'>"
                "✅ {} selected</p>".format(len(selected_svcs)), unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown("<hr style='border-color:#555;margin:10px 0'>", unsafe_allow_html=True)
            if st.button("🔄 Reset Catalog", use_container_width=True):
                st.session_state["uploaded_catalog_df"] = None
                st.session_state["uploaded_catalog_exp"] = None
                st.rerun()


# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div class='dark-header'>"
    "<div class='tag'>IT Procurement Analytics</div>"
    "<h1>Procurement Intelligence Dashboard</h1>"
    "<p class='subtitle'>Browse quotations · Compare prices · Upload &amp; score · AI insights</p>"
    "</div>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════
if not NO_DATA:
    k1, k2, k3, k4 = st.columns(4)
    src = d_filt if not d_filt.empty else df_exploded
    kpi(k1, src["File Name"].nunique(), "Total Quotes", "#D04A02")
    kpi(k2, src["Service"].nunique(), "Unique Services", "#295477")
    kpi(k3, src["Vendor"].nunique(), "Vendors", "#299D8F")
    kpi(k4, src["Category"].nunique(), "Categories", "#2D2D2D")
    st.markdown("")


# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
if NO_DATA:
    tab5 = st.tabs(["🗂 Upload Catalog"])[0]
    tab1 = tab2 = tab3 = tab4 = None
else:
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Analytics", "📋 Browse & Compare", "📤 Upload & Score", "📄 Data Table", "🗂 Upload Catalog"])


# ════════════════════════════════════════════════════════════
# TAB 1 — ANALYTICS
# ════════════════════════════════════════════════════════════
if tab1 is not None:
    with tab1:
        use_df = d_filt if not d_filt.empty else df_exploded

        col_l, col_r = st.columns(2, gap="large")

        with col_l:
            section_title("SERVICE OVERLAP ANALYSIS", "Orange = quoted by multiple vendors")
            shared = (use_df.groupby("Service")["Vendor"].nunique()
                      .sort_values(ascending=False).head(20).reset_index())
            shared.columns = ["Service", "Vendor Count"]
            shared["Color"] = shared["Vendor Count"].apply(lambda x: "#D04A02" if x > 1 else "#C0C0C0")
            fig1 = go.Figure(go.Bar(
                x=shared["Vendor Count"], y=shared["Service"].str[:42],
                orientation="h", marker_color=shared["Color"], marker_line_width=0,
                text=shared["Vendor Count"], textposition="outside", textfont=dict(size=10)))
            fig1.update_layout(
                height=480, plot_bgcolor=CBG, paper_bgcolor=CBG,
                margin=dict(l=5, r=40, t=20, b=10), font=CFONT,
                xaxis=dict(title="Vendors", showgrid=True, gridcolor="#E0E0E0", zeroline=False),
                yaxis=dict(autorange="reversed", tickfont=dict(size=9.5)), bargap=0.35)
            st.plotly_chart(fig1, use_container_width=True)

        with col_r:
            section_title("VENDOR SERVICE COVERAGE", "Higher = broader capability")
            spv = (use_df.groupby("Vendor")["Service"].nunique()
                   .sort_values(ascending=False).reset_index())
            spv.columns = ["Vendor", "Count"]
            spv["Color"] = [vendor_color_map.get(v, "#8C8C8C") for v in spv["Vendor"]]
            fig2 = go.Figure(go.Bar(
                x=spv["Vendor"], y=spv["Count"], marker_color=spv["Color"],
                marker_line_width=0, text=spv["Count"], textposition="outside", textfont=dict(size=10)))
            fig2.update_layout(
                height=480, plot_bgcolor=CBG, paper_bgcolor=CBG,
                margin=dict(l=5, r=10, t=20, b=10), font=CFONT,
                yaxis=dict(title="Services", showgrid=True, gridcolor="#E0E0E0", zeroline=False),
                xaxis=dict(tickangle=-35, tickfont=dict(size=9.5)), bargap=0.35)
            st.plotly_chart(fig2, use_container_width=True)

        section_title("CATEGORY DISTRIBUTION")
        cat_c = (use_df.drop_duplicates(subset=["Category", "File Name"])
                 .groupby("Category").size().reset_index())
        cat_c.columns = ["Category", "Count"]
        if not cat_c.empty:
            fig3 = px.pie(cat_c, names="Category", values="Count", hole=0.50, color_discrete_sequence=COLORS)
            fig3.update_traces(textposition="outside", textinfo="label+percent", textfont_size=11, pull=[0.03]*len(cat_c))
            fig3.update_layout(height=380, margin=dict(l=20, r=20, t=20, b=20),
                               paper_bgcolor=CBG, font=CFONT,
                               legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=10)))
            st.plotly_chart(fig3, use_container_width=True)

        st.markdown("")
        section_title("AI VENDOR SUMMARY")
        svc_by_vendor = {}
        for v in df_master["Vendor"].unique():
            svc_by_vendor[v] = list(df_exploded[df_exploded["Vendor"] == v]["Service"].unique())
        ai_box(ai_service_summary(svc_by_vendor))


# ════════════════════════════════════════════════════════════
# TAB 2 — BROWSE & COMPARE
# ════════════════════════════════════════════════════════════
if tab2 is not None:
    with tab2:
        if not selected_svcs:
            st.info("👈 Select services from the sidebar to browse matching quotations.")
        else:
            use_filt = d_filt if not d_filt.empty else df_exploded
            d_sel = use_filt[use_filt["Service"].isin(selected_svcs)].copy()

            if d_sel.empty:
                st.warning("No results found.")
            else:
                # Vendor coverage summary
                vsmap = defaultdict(set)
                for _, r in d_sel.iterrows():
                    vsmap[r["Vendor"]].add(r["Service"])

                vendors_all = sorted([v for v, s in vsmap.items() if set(selected_svcs).issubset(s)])
                vendors_some = sorted([v for v, s in vsmap.items() if not set(selected_svcs).issubset(s)])

                if len(selected_svcs) > 1:
                    if vendors_all:
                        st.success("✅ {} vendor(s) cover ALL {} services: {}".format(
                            len(vendors_all), len(selected_svcs),
                            " · ".join(["**{}**".format(v) for v in vendors_all])))
                    else:
                        st.warning("No single vendor covers all {} services.".format(len(selected_svcs)))

                section_title("QUOTATION COMPARISON", "Price Score: 100=cheapest, 0=most expensive")

                for svc in selected_svcs:
                    d_svc = (d_sel[d_sel["Service"] == svc]
                             .drop_duplicates(subset=["Vendor", "File Name"])
                             .sort_values("Vendor"))
                    vc = d_svc["Vendor"].nunique()
                    tag = "SHARED" if vc > 1 else "SINGLE VENDOR"

                    with st.expander("🔹 {}  —  {} vendor(s) · {} file(s) · {}".format(svc, vc, len(d_svc), tag),
                                     expanded=True):

                        # Collect all prices for this service
                        all_prices_svc = []
                        for _, r in d_svc.iterrows():
                            qp = _parse_num(str(r.get("Quoted Price", "")).strip())
                            if qp > 0:
                                all_prices_svc.append(qp)

                        # Build table
                        html_rows = []
                        html_rows.append(
                            "<table class='comp-table'><thead><tr>"
                            "<th style='width:14%'>Vendor</th>"
                            "<th style='width:12%'>Category</th>"
                            "<th style='width:22%'>File Name</th>"
                            "<th style='width:13%'>Quoted Price</th>"
                            "<th style='width:11%'>Price Score</th>"
                            "<th style='width:14%'>Verdict</th>"
                            "<th style='width:14%'>Similarity</th>"
                            "</tr></thead><tbody>")

                        for i, (_, row) in enumerate(d_svc.iterrows()):
                            bg = "#ffffff" if i % 2 == 0 else "#F9F9F9"
                            color = vendor_color_map.get(row["Vendor"], "#8C8C8C")
                            fname = str(row.get("File Name", "")).strip()

                            v_cell = vendor_pill(row["Vendor"], color)
                            fn_cell = "<span style='font-family:monospace;font-size:0.80em'>{}</span>".format(fname[:35])

                            qp = _parse_num(str(row.get("Quoted Price", "")).strip())
                            qp_cell = "<span style='color:#22992E;font-weight:700;font-family:monospace'>{}</span>".format(
                                _fmt(qp)) if qp > 0 else "<span style='color:#bbb'>—</span>"

                            # Price score
                            others = [p for p in all_prices_svc if p != qp and p > 0]
                            if qp > 0 and others:
                                ps, _ = price_score(qp, all_prices_svc)
                                sc = score_color(ps)
                                ps_cell = "<span style='font-weight:800;color:{}'>{}/100</span>".format(sc, ps if ps is not None else "—")
                                v_html = verdict_html(ps)
                            elif qp > 0 and len(all_prices_svc) == 1:
                                ps_cell = "<span style='color:#295477;font-weight:700'>Baseline</span>"
                                v_html = "<span class='verdict-avg'>Only quote</span>"
                            else:
                                ps_cell = "<span style='color:#bbb'>—</span>"
                                v_html = "<span style='color:#bbb'>—</span>"

                            # Similarity (based on service match with other vendors)
                            vendor_svcs = set(d_sel[d_sel["Vendor"] == row["Vendor"]]["Service"].unique())
                            all_svc_set = set(selected_svcs)
                            sim = round(len(vendor_svcs & all_svc_set) / max(len(all_svc_set), 1) * 100, 0)
                            sim_color = "#22992E" if sim >= 70 else ("#FFB600" if sim >= 40 else "#E0301E")
                            sim_cell = "<span style='font-weight:700;color:{}'>{:.0f}%</span>".format(sim_color, sim)

                            html_rows.append(
                                "<tr style='background:{}'><td>{}</td><td style='color:#555'>{}</td>"
                                "<td>{}</td><td>{}</td><td style='text-align:center'>{}</td>"
                                "<td>{}</td><td style='text-align:center'>{}</td></tr>".format(
                                    bg, v_cell, row["Category"], fn_cell, qp_cell, ps_cell, v_html, sim_cell))

                        html_rows.append("</tbody></table>")
                        st.markdown("".join(html_rows), unsafe_allow_html=True)

                        # Price chart if multiple prices
                        if len(all_prices_svc) >= 2:
                            st.markdown("")
                            chart_data = []
                            for _, r in d_svc.iterrows():
                                qp = _parse_num(str(r.get("Quoted Price", "")).strip())
                                if qp > 0:
                                    chart_data.append({
                                        "Label": r["Vendor"],
                                        "Price": qp,
                                        "Color": vendor_color_map.get(r["Vendor"], "#8C8C8C"),
                                    })
                            if chart_data:
                                pdf_c = pd.DataFrame(chart_data)
                                avg_p = sum(r["Price"] for r in chart_data) / len(chart_data)
                                mf = go.Figure(go.Bar(
                                    x=pdf_c["Label"], y=pdf_c["Price"],
                                    marker_color=pdf_c["Color"], marker_line_width=0,
                                    text=pdf_c["Price"].apply(_fmt), textposition="outside"))
                                mf.add_hline(y=avg_p, line_dash="dash", line_color="#FFB600", line_width=2,
                                             annotation_text="Avg: {}".format(_fmt(avg_p)),
                                             annotation_position="top right")
                                mf.update_layout(
                                    height=280, plot_bgcolor=CBG, paper_bgcolor=CBG,
                                    margin=dict(l=5, r=10, t=20, b=10), font=CFONT,
                                    yaxis=dict(title="Price", showgrid=True, gridcolor="#E0E0E0", zeroline=False),
                                    xaxis=dict(tickangle=-20), bargap=0.4)
                                st.plotly_chart(mf, use_container_width=True)

                        # AI insight for this service
                        if all_prices_svc:
                            vp_map = {}
                            for _, r in d_svc.iterrows():
                                qp = _parse_num(str(r.get("Quoted Price", "")).strip())
                                if qp > 0:
                                    vp_map[r["Vendor"]] = qp
                            ai_box(ai_price_insight(
                                min(all_prices_svc), all_prices_svc, vp_map))


# ════════════════════════════════════════════════════════════
# TAB 3 — UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
if tab3 is not None:
    with tab3:
        st.markdown(
            "<div class='dark-header'>"
            "<div class='tag'>New Quotation Analysis</div>"
            "<h1>Upload &amp; Score Quotation</h1>"
            "<p class='subtitle'>Upload a new quotation → auto-extract price → score vs master catalog</p>"
            "</div>", unsafe_allow_html=True)

        # Demo file download
        section_title("DEMO FILE")
        st.markdown(
            "<p style='font-size:0.85em;color:#555;margin-bottom:8px'>"
            "Download this sample quotation to test the scoring pipeline:</p>", unsafe_allow_html=True)

        demo_bytes = generate_demo_quotation_xlsx()
        st.download_button(
            label="📥 Download Demo Quotation (NovaTech_Quote.xlsx)",
            data=demo_bytes,
            file_name="NovaTech_Network_Quote.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="demo_dl")

        st.markdown("<hr style='border-color:#e0e0e0;margin:16px 0'>", unsafe_allow_html=True)

        section_title("STEP 1 — UPLOAD QUOTATION FILE")
        uploaded = st.file_uploader("Upload", type=["pdf", "xlsx", "xls", "docx"],
                                    label_visibility="collapsed", key="quot_upload")

        if uploaded is not None:
            content = uploaded.read()
            ext = uploaded.name.rsplit(".", 1)[-1].lower()
            fname_up = uploaded.name
            st.success("✅ Uploaded: **{}** ({} KB)".format(fname_up, round(len(content)/1024, 1)))

            section_title("STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting price from document…"):
                result = extract_price_from_bytes(content, ext)
                new_price = result["price_num"]
                new_text = result["text"]

            if new_price > 0:
                st.markdown(
                    "<div class='score-card score-low'>"
                    "<div class='score-label' style='color:#22992E'>Extracted Price</div>"
                    "<div class='score-num' style='color:#22992E'>{}</div>"
                    "</div>".format(_fmt(new_price)), unsafe_allow_html=True)
            else:
                st.warning("Could not auto-extract price. Enter manually below.")
                manual = st.number_input("Enter price (USD)", min_value=0.0, step=100.0, value=0.0, key="man_price")
                if manual > 0:
                    new_price = manual

            # Show extracted text preview
            if new_text:
                with st.expander("📝 Extracted Document Text (preview)", expanded=False):
                    st.text(new_text[:2000])

            section_title("STEP 3 — SELECT MATCHING SERVICES")
            st.markdown("<p style='font-size:0.83em;color:#555'>Select the services this quotation covers:</p>",
                        unsafe_allow_html=True)

            all_svcs_up = sorted([s for s in df_exploded["Service"].unique() if str(s).strip() not in ["", "nan"]])
            svc_search_up = st.text_input("Search services", placeholder="e.g. Firewall, Cloud…",
                                          key="svc_up_search", label_visibility="collapsed")
            if svc_search_up:
                all_svcs_up = [s for s in all_svcs_up if svc_search_up.lower() in s.lower()]

            new_services = st.multiselect("Select services", options=all_svcs_up, default=[],
                                          label_visibility="collapsed", key="upload_svc_select")

            cat_filter_up = st.selectbox("Filter by category", options=["All"] + sorted([
                c for c in df_master["Category"].unique() if str(c).strip() not in ["", "nan"]]),
                key="cat_up_filter")

            section_title("STEP 4 — COMPARISON RESULTS")

            if not new_services and new_price <= 0:
                st.info("Select services and/or provide a price to see comparison.")
            else:
                # Find matching historical quotes
                if new_services:
                    candidates = df_exploded[df_exploded["Service"].isin(new_services)].copy()
                else:
                    candidates = df_exploded.copy()

                if cat_filter_up != "All":
                    candidates = candidates[candidates["Category"] == cat_filter_up]

                cand_files = (candidates.drop_duplicates(subset=["File Name", "Vendor"])
                              [["File Name", "Vendor", "Category", "Quoted Price"]].copy())

                if cand_files.empty:
                    st.warning("No matching historical quotes found.")
                else:
                    # Collect historical prices
                    hist_prices = []
                    vendor_p_map = {}
                    for _, r in cand_files.iterrows():
                        qp = _parse_num(str(r.get("Quoted Price", "")).strip())
                        if qp > 0:
                            hist_prices.append(qp)
                            vendor_p_map[r["Vendor"]] = qp

                    if new_price > 0 and hist_prices:
                        ps, ps_label = price_score(new_price, hist_prices)
                        avg_h = sum(hist_prices) / len(hist_prices)
                        mn_h, mx_h = min(hist_prices), max(hist_prices)

                        # Similarity score
                        if new_services:
                            hist_svcs_set = set(candidates["Service"].unique())
                            new_svcs_set = set(new_services)
                            sim_score = round(len(new_svcs_set & hist_svcs_set) / max(len(new_svcs_set | hist_svcs_set), 1) * 100, 1)
                        else:
                            sim_score = 0

                        # Score cards row
                        sc1, sc2, sc3, sc4 = st.columns(4)

                        sc1.markdown(
                            "<div class='score-card {}'>"
                            "<div class='score-label' style='color:{}'>Price Score</div>"
                            "<div class='score-num' style='color:{}'>{}/100</div>"
                            "<div style='font-size:0.73em;color:#555;margin-top:3px'>vs {} quotes</div>"
                            "</div>".format(score_css(ps), score_color(ps), score_color(ps),
                                            ps if ps is not None else "N/A", len(hist_prices)),
                            unsafe_allow_html=True)

                        sc2.markdown(
                            "<div class='score-card score-medium'>"
                            "<div class='score-label' style='color:#D04A02'>Your Price</div>"
                            "<div class='score-num' style='color:#D04A02'>{}</div>"
                            "</div>".format(_fmt(new_price)), unsafe_allow_html=True)

                        sc3.markdown(
                            "<div class='score-card score-medium'>"
                            "<div class='score-label' style='color:#295477'>Historical Avg</div>"
                            "<div class='score-num' style='color:#295477'>{}</div>"
                            "<div style='font-size:0.72em;color:#555;margin-top:3px'>min {} · max {}</div>"
                            "</div>".format(_fmt(avg_h), _fmt(mn_h), _fmt(mx_h)), unsafe_allow_html=True)

                        sim_c = "#22992E" if sim_score >= 70 else ("#FFB600" if sim_score >= 40 else "#E0301E")
                        sc4.markdown(
                            "<div class='score-card' style='border-color:{}'>"
                            "<div class='score-label' style='color:{}'>Service Match</div>"
                            "<div class='score-num' style='color:{}'>{:.0f}%</div>"
                            "</div>".format(sim_c, sim_c, sim_c, sim_score), unsafe_allow_html=True)

                        # Verdict
                        st.markdown("")
                        if ps is not None and ps >= 70:
                            st.markdown(
                                "<div style='background:#E8F5E9;border-left:5px solid #22992E;"
                                "padding:14px 18px;border-radius:4px;margin:8px 0'>"
                                "<span style='font-size:1.1em;font-weight:700;color:#22992E'>"
                                "✅ VERDICT: COMPETITIVE</span><br>"
                                "<span style='font-size:0.87em;color:#2D2D2D'>{}</span></div>".format(ps_label),
                                unsafe_allow_html=True)
                        elif ps is not None and ps >= 40:
                            st.markdown(
                                "<div style='background:#FFF8E1;border-left:5px solid #FFB600;"
                                "padding:14px 18px;border-radius:4px;margin:8px 0'>"
                                "<span style='font-size:1.1em;font-weight:700;color:#E6A000'>"
                                "⚠ VERDICT: AVERAGE PRICING</span><br>"
                                "<span style='font-size:0.87em;color:#2D2D2D'>{}</span></div>".format(ps_label),
                                unsafe_allow_html=True)
                        else:
                            st.markdown(
                                "<div style='background:#FFEBEE;border-left:5px solid #E0301E;"
                                "padding:14px 18px;border-radius:4px;margin:8px 0'>"
                                "<span style='font-size:1.1em;font-weight:700;color:#E0301E'>"
                                "❌ VERDICT: OVERPRICED — NEGOTIATE</span><br>"
                                "<span style='font-size:0.87em;color:#2D2D2D'>{}</span></div>".format(ps_label),
                                unsafe_allow_html=True)

                        st.markdown("")
                        ai_box(ai_price_insight(new_price, hist_prices, vendor_p_map))

                        # Comparison chart
                        st.markdown("")
                        section_title("PRICE POSITIONING CHART")
                        chart_data = []
                        for _, r in cand_files.iterrows():
                            qp = _parse_num(str(r.get("Quoted Price", "")).strip())
                            if qp > 0:
                                chart_data.append({
                                    "Label": r["Vendor"],
                                    "Price": qp,
                                    "Type": "Historical",
                                    "Color": vendor_color_map.get(r["Vendor"], "#8C8C8C"),
                                })
                        chart_data.append({
                            "Label": "YOUR QUOTE",
                            "Price": new_price,
                            "Type": "New Upload",
                            "Color": "#D04A02",
                        })
                        cdf = pd.DataFrame(chart_data).sort_values("Price")
                        cf = go.Figure()
                        hist_df = cdf[cdf["Type"] == "Historical"]
                        new_df = cdf[cdf["Type"] == "New Upload"]
                        cf.add_trace(go.Bar(
                            x=hist_df["Label"], y=hist_df["Price"],
                            marker_color=hist_df["Color"], marker_line_width=0,
                            name="Historical", text=hist_df["Price"].apply(_fmt), textposition="outside"))
                        cf.add_trace(go.Bar(
                            x=new_df["Label"], y=new_df["Price"],
                            marker_color="#D04A02", marker_line_width=0,
                            name="Your Quote", text=new_df["Price"].apply(_fmt), textposition="outside"))
                        cf.add_hline(y=avg_h, line_dash="dash", line_color="#FFB600", line_width=2,
                                     annotation_text="Avg: {}".format(_fmt(avg_h)),
                                     annotation_position="top right")
                        cf.update_layout(
                            height=360, plot_bgcolor=CBG, paper_bgcolor=CBG,
                            margin=dict(l=5, r=10, t=20, b=10), font=CFONT, barmode="group",
                            yaxis=dict(title="Price (USD)", showgrid=True, gridcolor="#E0E0E0", zeroline=False),
                            xaxis=dict(tickangle=-25),
                            legend=dict(orientation="h", x=0, y=1.05), bargap=0.25)
                        st.plotly_chart(cf, use_container_width=True)

                        # Detailed comparison table
                        st.markdown("")
                        section_title("DETAILED COMPARISON TABLE")
                        html_t = ["<table class='comp-table'><thead><tr>"
                                  "<th style='width:16%'>Vendor</th>"
                                  "<th style='width:14%'>Category</th>"
                                  "<th style='width:22%'>File</th>"
                                  "<th style='width:13%'>Hist. Price</th>"
                                  "<th style='width:12%'>Δ vs Yours</th>"
                                  "<th style='width:12%'>Price Score</th>"
                                  "<th style='width:11%'>Verdict</th>"
                                  "</tr></thead><tbody>"]

                        for i, (_, r) in enumerate(cand_files.iterrows()):
                            bg = "#ffffff" if i % 2 == 0 else "#F9F9F9"
                            qp = _parse_num(str(r.get("Quoted Price", "")).strip())
                            color = vendor_color_map.get(r["Vendor"], "#8C8C8C")

                            if qp > 0 and new_price > 0:
                                delta_pct = round((new_price - qp) / qp * 100, 1)
                                if delta_pct > 0:
                                    delta_cell = "<span style='color:#E0301E;font-weight:700'>+{}%</span>".format(delta_pct)
                                elif delta_pct < 0:
                                    delta_cell = "<span style='color:#22992E;font-weight:700'>{}%</span>".format(delta_pct)
                                else:
                                    delta_cell = "<span style='color:#555'>0%</span>"

                                ps_r, _ = price_score(qp, hist_prices)
                                ps_cell = "<span style='font-weight:700;color:{}'>{}/100</span>".format(
                                    score_color(ps_r), ps_r if ps_r is not None else "—")
                                v_cell = verdict_html(ps_r)
                            else:
                                delta_cell = "<span style='color:#bbb'>—</span>"
                                ps_cell = "<span style='color:#bbb'>—</span>"
                                v_cell = "<span style='color:#bbb'>—</span>"

                            html_t.append(
                                "<tr style='background:{}'><td>{}</td><td style='color:#555'>{}</td>"
                                "<td><span style='font-size:0.80em'>{}</span></td>"
                                "<td style='font-family:monospace;color:#295477;font-weight:700'>{}</td>"
                                "<td style='text-align:center'>{}</td>"
                                "<td style='text-align:center'>{}</td>"
                                "<td>{}</td></tr>".format(
                                    bg, vendor_pill(r["Vendor"], color), r["Category"],
                                    str(r["File Name"])[:30], _fmt(qp),
                                    delta_cell, ps_cell, v_cell))

                        html_t.append("</tbody></table>")
                        st.markdown("".join(html_t), unsafe_allow_html=True)

                    elif new_price > 0:
                        st.warning("No historical prices available for comparison. Showing file matches only.")
                        st.dataframe(cand_files.head(20), use_container_width=True)
                    else:
                        st.info("Enter a price or select services to see the scoring comparison.")


# ════════════════════════════════════════════════════════════
# TAB 4 — DATA TABLE
# ════════════════════════════════════════════════════════════
if tab4 is not None:
    with tab4:
        dm = df_master.copy()
        if selected_cat != "All":
            dm = dm[dm["Category"] == selected_cat]
        if selected_vendor != "All":
            dm = dm[dm["Vendor"] == selected_vendor]
        st.dataframe(dm.drop(columns=["Services List", "Hyperlink"], errors="ignore"),
                     use_container_width=True, height=500)


# ════════════════════════════════════════════════════════════
# TAB 5 — UPLOAD CATALOG
# ════════════════════════════════════════════════════════════
with tab5:
    st.markdown(
        "<div class='dark-header'>"
        "<div class='tag'>Catalog Management</div>"
        "<h1>Upload Master Catalog</h1>"
        "<p class='subtitle'>Upload any Excel/CSV catalog — AI auto-detects columns and builds the dashboard.</p>"
        "</div>", unsafe_allow_html=True)

    if DATA_SOURCE == "uploaded":
        st.success("✅ Using uploaded catalog: **{}** rows, **{}** vendors, **{}** services".format(
            len(df_master), df_master["Vendor"].nunique(), df_exploded["Service"].nunique()))

    with st.expander("📖 How to prepare your catalog file", expanded=False):
        st.markdown("""
**Required columns** (auto-detected):

| Column | Accepted Names |
|---|---|
| **Category** | Category, Type, Domain |
| **Vendor** | Vendor, Supplier, Company, Provider |
| **File Name** | File Name, Filename, Document |
| **Services** | Comments, Services, Description, Scope |
| **Price** | Quoted Price, Price, Cost, Amount |
| **Link** | File Link, URL, Hyperlink |

**Formats:** `.xlsx`, `.xls`, `.csv`
        """)

    section_title("UPLOAD CATALOG FILE")
    catalog_file = st.file_uploader("Upload Catalog", type=["xlsx", "xls", "csv"],
                                    label_visibility="collapsed", key="catalog_upload")

    if catalog_file is not None:
        file_bytes = catalog_file.read()
        fname_cat = catalog_file.name

        st.markdown(
            "<div class='catalog-step'><div class='catalog-step-num'>Step 1 — File Received</div>"
            "<b>{}</b> — {} KB</div>".format(fname_cat, round(len(file_bytes)/1024, 1)),
            unsafe_allow_html=True)

        with st.spinner("AI analyzing catalog…"):
            df_new, df_exp_new, err = process_uploaded_catalog(file_bytes, fname_cat)

        if err:
            st.error("❌ {}".format(err))
        elif df_new is None:
            st.error("❌ Could not process file.")
        else:
            st.markdown(
                "<div class='catalog-step'><div class='catalog-step-num'>Step 2 — Detected</div>"
                "<b>{}</b> rows · <b>{}</b> vendors · <b>{}</b> categories</div>".format(
                    len(df_new), df_new["Vendor"].nunique(), df_new["Category"].nunique()),
                unsafe_allow_html=True)

            # AI Analysis
            insights = ai_analyze_catalog(df_new, df_exp_new)

            ins_cols = st.columns(2)
            items = [
                ("Overview", insights.get("overview", ""), "#295477"),
                ("Top Vendor", insights.get("top_vendor", ""), "#299D8F"),
                ("Most Competitive", insights.get("competitive", ""), "#D04A02"),
                ("Category Focus", insights.get("category", ""), "#FFB600"),
            ]
            if insights.get("pricing"):
                items.append(("Pricing", insights["pricing"], "#22992E"))

            for idx, (title, content, color) in enumerate(items):
                if content:
                    ins_cols[idx % 2].markdown(
                        "<div class='insight-card'>"
                        "<div style='font-size:0.70em;font-weight:700;letter-spacing:1px;"
                        "text-transform:uppercase;color:{};margin-bottom:5px'>{}</div>"
                        "<div style='font-size:0.86em;color:#2D2D2D'>{}</div></div>".format(color, title, content),
                        unsafe_allow_html=True)

            recs = insights.get("recommendations", [])
            if recs:
                section_title("AI RECOMMENDATIONS")
                for rec in recs:
                    st.markdown(
                        "<div style='background:#FFF3F0;border-left:3px solid #D04A02;"
                        "padding:7px 12px;border-radius:2px;margin-bottom:5px;font-size:0.85em'>▸ {}</div>".format(rec),
                        unsafe_allow_html=True)

            # Preview charts
            st.markdown("")
            section_title("CATALOG PREVIEW")
            pc1, pc2 = st.columns(2)
            with pc1:
                spv_n = df_exp_new.groupby("Vendor")["Service"].nunique().sort_values(ascending=False).reset_index()
                spv_n.columns = ["Vendor", "Services"]
                vc_n = {v: get_color(i) for i, v in enumerate(spv_n["Vendor"])}
                pf1 = go.Figure(go.Bar(
                    x=spv_n["Vendor"], y=spv_n["Services"],
                    marker_color=[vc_n.get(v, "#8C8C8C") for v in spv_n["Vendor"]],
                    marker_line_width=0, text=spv_n["Services"], textposition="outside"))
                pf1.update_layout(title="Services per Vendor", height=300, plot_bgcolor=CBG,
                                  paper_bgcolor=CBG, margin=dict(l=5, r=10, t=40, b=10), font=CFONT,
                                  yaxis=dict(showgrid=True, gridcolor="#E0E0E0", zeroline=False),
                                  xaxis=dict(tickangle=-30), bargap=0.35)
                st.plotly_chart(pf1, use_container_width=True)

            with pc2:
                cat_n = df_new.drop_duplicates(subset=["Category", "File Name"]).groupby("Category").size().reset_index()
                cat_n.columns = ["Category", "Count"]
                if not cat_n.empty:
                    pf2 = px.pie(cat_n, names="Category", values="Count", hole=0.45, color_discrete_sequence=COLORS)
                    pf2.update_traces(textposition="outside", textinfo="label+percent", textfont_size=10)
                    pf2.update_layout(title="Categories", height=300, margin=dict(l=10, r=10, t=40, b=10),
                                      paper_bgcolor=CBG, font=CFONT)
                    st.plotly_chart(pf2, use_container_width=True)

            st.markdown("")
            section_title("DATA PREVIEW")
            st.dataframe(df_new.drop(columns=["Services List", "Hyperlink"], errors="ignore").head(20),
                         use_container_width=True, height=300)

            st.markdown("")
            ac, _ = st.columns([2, 3])
            if ac.button("✅ Apply This Catalog", type="primary", use_container_width=True, key="apply_cat"):
                st.session_state["uploaded_catalog_df"] = df_new
                st.session_state["uploaded_catalog_exp"] = df_exp_new
                st.success("Catalog applied! Dashboard now uses your uploaded data.")
                st.rerun()
