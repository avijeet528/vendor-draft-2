# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  PwC Brand | Source Sans Pro | Streamlit
#  Demo-ready with verdict system
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
    font-family:'Source Sans Pro','Helvetica Neue',
                Arial,sans-serif !important;
}
.main .block-container {
    background-color:#F3F3F3 !important;
    padding-top:1.5rem;
    max-width:100% !important;
    padding-left:2rem !important;
    padding-right:2rem !important;
}
#MainMenu {visibility:hidden;}
footer    {visibility:hidden;}
header    {visibility:hidden;}
[data-testid="collapsedControl"]{display:none !important;}

section[data-testid="stSidebar"]{
    background-color:#2D2D2D !important;
    border-right:3px solid #D04A02;
    min-width:300px !important;
    max-width:300px !important;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div{
    color:#F0F0F0 !important;
    font-family:'Source Sans Pro',sans-serif !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"]{
    background-color:#FFFFFF !important;
    border-radius:2px !important;
    border:1px solid #999 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] *{
    color:#2D2D2D !important;
}
section[data-testid="stSidebar"] div[data-baseweb="input"]{
    background-color:#FFFFFF !important;
    border-radius:2px !important;
}
section[data-testid="stSidebar"] div[data-baseweb="input"] input{
    color:#2D2D2D !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"]{
    background-color:#D04A02 !important;
    border-radius:2px !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span{
    color:white !important;
}
.kpi-box{
    border-radius:4px;padding:18px 10px;
    text-align:center;color:white;
    border-left:5px solid rgba(255,255,255,0.25);
}
.kpi-value{
    font-size:2.1em;font-weight:700;
    margin:0;line-height:1.1;letter-spacing:-0.5px;
}
.kpi-label{
    font-size:0.78em;font-weight:700;opacity:0.9;
    margin-top:5px;letter-spacing:0.8px;
    text-transform:uppercase;
}
button[data-baseweb="tab"]{
    font-weight:600 !important;
    font-size:0.92em !important;
    color:#7D7D7D !important;
}
button[data-baseweb="tab"][aria-selected="true"]{
    color:#D04A02 !important;
    border-bottom:3px solid #D04A02 !important;
}
div[data-testid="stExpander"] details>summary{
    list-style:none !important;
    padding-left:12px !important;
}
div[data-testid="stExpander"] details>summary::before,
div[data-testid="stExpander"] details>summary::after,
div[data-testid="stExpander"] details>summary::-webkit-details-marker,
div[data-testid="stExpander"] details>summary::marker{
    display:none !important;content:"" !important;width:0 !important;
}
div[data-testid="stExpander"] details summary p{
    font-weight:700;font-size:0.94em;
    color:#2D2D2D !important;
    padding-left:0 !important;margin-left:0 !important;
}
div[data-testid="stExpander"] details{
    border:1px solid #ddd;border-radius:4px;
    margin-bottom:10px;padding:2px 0;
}
.comp-table{
    width:100%;border-collapse:collapse;
    table-layout:fixed;font-size:0.83em;
    border:1px solid #e0e0e0;
}
.comp-table thead tr{background:#2D2D2D;}
.comp-table thead th{
    padding:10px 10px;text-align:left;
    font-weight:700;font-size:0.80em;
    letter-spacing:0.4px;text-transform:uppercase;
    color:white !important;border:none;word-break:break-word;
}
.comp-table tbody tr:nth-child(even){background:#F3F3F3;}
.comp-table tbody tr:hover{background:#FCE8DC;}
.comp-table tbody td{
    padding:8px 10px;border-bottom:1px solid #e8e8e8;
    vertical-align:middle;word-break:break-word;
    font-size:0.82em;color:#2D2D2D;
}
.comp-table th:nth-child(1),.comp-table td:nth-child(1){width:13%;}
.comp-table th:nth-child(2),.comp-table td:nth-child(2){width:12%;}
.comp-table th:nth-child(3),.comp-table td:nth-child(3){width:22%;}
.comp-table th:nth-child(4),.comp-table td:nth-child(4){width:10%;}
.comp-table th:nth-child(5),.comp-table td:nth-child(5){width:10%;}
.comp-table th:nth-child(6),.comp-table td:nth-child(6){width:10%;}
.comp-table th:nth-child(7),.comp-table td:nth-child(7){width:13%;}
.comp-table th:nth-child(8),.comp-table td:nth-child(8){width:10%;}
.vendor-badge{
    display:inline-block;padding:3px 8px;
    border-radius:2px;color:white;
    font-size:0.78em;font-weight:700;
    white-space:nowrap;overflow:hidden;
    text-overflow:ellipsis;max-width:100%;
    box-sizing:border-box;
}
.score-card{
    border-radius:4px;padding:14px 16px;
    margin-bottom:10px;border-left:5px solid #D04A02;
}
.score-card.green {background:#F0FFF4;border-color:#22992E;}
.score-card.yellow{background:#FFF8E1;border-color:#FFB600;}
.score-card.red   {background:#FFF3F0;border-color:#E0301E;}
.ai-box{
    background:#F8F0FF;border-left:5px solid #6E2585;
    border-radius:4px;padding:14px 18px;margin:10px 0;
}
.verdict-green {
    background:#F0FFF4;border:2px solid #22992E;
    border-radius:4px;padding:12px 16px;
    color:#22992E;font-weight:700;
}
.verdict-yellow{
    background:#FFF8E1;border:2px solid #FFB600;
    border-radius:4px;padding:12px 16px;
    color:#856404;font-weight:700;
}
.verdict-red   {
    background:#FFF3F0;border:2px solid #E0301E;
    border-radius:4px;padding:12px 16px;
    color:#E0301E;font-weight:700;
}
.demo-banner{
    background:linear-gradient(
        135deg,#D04A02 0%,#B83D00 100%);
    color:white;padding:10px 18px;
    border-radius:4px;margin-bottom:16px;
    font-size:0.88em;font-weight:600;
}
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
def get_color(i): return COLORS[i % len(COLORS)]
CFONT = dict(
    family="Source Sans Pro,Helvetica Neue,Arial",
    size=11,color="#2D2D2D")
CBG = "#F3F3F3"


# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"""
    (?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}(?:[,]\d{3})*(?:\.\d{1,2})?
    |(?:[\$\€\£]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?
    |\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?
    """,re.VERBOSE|re.IGNORECASE)
TOTAL_KW = [
    "grand total","total amount","total price","amount due",
    "net total","total cost","total value","quote total",
    "subtotal","estimated total","total",
]

def _parse_num(s):
    try: return float(re.sub(r"[^\d.]","",str(s)) or "0")
    except: return 0.0

def _fmt(val):
    try:
        v = float(re.sub(r"[^\d.]","",str(val)) or "0")
        if v <= 0: return "—"
        return "${:,.2f}".format(v)
    except: return str(val)

def _best_price(text):
    tl = text.lower()
    for kw in TOTAL_KW:
        idx = tl.find(kw)
        if idx == -1: continue
        snip  = text[max(0,idx-20):idx+300]
        hits  = PRICE_RE.findall(snip)
        valid = [h.strip() for h in hits if _parse_num(h)>=50]
        if valid: return max(valid,key=_parse_num)
    all_h = PRICE_RE.findall(text)
    valid = [h.strip() for h in all_h if _parse_num(h)>=100]
    if valid: return max(valid,key=_parse_num)
    return ""

def _text_from_bytes(content,ext):
    text=""; ext=ext.lower().strip(".")
    try:
        if ext=="pdf":
            if not PDF_OK: return ""
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for p in pdf.pages:
                    t=p.extract_text()
                    if t: text+=t+"\n"
        elif ext in ("xlsx","xls"):
            wb=openpyxl.load_workbook(
                io.BytesIO(content),data_only=True,read_only=True)
            rows_text=[]
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rs="  ".join(
                        str(c) for c in row if c is not None)
                    if rs.strip(): rows_text.append(rs)
            text="\n".join(rows_text); wb.close()
        elif ext=="docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml=z.read("word/document.xml").decode(
                        "utf-8",errors="ignore")
                    text=re.sub(r"<[^>]+>"," ",xml)
                    text=re.sub(r"\s{2,}","\n",text)
    except: pass
    return text

def extract_price_from_bytes(content,ext):
    text=_text_from_bytes(content,ext)
    price=_best_price(text)
    return {"price":price,
            "price_num":_parse_num(price) if price else 0.0,
            "text":text[:5000]}

def extract_price_from_file(filepath):
    try:
        with open(filepath,"rb") as f:
            content=f.read()
        ext=filepath.rsplit(".",1)[-1].lower()
        return extract_price_from_bytes(content,ext)
    except:
        return {"price":"","price_num":0.0,"text":""}


# ════════════════════════════════════════════════════════════
# SCORING & VERDICT
# ════════════════════════════════════════════════════════════
def price_score(new_price,hist_prices):
    valid=[p for p in hist_prices if p>0]
    if not valid or new_price<=0:
        return None,"No comparison data",0,0,0
    mn=min(valid); mx=max(valid)
    avg=sum(valid)/len(valid)
    if mx==mn: return 50,"Same as historical average",avg,mn,mx
    score=round((1-(new_price-mn)/(mx-mn))*100,1)
    score=max(0,min(100,score))
    pct=round((new_price-avg)/avg*100,1)
    if new_price<avg:
        label="{}% BELOW average — COMPETITIVE".format(abs(pct))
    elif new_price>avg:
        label="{}% ABOVE average — REVIEW NEEDED".format(abs(pct))
    else:
        label="Matches historical average"
    return score,label,avg,mn,mx

def score_color(s):
    if s is None: return "#8C8C8C"
    if s>=70: return "#22992E"
    if s>=40: return "#FFB600"
    return "#E0301E"

def score_css(s):
    if s is None: return "yellow"
    if s>=70: return "green"
    if s>=40: return "yellow"
    return "red"

def get_verdict(ps):
    if ps is None:
        return "⚪ No Data","No comparison data available.","#8C8C8C"
    if ps>=70:
        return (
            "✅ COMPETITIVE",
            "This quote is priced competitively vs "
            "historical data. Proceed with confidence.",
            "#22992E")
    if ps>=40:
        return (
            "🟡 AVERAGE",
            "This quote is within average range. "
            "Negotiate for a small discount if possible.",
            "#856404")
    return (
        "🔴 HIGH — NEGOTIATE",
        "This quote is above the historical average. "
        "Strongly recommend negotiating or seeking "
        "alternative vendors.",
        "#E0301E")


# ════════════════════════════════════════════════════════════
# AI INSIGHTS
# ════════════════════════════════════════════════════════════
def ai_service_summary(df_master,df_exploded):
    svc_by_v={}
    for v in df_master["Vendor"].unique():
        svc_by_v[v]=list(
            df_exploded[df_exploded["Vendor"]==v
                       ]["Service"].unique())
    if not svc_by_v: return "No vendor data."
    best  =max(svc_by_v,key=lambda v:len(svc_by_v[v]))
    n_best=len(svc_by_v[best])
    total =len(set(s for svcs in svc_by_v.values()
                   for s in svcs))
    shared=[s for s in set(
                s for svcs in svc_by_v.values()
                for s in svcs)
            if sum(1 for svcs in svc_by_v.values()
                   if s in svcs)>1]
    lines=["**{}** covers the most services "
           "({} of {} total).".format(best,n_best,total)]
    if shared:
        lines.append(
            "**{}** service(s) offered by multiple vendors "
            "— ideal for competitive benchmarking.".format(
                len(shared)))
    return " ".join(lines)

def ai_price_insight(new_price,hist_prices,vendor_prices):
    valid=[p for p in hist_prices if p>0]
    if not valid or new_price<=0:
        return "Insufficient data for price analysis."
    avg=sum(valid)/len(valid)
    mn=min(valid); mx=max(valid)
    pct=round((new_price-avg)/avg*100,1)
    lines=[]
    if new_price<=mn:
        lines.append("This quote is the **lowest price** "
                     "seen — excellent value.")
    elif new_price>=mx:
        lines.append("This quote is **above all historical "
                     "prices** — negotiate strongly.")
    elif pct>15:
        lines.append("Quote is **{}% above** average. "
                     "Request a revised quote.".format(
                         abs(pct)))
    elif pct<-15:
        lines.append("Quote is **{}% below** average — "
                     "very competitive.".format(abs(pct)))
    else:
        lines.append("Quote is **within normal range** "
                     "({}% vs average).".format(pct))
    if vendor_prices:
        best_v=min(vendor_prices,key=vendor_prices.get)
        lines.append(
            "**{}** has historically offered the lowest "
            "prices.".format(best_v))
    return " ".join(lines)

def generate_selection_verdict(
        selected_svcs, d_sel, vendor_prices_map,
        df_exploded):
    """
    Generates a comprehensive verdict when user
    selects services showing price comparison
    and recommendation.
    """
    if d_sel.empty or not selected_svcs:
        return None

    vsmap=defaultdict(set)
    for _,r in d_sel.iterrows():
        vsmap[r["Vendor"]].add(r["Service"])

    # Vendors covering ALL selected services
    full_cover=[v for v,s in vsmap.items()
                if set(selected_svcs).issubset(s)]

    # All prices for selected services
    all_prices=[]
    vendor_min={}
    for v,p in vendor_prices_map.items():
        if p>0:
            all_prices.append(p)
            vendor_min[v]=p

    if not all_prices:
        return {
            "title"       :"📊 Vendor Coverage Summary",
            "vendors_all" :full_cover,
            "vendors_some":sorted(
                [v for v in vsmap if v not in full_cover]),
            "has_prices"  :False,
            "lines"       :[],
        }

    avg_p  =sum(all_prices)/len(all_prices)
    min_p  =min(all_prices)
    max_p  =max(all_prices)
    best_v =min(vendor_min,key=vendor_min.get)
    worst_v=max(vendor_min,key=vendor_min.get)
    spread =round((max_p-min_p)/min_p*100,1) if min_p>0 else 0

    lines=[]
    lines.append(
        "**{}** vendor(s) quoted for the selected "
        "service(s).".format(len(vendor_min)))
    lines.append(
        "Price range: **{}** — **{}** "
        "(spread: **{}%**).".format(
            _fmt(min_p),_fmt(max_p),spread))
    lines.append(
        "Average quoted price: **{}**.".format(
            _fmt(avg_p)))

    if full_cover:
        lines.append(
            "**{}** offer(s) ALL selected services — "
            "best for consolidated procurement.".format(
                ", ".join(full_cover)))
    else:
        lines.append(
            "No single vendor covers all selected "
            "services — consider multi-vendor approach.")

    lines.append(
        "**Best price:** {} at {} — "
        "{}% below average.".format(
            best_v, _fmt(min_p),
            round((avg_p-min_p)/avg_p*100,1)))

    if spread > 20:
        lines.append(
            "⚠️ Large price spread ({}%) detected — "
            "significant negotiation opportunity "
            "exists.".format(spread))

    return {
        "title"        :"📊 Procurement Verdict",
        "vendors_all"  :full_cover,
        "vendors_some" :sorted(
            [v for v in vsmap if v not in full_cover]),
        "has_prices"   :True,
        "lines"        :lines,
        "best_vendor"  :best_v,
        "best_price"   :min_p,
        "worst_vendor" :worst_v,
        "worst_price"  :max_p,
        "avg_price"    :avg_p,
        "spread"       :spread,
        "vendor_prices":vendor_min,
    }


# ════════════════════════════════════════════════════════════
# HYPERLINK EXTRACTION
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path):
    link_map={}
    try:
        wb=openpyxl.load_workbook(file_path)
        ws=wb.active
        fn_col=hr=None
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and
                        str(cell.value).strip().lower()
                        =="file name"):
                    fn_col=cell.column
                    hr=cell.row; break
            if fn_col: break
        if fn_col:
            for row in ws.iter_rows(
                    min_row=hr+1,
                    min_col=fn_col,max_col=fn_col):
                cell=row[0]
                if cell.value and cell.hyperlink:
                    link_map[str(cell.value).strip()]=\
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
    FILE_PATH="Master Catalog.xlsx"
    if not os.path.exists(FILE_PATH):
        return None,None

    raw=pd.read_excel(
        FILE_PATH,engine="openpyxl",header=None)
    header_row=None
    for i,row in raw.iterrows():
        vals=[str(v).strip().lower()
              for v in row.values if pd.notna(v)]
        if (any("category" in v for v in vals) and
                any("file" in v for v in vals)):
            header_row=i; break
    if header_row is None: return None,None

    df=pd.read_excel(
        FILE_PATH,engine="openpyxl",header=header_row)
    df=df.loc[:,df.columns.notna()]
    df.columns=[str(c).strip() for c in df.columns]
    df.dropna(how="all",inplace=True)

    col_map={}
    for c in df.columns:
        cl=str(c).lower().strip()
        if cl=="category": col_map["Category"]=c
        elif "vendor" in cl or "type" in cl:
            col_map["Vendor"]=c
        elif cl=="file name": col_map["File Name"]=c
        elif cl=="file link": col_map["File Link"]=c
        elif cl=="file url":  col_map["File URL"]=c
        elif "comment" in cl: col_map["Comments"]=c
        elif "quoted" in cl or "price" in cl:
            col_map["Quoted Price"]=c

    df.rename(
        columns={v:k for k,v in col_map.items()},
        inplace=True)

    keep=["Category","Vendor","File Name","Comments"]
    for e in ["File Link","File URL","Quoted Price"]:
        if e in df.columns: keep.append(e)
    df=df[[c for c in keep if c in df.columns]].copy()

    df=df[~(
        df["Category"].astype(str).str.strip().isin(
            ["","nan"]) &
        df["Vendor"].astype(str).str.strip().isin(
            ["","nan"]))].copy()

    for col in df.columns:
        df[col]=df[col].fillna("").astype(str).str.strip()
    df.reset_index(drop=True,inplace=True)

    hmap=extract_hyperlinks(FILE_PATH)
    df["Hyperlink"]=df["File Name"].map(hmap).fillna("")
    for fb in ["File Link","File URL"]:
        if fb in df.columns:
            df["Hyperlink"]=df.apply(
                lambda r:r["Hyperlink"]
                if r["Hyperlink"] not in ["","nan"]
                else r[fb],axis=1)

    def parse_svc(v):
        if not v or str(v).strip() in ["","nan"]:
            return ["(unspecified)"]
        parts=[s.strip()
               for s in str(v).split("\n") if s.strip()]
        return parts or ["(unspecified)"]

    df["Services List"]=df["Comments"].apply(parse_svc)

    df_exp=df.explode("Services List").copy()
    df_exp.rename(
        columns={"Services List":"Service"},inplace=True)
    df_exp["Service"]=df_exp["Service"].str.strip()
    df_exp=df_exp[
        ~df_exp["Service"].isin(
            ["","(unspecified)","nan"])
    ].reset_index(drop=True)

    return df,df_exp


# ════════════════════════════════════════════════════════════
# PROCESS UPLOADED CATALOG
# ════════════════════════════════════════════════════════════
def process_uploaded_catalog(file_bytes,filename):
    try:
        ext=filename.rsplit(".",1)[-1].lower()
        if ext in ("xlsx","xls"):
            raw=pd.read_excel(
                io.BytesIO(file_bytes),
                engine="openpyxl",header=None)
        elif ext=="csv":
            raw=pd.read_csv(
                io.BytesIO(file_bytes),header=None)
        else:
            return None,None,"Unsupported file type."

        header_row=0
        for i,row in raw.iterrows():
            vals=[str(v).strip().lower()
                  for v in row.values if pd.notna(v)]
            joined=" ".join(vals)
            has_v=any(k in joined
                for k in ["vendor","supplier","company"])
            has_f=any(k in joined
                for k in ["file","document","attachment"])
            has_c=any(k in joined
                for k in ["category","type","domain"])
            if (has_v or has_c) and has_f:
                header_row=i; break

        if ext in ("xlsx","xls"):
            df=pd.read_excel(
                io.BytesIO(file_bytes),
                engine="openpyxl",header=header_row)
        else:
            df=pd.read_csv(
                io.BytesIO(file_bytes),header=header_row)

        df=df.loc[:,df.columns.notna()]
        df.columns=[str(c).strip() for c in df.columns]
        df.dropna(how="all",inplace=True)

        col_map={}
        for c in df.columns:
            cl=str(c).lower().strip()
            if any(k in cl for k in
                   ["category","type","domain"]):
                if "Category" not in col_map:
                    col_map["Category"]=c
            elif any(k in cl for k in
                     ["vendor","supplier","company"]):
                if "Vendor" not in col_map:
                    col_map["Vendor"]=c
            elif any(k in cl for k in
                     ["file name","filename","document"]):
                if "File Name" not in col_map:
                    col_map["File Name"]=c
            elif any(k in cl for k in
                     ["link","url","hyperlink"]):
                if "File Link" not in col_map:
                    col_map["File Link"]=c
            elif any(k in cl for k in
                     ["comment","service","description",
                      "scope"]):
                if "Comments" not in col_map:
                    col_map["Comments"]=c
            elif any(k in cl for k in
                     ["price","cost","amount","value"]):
                if "Quoted Price" not in col_map:
                    col_map["Quoted Price"]=c

        df.rename(
            columns={v:k for k,v in col_map.items()},
            inplace=True)

        for req in ["Category","Vendor","File Name"]:
            if req not in df.columns: df[req]=""
        if "Comments" not in df.columns:
            df["Comments"]=""

        keep=["Category","Vendor","File Name","Comments"]
        for e in ["File Link","Quoted Price"]:
            if e in df.columns: keep.append(e)
        df=df[[c for c in keep if c in df.columns]].copy()

        df=df[~(
            df["Category"].astype(str).str.strip().isin(
                ["","nan"]) &
            df["Vendor"].astype(str).str.strip().isin(
                ["","nan"]))].copy()

        for col in df.columns:
            df[col]=df[col].fillna("").astype(
                str).str.strip()
        df.reset_index(drop=True,inplace=True)
        df["Hyperlink"]=""

        def parse_svc(v):
            if not v or str(v).strip() in ["","nan"]:
                return ["(unspecified)"]
            parts=[s.strip()
                   for s in str(v).split("\n")
                   if s.strip()]
            return parts or ["(unspecified)"]

        df["Services List"]=df["Comments"].apply(parse_svc)

        df_exp=df.explode("Services List").copy()
        df_exp.rename(
            columns={"Services List":"Service"},
            inplace=True)
        df_exp["Service"]=df_exp["Service"].str.strip()
        df_exp=df_exp[
            ~df_exp["Service"].isin(
                ["","(unspecified)","nan"])
        ].reset_index(drop=True)

        return df,df_exp,None
    except Exception as e:
        return None,None,str(e)


# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
if "uploaded_catalog_df" in st.session_state and \
        st.session_state["uploaded_catalog_df"] is not None:
    df_master  =st.session_state["uploaded_catalog_df"]
    df_exploded=st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE="uploaded"
else:
    df_master,df_exploded=load_data()
    DATA_SOURCE="file"

NO_DATA=df_master is None or df_exploded is None

if not NO_DATA:
    vendor_color_map={
        v:get_color(i)
        for i,v in enumerate(
            sorted(df_master["Vendor"].unique()))}
else:
    vendor_color_map={}

DEMO_DIR="demo_quotes"
DEMO_QUOT="DEMO_New_Quotation_AlphaNetworks.xlsx"


# ════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════
def sb_label(txt):
    st.markdown(
        "<p style='color:#F0F0F0;font-weight:700;"
        "font-size:0.85em;margin:12px 0 4px;"
        "letter-spacing:0.5px;text-transform:uppercase'>"
        "{}</p>".format(txt),
        unsafe_allow_html=True)

def section_title(txt,caption=""):
    st.markdown(
        "<div style='font-size:0.78em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:4px'>{}</div>".format(
            txt),unsafe_allow_html=True)
    if caption: st.caption(caption)

def vendor_pill(v,color):
    return ("<span class='vendor-badge' "
            "style='background:{}'>{}</span>".format(
                color,v))

def ai_box(content):
    st.markdown(
        "<div class='ai-box'>"
        "<div style='font-size:0.68em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#6E2585;margin-bottom:6px'>"
        "AI Insight</div>{}</div>".format(content),
        unsafe_allow_html=True)

def resolve_url(row):
    url=str(row.get("Hyperlink","")).strip()
    if not url or url=="nan":
        url=str(row.get("File Link","")).strip()
    if not url or url=="nan":
        url=str(row.get("File URL","")).strip()
    return "" if url=="nan" else url

def kpi(col,val,lbl,bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div>"
        "</div>".format(bg,val,lbl),
        unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
selected_svcs  =[]
selected_cat   ="All"
selected_vendor="All"
d_filt         =pd.DataFrame()

if not NO_DATA:
    with st.sidebar:
        st.markdown(
            "<div style='text-align:center;"
            "padding:20px 0 14px'>"
            "<div style='font-size:2em'>📋</div>"
            "<div style='font-size:1.05em;"
            "font-weight:700;color:white;"
            "margin:5px 0 2px'>IT Procurement</div>"
            "<div style='font-size:0.70em;color:#aaa;"
            "letter-spacing:1px;text-transform:uppercase'>"
            "Intelligence Dashboard</div></div>"
            "<hr style='border-color:#D04A02;"
            "border-width:2px;margin:0 0 14px'>",
            unsafe_allow_html=True)

        if DATA_SOURCE=="uploaded":
            st.markdown(
                "<div style='background:#D04A02;"
                "color:white;padding:6px 10px;"
                "border-radius:2px;font-size:0.75em;"
                "font-weight:700;text-align:center;"
                "margin-bottom:10px'>"
                "📤 USING UPLOADED CATALOG</div>",
                unsafe_allow_html=True)

        # Category
        sb_label("📂 Category")
        all_cats=["All"]+sorted([
            c for c in df_master["Category"].unique()
            if str(c).strip() not in ["","nan"]])
        selected_cat=st.selectbox(
            "Category",all_cats,
            label_visibility="collapsed")

        # Vendor
        sb_label("🏢 Vendor")
        vpool=(df_master if selected_cat=="All"
               else df_master[
                   df_master["Category"]==selected_cat])
        all_vendors=["All"]+sorted([
            v for v in vpool["Vendor"].unique()
            if str(v).strip() not in ["","nan"]])
        selected_vendor=st.selectbox(
            "Vendor",all_vendors,
            label_visibility="collapsed")

        st.markdown(
            "<hr style='border-color:#555;margin:12px 0'>",
            unsafe_allow_html=True)

        # Filter data
        d_filt=df_exploded.copy()
        if selected_cat!="All":
            d_filt=d_filt[d_filt["Category"]==selected_cat]
        if selected_vendor!="All":
            d_filt=d_filt[d_filt["Vendor"]==selected_vendor]

        # Service search
        sb_label("🔍 Search Services")
        svc_search=st.text_input(
            "Search",placeholder="e.g. Cisco, Oracle…",
            label_visibility="collapsed")

        avail=sorted([
            s for s in d_filt["Service"].unique()
            if str(s).strip() not in ["","nan"]])
        if svc_search:
            avail=[s for s in avail
                   if svc_search.lower() in s.lower()]

        # Simple multiselect — no subcategories
        sb_label("🛠 Select Services ({} available)".format(
            len(avail)))
        selected_svcs=st.multiselect(
            "Services",
            options=avail,
            default=[],
            label_visibility="collapsed",
            help="Select services to compare vendors & prices")

        st.markdown(
            "<hr style='border-color:#555;margin:12px 0'>",
            unsafe_allow_html=True)

        st.markdown(
            "<p style='color:#888;font-size:0.78em;"
            "margin:2px 0'>"
            "📄 {} quotes | 🛠 {} services | "
            "🏢 {} vendors</p>".format(
                len(df_master),
                df_exploded["Service"].nunique(),
                df_master["Vendor"].nunique()),
            unsafe_allow_html=True)

        if selected_svcs:
            st.markdown(
                "<p style='color:#D04A02;"
                "font-size:0.80em;font-weight:700;"
                "margin:4px 0'>✅ {} service(s) "
                "selected</p>".format(
                    len(selected_svcs)),
                unsafe_allow_html=True)

        if DATA_SOURCE=="uploaded":
            st.markdown(
                "<hr style='border-color:#555;"
                "margin:12px 0'>",
                unsafe_allow_html=True)
            if st.button(
                "🔄 Reset to Default Catalog",
                use_container_width=True):
                st.session_state[
                    "uploaded_catalog_df"]=None
                st.session_state[
                    "uploaded_catalog_exp"]=None
                st.rerun()


# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div style='background:#2D2D2D;color:white;"
    "padding:20px 28px;border-radius:4px;"
    "border-left:6px solid #D04A02;"
    "margin-bottom:22px'>"
    "<div style='font-size:0.72em;font-weight:700;"
    "letter-spacing:2px;text-transform:uppercase;"
    "color:#D04A02;margin-bottom:5px'>"
    "IT Procurement Analytics</div>"
    "<h1 style='margin:0;font-size:1.45em;"
    "font-weight:700;color:white'>"
    "Procurement Intelligence Dashboard</h1>"
    "<p style='margin:6px 0 0;opacity:0.6;"
    "font-size:0.85em'>"
    "Browse quotations · Compare prices · "
    "Upload &amp; score · AI insights · Verdict system"
    "</p></div>",
    unsafe_allow_html=True)

# Demo banner
if os.path.exists(DEMO_DIR):
    st.markdown(
        "<div class='demo-banner'>"
        "🎯 <b>DEMO MODE</b> — "
        "Sample quotation data loaded. "
        "Upload <b>demo_quotes/{}</b> in the "
        "Upload &amp; Score tab to see the "
        "verdict system in action."
        "</div>".format(DEMO_QUOT),
        unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════
if not NO_DATA:
    use_filt=(d_filt if not d_filt.empty
              else df_exploded)
    k1,k2,k3,k4=st.columns(4)
    kpi(k1,
        use_filt["File Name"].nunique() if not use_filt.empty
        else df_master["File Name"].nunique(),
        "Total Quotes","#D04A02")
    kpi(k2,
        use_filt["Service"].nunique() if not use_filt.empty
        else df_exploded["Service"].nunique(),
        "Unique Services","#295477")
    kpi(k3,
        use_filt["Vendor"].nunique() if not use_filt.empty
        else df_master["Vendor"].nunique(),
        "Vendors","#299D8F")
    kpi(k4,
        use_filt["Category"].nunique() if not use_filt.empty
        else df_master["Category"].nunique(),
        "Categories","#2D2D2D")
    st.markdown("<br>",unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
tab1,tab2,tab3,tab4,tab5=st.tabs([
    "📊 Analytics",
    "📋 Browse & Verdict",
    "📤 Upload & Score",
    "📄 Data Table",
    "🗂 Upload Catalog",
])


# ════════════════════════════════════════════════════════════
# TAB 1 — ANALYTICS
# ════════════════════════════════════════════════════════════
with tab1:
    if NO_DATA:
        st.info("No catalog. Go to 🗂 Upload Catalog.")
    else:
        use_df=(d_filt if not d_filt.empty
                else df_exploded)
        col_l,col_r=st.columns(2,gap="large")

        with col_l:
            section_title(
                "SERVICE OVERLAP ANALYSIS",
                "Orange = quoted by multiple vendors.")
            shared=(use_df.groupby("Service")[
                "Vendor"].nunique()
                .sort_values(ascending=False)
                .head(20).reset_index())
            shared.columns=["Service","Vendor Count"]
            shared["Color"]=shared["Vendor Count"].apply(
                lambda x:"#D04A02" if x>1 else "#C0C0C0")
            fig1=go.Figure(go.Bar(
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
                plot_bgcolor=CBG,paper_bgcolor=CBG,
                margin=dict(l=5,r=40,t=20,b=10),
                font=CFONT,
                xaxis=dict(
                    title="Vendors",showgrid=True,
                    gridcolor="#E0E0E0",zeroline=False),
                yaxis=dict(
                    autorange="reversed",
                    tickfont=dict(size=9.5)),
                bargap=0.35)
            st.plotly_chart(fig1,use_container_width=True)

        with col_r:
            section_title(
                "VENDOR SERVICE COVERAGE",
                "Higher = broader capability.")
            spv=(use_df.groupby("Vendor")[
                "Service"].nunique()
                .sort_values(ascending=False)
                .reset_index())
            spv.columns=["Vendor","Count"]
            spv["Color"]=[
                vendor_color_map.get(v,"#8C8C8C")
                for v in spv["Vendor"]]
            fig2=go.Figure(go.Bar(
                x=spv["Vendor"],y=spv["Count"],
                marker_color=spv["Color"],
                marker_line_width=0,
                text=spv["Count"],
                textposition="outside",
                textfont=dict(size=10),
            ))
            fig2.update_layout(
                height=480,
                plot_bgcolor=CBG,paper_bgcolor=CBG,
                margin=dict(l=5,r=10,t=20,b=10),
                font=CFONT,
                yaxis=dict(
                    title="Unique Services",
                    showgrid=True,gridcolor="#E0E0E0",
                    zeroline=False),
                xaxis=dict(
                    tickangle=-35,
                    tickfont=dict(size=9.5)),
                bargap=0.35)
            st.plotly_chart(fig2,use_container_width=True)

        section_title(
            "PROCUREMENT CATEGORY DISTRIBUTION",
            "Share of quote files across categories.")
        cat_c=(use_df.drop_duplicates(
                   subset=["Category","File Name"])
               .groupby("Category").size()
               .reset_index())
        cat_c.columns=["Category","Count"]
        if not cat_c.empty:
            fig3=px.pie(
                cat_c,names="Category",values="Count",
                hole=0.50,
                color_discrete_sequence=COLORS)
            fig3.update_traces(
                textposition="outside",
                textinfo="label+percent",
                textfont_size=11,
                pull=[0.03]*len(cat_c))
            fig3.update_layout(
                height=380,
                margin=dict(l=20,r=20,t=20,b=20),
                paper_bgcolor=CBG,font=CFONT,
                legend=dict(
                    orientation="v",x=1.02,y=0.5,
                    font=dict(size=10)))
            st.plotly_chart(
                fig3,use_container_width=True)

        st.markdown("<br>",unsafe_allow_html=True)
        section_title("AI VENDOR SUMMARY")
        ai_box(ai_service_summary(df_master,df_exploded))


# ════════════════════════════════════════════════════════════
# TAB 2 — BROWSE & VERDICT
# ════════════════════════════════════════════════════════════
with tab2:
    if NO_DATA:
        st.info("No catalog. Go to 🗂 Upload Catalog.")
    elif not selected_svcs:
        st.info(
            "👈 Select services from the sidebar to "
            "browse quotations and see the verdict.")

        # Show demo tip
        if os.path.exists(DEMO_DIR):
            st.markdown(
                "<div style='background:white;"
                "border:1px solid #e0e0e0;"
                "border-radius:4px;"
                "padding:16px 20px;margin-top:12px'>"
                "<div style='font-size:0.72em;"
                "font-weight:700;letter-spacing:1px;"
                "text-transform:uppercase;"
                "color:#D04A02;margin-bottom:8px'>"
                "💡 Try These Demo Services</div>"
                "<p style='font-size:0.87em;"
                "color:#555;margin:0'>"
                "Select <b>Cisco Catalyst C9300</b> "
                "from the sidebar to see 3 vendors "
                "quoting for the same service with "
                "price comparison &amp; verdict."
                "</p></div>",
                unsafe_allow_html=True)
    else:
        use_filt2=(d_filt if not d_filt.empty
                   else df_exploded)
        d_sel=use_filt2[
            use_filt2["Service"].isin(
                selected_svcs)].copy()

        if d_sel.empty:
            st.warning("No results found.")
        else:
            has_price="Quoted Price" in d_sel.columns

            # Collect all prices per vendor
            vendor_prices_map={}
            for _,r in d_sel.drop_duplicates(
                    subset=["Vendor","File Name"]).iterrows():
                qp=_parse_num(
                    str(r.get("Quoted Price","")).strip())
                ck="px_{}".format(
                    str(r.get("File Name","")).strip())
                ca=st.session_state.get(ck)
                ep=ca["price_num"] if ca else 0.0
                ref=ep if ep>0 else qp
                v=r["Vendor"]
                if ref>0:
                    if v not in vendor_prices_map:
                        vendor_prices_map[v]=ref
                    else:
                        vendor_prices_map[v]=min(
                            vendor_prices_map[v],ref)

            # Generate verdict
            verdict=generate_selection_verdict(
                selected_svcs,d_sel,
                vendor_prices_map,df_exploded)

            # ── VERDICT BANNER ──────────────────────────
            if verdict:
                st.markdown(
                    "<div style='font-size:0.78em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#D04A02;margin-bottom:10px'>"
                    "PROCUREMENT VERDICT</div>",
                    unsafe_allow_html=True)

                if verdict["has_prices"]:
                    # Score cards
                    vc1,vc2,vc3,vc4=st.columns(4)

                    best_p=verdict.get("best_price",0)
                    worst_p=verdict.get("worst_price",0)
                    avg_p=verdict.get("avg_price",0)
                    spread=verdict.get("spread",0)
                    best_v=verdict.get("best_vendor","")

                    vc1.markdown(
                        "<div class='score-card green'>"
                        "<div style='font-size:0.68em;"
                        "font-weight:700;"
                        "letter-spacing:1px;"
                        "text-transform:uppercase;"
                        "color:#22992E'>Best Price</div>"
                        "<div style='font-size:1.6em;"
                        "font-weight:800;color:#22992E'>"
                        "{}</div>"
                        "<div style='font-size:0.75em;"
                        "color:#555;margin-top:2px'>"
                        "by {}</div>"
                        "</div>".format(
                            _fmt(best_p),best_v),
                        unsafe_allow_html=True)

                    vc2.markdown(
                        "<div class='score-card yellow'>"
                        "<div style='font-size:0.68em;"
                        "font-weight:700;"
                        "letter-spacing:1px;"
                        "text-transform:uppercase;"
                        "color:#856404'>Avg Price</div>"
                        "<div style='font-size:1.6em;"
                        "font-weight:800;"
                        "color:#856404'>{}</div>"
                        "<div style='font-size:0.75em;"
                        "color:#555;margin-top:2px'>"
                        "across {} vendors</div>"
                        "</div>".format(
                            _fmt(avg_p),
                            len(vendor_prices_map)),
                        unsafe_allow_html=True)

                    vc3.markdown(
                        "<div class='score-card red'>"
                        "<div style='font-size:0.68em;"
                        "font-weight:700;"
                        "letter-spacing:1px;"
                        "text-transform:uppercase;"
                        "color:#E0301E'>Highest Price</div>"
                        "<div style='font-size:1.6em;"
                        "font-weight:800;color:#E0301E'>"
                        "{}</div>"
                        "<div style='font-size:0.75em;"
                        "color:#555;margin-top:2px'>"
                        "by {}</div>"
                        "</div>".format(
                            _fmt(worst_p),
                            verdict.get("worst_vendor","")),
                        unsafe_allow_html=True)

                    vc4.markdown(
                        "<div class='score-card {}' >"
                        "<div style='font-size:0.68em;"
                        "font-weight:700;"
                        "letter-spacing:1px;"
                        "text-transform:uppercase;"
                        "color:{}'>Price Spread</div>"
                        "<div style='font-size:1.6em;"
                        "font-weight:800;color:{}'>"
                        "{}%</div>"
                        "<div style='font-size:0.75em;"
                        "color:#555;margin-top:2px'>"
                        "negotiation room</div>"
                        "</div>".format(
                            "red" if spread>20
                            else "yellow" if spread>10
                            else "green",
                            score_color(
                                100-min(spread*2,100)),
                            score_color(
                                100-min(spread*2,100)),
                            spread),
                        unsafe_allow_html=True)

                    st.markdown(
                        "<br>",unsafe_allow_html=True)

                    # Verdict box
                    for line in verdict["lines"]:
                        st.markdown(
                            "<div style='background:white;"
                            "border-left:4px solid #D04A02;"
                            "padding:8px 14px;"
                            "border-radius:2px;"
                            "margin-bottom:6px;"
                            "font-size:0.87em'>"
                            "▸ {}</div>".format(line),
                            unsafe_allow_html=True)

                    # Price comparison chart
                    st.markdown(
                        "<br>",unsafe_allow_html=True)
                    section_title(
                        "PRICE COMPARISON — "
                        "ALL VENDORS FOR SELECTED SERVICES")

                    chart_rows=[]
                    for v,p in sorted(
                            vendor_prices_map.items(),
                            key=lambda x:x[1]):
                        ps,_,_,_,_=price_score(
                            p,list(
                                vendor_prices_map.values()))
                        vc_css=score_css(ps)
                        chart_rows.append({
                            "Vendor":v,
                            "Price":p,
                            "Color":vendor_color_map.get(
                                v,"#8C8C8C"),
                        })

                    if chart_rows:
                        cdf=pd.DataFrame(chart_rows)
                        cf=go.Figure(go.Bar(
                            x=cdf["Vendor"],
                            y=cdf["Price"],
                            marker_color=cdf["Color"],
                            marker_line_width=0,
                            text=cdf["Price"].apply(
                                _fmt),
                            textposition="outside",
                        ))
                        cf.add_hline(
                            y=avg_p,
                            line_dash="dash",
                            line_color="#FFB600",
                            line_width=2,
                            annotation_text="Avg: {}".format(
                                _fmt(avg_p)),
                            annotation_position=
                                "top right")
                        cf.update_layout(
                            height=320,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(
                                l=5,r=10,t=20,b=10),
                            font=CFONT,
                            yaxis=dict(
                                title="Price (USD)",
                                showgrid=True,
                                gridcolor="#E0E0E0",
                                zeroline=False),
                            xaxis=dict(tickangle=-20),
                            bargap=0.4)
                        st.plotly_chart(
                            cf,use_container_width=True)

                    # Per-vendor score table
                    st.markdown(
                        "<br>",unsafe_allow_html=True)
                    section_title(
                        "VENDOR PRICE SCORE CARD")

                    all_p_vals=list(
                        vendor_prices_map.values())
                    rows=["<table class='comp-table'>"
                          "<thead><tr>"
                          "<th>Vendor</th>"
                          "<th>Price</th>"
                          "<th>vs Average</th>"
                          "<th>Price Score</th>"
                          "<th>Verdict</th>"
                          "</tr></thead><tbody>"]

                    for i,(v,p) in enumerate(
                            sorted(
                                vendor_prices_map.items(),
                                key=lambda x:x[1])):
                        bg="white" if i%2==0 else "#F3F3F3"
                        vc=vendor_color_map.get(
                            v,"#8C8C8C")
                        others=[px for px in all_p_vals
                                if px!=p]
                        ps=None
                        if p>0 and others:
                            ps,_,_,_,_=price_score(
                                p,others)
                        sc=score_color(ps)
                        pct_vs_avg=(
                            round((p-avg_p)/avg_p*100,1)
                            if avg_p>0 else 0)
                        vs_avg_txt=(
                            "{}% below avg".format(
                                abs(pct_vs_avg))
                            if pct_vs_avg<0
                            else "{}% above avg".format(
                                abs(pct_vs_avg))
                            if pct_vs_avg>0
                            else "At average")
                        vs_color=(
                            "#22992E" if pct_vs_avg<0
                            else "#E0301E"
                            if pct_vs_avg>5
                            else "#856404")
                        vt,_,vc_col=get_verdict(ps)
                        rows.append(
                            "<tr style='background:{}'>"
                            "<td>{}</td>"
                            "<td style='font-family:"
                            "monospace;font-weight:700;"
                            "color:#2D2D2D'>{}</td>"
                            "<td style='color:{}'>"
                            "{}</td>"
                            "<td style='text-align:center'>"
                            "<span style='font-weight:800;"
                            "font-size:1.1em;color:{}'>"
                            "{}/100</span></td>"
                            "<td><span style='color:{};"
                            "font-weight:700;font-size:"
                            "0.85em'>{}</span></td>"
                            "</tr>".format(
                                bg,
                                vendor_pill(v,vc),
                                _fmt(p),
                                vs_color,vs_avg_txt,
                                sc,
                                ps if ps is not None
                                else "—",
                                vc_col,vt))

                    rows.append("</tbody></table>")
                    st.markdown(
                        "".join(rows),
                        unsafe_allow_html=True)

                else:
                    st.info(
                        "No price data available. "
                        "Click 'Extract Prices' in the "
                        "quotation tables below to "
                        "enable scoring.")

                st.markdown(
                    "<hr style='border:none;"
                    "border-top:2px solid #D04A02;"
                    "margin:20px 0'>",
                    unsafe_allow_html=True)

            # ── Per-service tables ──────────────────────
            section_title(
                "QUOTATION FILES — PER SERVICE")

            vsmap=defaultdict(set)
            for _,r in d_sel.iterrows():
                vsmap[r["Vendor"]].add(r["Service"])
            vendors_all=sorted([
                v for v,s in vsmap.items()
                if set(selected_svcs).issubset(s)])

            if len(selected_svcs)>1:
                if vendors_all:
                    st.success(
                        "✅ {} vendor(s) cover ALL {} "
                        "services: {}".format(
                            len(vendors_all),
                            len(selected_svcs),
                            " · ".join(
                                ["**{}**".format(v)
                                 for v in vendors_all])))
                else:
                    st.warning(
                        "No single vendor covers all "
                        "{} selected services.".format(
                            len(selected_svcs)))

            for svc in selected_svcs:
                d_svc=(d_sel[d_sel["Service"]==svc]
                       .drop_duplicates(
                           subset=["Vendor","File Name"])
                       .sort_values("Vendor"))
                vc=d_svc["Vendor"].nunique()
                s_tag=("SHARED" if vc>1
                        else "SINGLE VENDOR")

                with st.expander(
                    "{}  —  {} vendor(s) · {} file(s) · {}"
                    .format(svc,vc,len(d_svc),s_tag),
                    expanded=True):

                    pills=" ".join([
                        vendor_pill(
                            v,vendor_color_map.get(
                                v,"#8C8C8C"))
                        for v in sorted(
                            d_svc["Vendor"].unique())])
                    st.markdown(
                        "<div style='margin-bottom:12px'>"
                        "<b style='font-size:0.87em'>"
                        "Vendors:</b>&nbsp;&nbsp;"
                        "{}</div>".format(pills),
                        unsafe_allow_html=True)

                    all_prices_svc=[]
                    for _,r in d_svc.iterrows():
                        qp=_parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp>0:
                            all_prices_svc.append(qp)
                        ck="px_{}".format(
                            str(r.get(
                                "File Name","")).strip())
                        ca=st.session_state.get(ck)
                        if ca and ca.get(
                                "price_num",0)>0:
                            all_prices_svc.append(
                                ca["price_num"])

                    rows=["<table class='comp-table'>"
                          "<thead><tr>"
                          "<th>Vendor</th>"
                          "<th>Category</th>"
                          "<th>File Name</th>"]
                    if has_price:
                        rows.append(
                            "<th>Quoted Price</th>")
                    rows.append(
                        "<th>Extracted</th>"
                        "<th>Score</th>"
                        "<th>Verdict</th>"
                        "<th>Open</th>"
                        "</tr></thead><tbody>")

                    for i,(idx,row) in enumerate(
                            d_svc.iterrows()):
                        bg=("white" if i%2==0
                            else "#F3F3F3")
                        vc=vendor_color_map.get(
                            row["Vendor"],"#8C8C8C")
                        fname=str(row.get(
                            "File Name","")).strip()
                        url=resolve_url(row)

                        qp_str=str(row.get(
                            "Quoted Price","")).strip()
                        qp_num=_parse_num(qp_str)

                        ck="px_{}".format(fname)
                        cached=st.session_state.get(ck)
                        if cached and cached.get(
                                "price_num",0)>0:
                            ep_num=cached["price_num"]
                            ep_fmt=_fmt(
                                cached["price"])
                            ref=ep_num
                        elif qp_num>0:
                            ep_fmt="—"
                            ref=qp_num
                        else:
                            ep_fmt="—"; ref=0

                        others=[
                            p for p in all_prices_svc
                            if p!=ref]
                        ps=None; vt="—"; vc_col="#bbb"
                        if ref>0 and others:
                            ps,_,_,_,_=price_score(
                                ref,others)
                            vt,_,vc_col=get_verdict(ps)

                        rows.append(
                            "<tr style='background:{}'>"
                            "<td>{}</td>"
                            "<td style='color:#555'>{}</td>"
                            "<td style='font-family:"
                            "monospace;font-size:0.79em;"
                            "word-break:break-all'>"
                            "{}</td>".format(
                                bg,
                                vendor_pill(
                                    row["Vendor"],vc),
                                row.get("Category",""),
                                fname))
                        if has_price:
                            rows.append(
                                "<td style='color:#22992E;"
                                "font-weight:700;"
                                "font-family:monospace'>"
                                "{}</td>".format(
                                    _fmt(qp_str)
                                    if qp_num>0
                                    else "—"))
                        sc=score_color(ps)
                        rows.append(
                            "<td style='color:#295477;"
                            "font-weight:700;"
                            "font-family:monospace'>"
                            "{}</td>"
                            "<td style='text-align:"
                            "center'><span style='"
                            "font-weight:800;"
                            "color:{}'>{}</span></td>"
                            "<td><span style='color:{};"
                            "font-weight:700;"
                            "font-size:0.82em'>"
                            "{}</span></td>"
                            "<td>{}</td></tr>".format(
                                ep_fmt,
                                sc,
                                "{}/100".format(ps)
                                if ps is not None
                                else "—",
                                vc_col,vt,
                                "<a href='{}' "
                                "target='_blank' "
                                "style='color:#D04A02;"
                                "font-weight:600;"
                                "text-decoration:none'>"
                                "Open</a>".format(url)
                                if url and url.startswith(
                                    "http")
                                else
                                "<a href='file://{}' "
                                "style='color:#D04A02;"
                                "font-weight:600;"
                                "text-decoration:none'>"
                                "📂 Local</a>".format(url)
                                if url
                                else "—"))

                    rows.append("</tbody></table>")
                    st.markdown(
                        "".join(rows),
                        unsafe_allow_html=True)

                    if all_prices_svc:
                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        ai_box(ai_price_insight(
                            0,all_prices_svc,
                            {v:p for v,p in
                             vendor_prices_map.items()
                             if p>0}))

                    st.markdown(
                        "<br>",unsafe_allow_html=True)
                    if st.button(
                        "Extract Prices — {}".format(
                            svc[:40]),
                        key="ep_{}".format(svc[:35]),
                        type="primary"):
                        prog=st.progress(0)
                        n=len(d_svc)
                        for ki,(idx2,row2) in enumerate(
                                d_svc.iterrows()):
                            fname2=str(row2.get(
                                "File Name","")).strip()
                            url2=resolve_url(row2)
                            ck2="px_{}".format(fname2)
                            if st.session_state.get(
                                    ck2) is None:
                                # Try local file first
                                local=os.path.join(
                                    DEMO_DIR,fname2)
                                if os.path.exists(local):
                                    res=extract_price_from_file(
                                        local)
                                    st.session_state[
                                        ck2]=res
                                elif (url2 and
                                      url2.startswith(
                                          "http")):
                                    from app import (
                                        extract_price_from_url
                                        if False else None)
                                    pass
                            prog.progress((ki+1)/n)
                        prog.empty()
                        st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 3 — UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
with tab3:
    if NO_DATA:
        st.info("No catalog. Go to 🗂 Upload Catalog.")
    else:
        st.markdown(
            "<div style='background:#2D2D2D;color:white;"
            "padding:14px 20px;border-radius:4px;"
            "border-left:6px solid #D04A02;"
            "margin-bottom:16px'>"
            "<div style='font-size:0.72em;"
            "font-weight:700;letter-spacing:2px;"
            "text-transform:uppercase;"
            "color:#D04A02;margin-bottom:4px'>"
            "New Quotation Analysis</div>"
            "<div style='font-size:1.0em;"
            "font-weight:700'>"
            "Upload a new quotation → auto-extract "
            "price → score &amp; verdict vs history"
            "</div></div>",
            unsafe_allow_html=True)

        if os.path.exists(
                os.path.join(DEMO_DIR,DEMO_QUOT)):
            st.markdown(
                "<div style='background:#FFF3F0;"
                "border:1px solid #D04A02;"
                "border-radius:4px;"
                "padding:10px 16px;margin-bottom:12px;"
                "font-size:0.85em'>"
                "🎯 <b>Demo file available:</b> "
                "Upload <code>demo_quotes/{}</code> "
                "to see a live verdict comparing "
                "Alpha Networks vs historical data."
                "</div>".format(DEMO_QUOT),
                unsafe_allow_html=True)

        section_title("STEP 1 — UPLOAD FILE")
        uploaded=st.file_uploader(
            "Upload",
            type=["pdf","xlsx","xls","docx"],
            label_visibility="collapsed")

        if uploaded is not None:
            content=uploaded.read()
            ext=uploaded.name.rsplit(".",1)[-1].lower()
            fname_up=uploaded.name
            st.success("Uploaded: **{}** ({} KB)".format(
                fname_up,round(len(content)/1024,1)))

            section_title("STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting price…"):
                result=extract_price_from_bytes(
                    content,ext)
                new_price=result["price_num"]
                new_text=result["text"]

            if new_price>0:
                st.markdown(
                    "<div class='score-card green'>"
                    "<div style='font-size:0.72em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#22992E'>Extracted Price</div>"
                    "<div style='font-size:2.0em;"
                    "font-weight:800;color:#22992E'>"
                    "{}</div>"
                    "</div>".format(_fmt(new_price)),
                    unsafe_allow_html=True)
            else:
                st.warning(
                    "Price not found automatically.")
                manual=st.number_input(
                    "Enter price manually (USD)",
                    min_value=0.0,step=100.0,value=0.0)
                if manual>0: new_price=manual

            section_title(
                "STEP 3 — SELECT SERVICES")
            all_svcs_up=sorted([
                s for s in df_exploded[
                    "Service"].unique()
                if str(s).strip() not in ["","nan"]])
            svc_search_up=st.text_input(
                "Filter",
                placeholder="Search services…",
                key="svc_up",
                label_visibility="collapsed")
            if svc_search_up:
                all_svcs_up=[
                    s for s in all_svcs_up
                    if svc_search_up.lower()
                    in s.lower()]

            new_services=st.multiselect(
                "Services in this quotation",
                options=all_svcs_up,
                key="new_svcs",
                label_visibility="collapsed")

            cat_filter_up=st.selectbox(
                "Filter historical by category",
                options=["All"]+sorted([
                    c for c in
                    df_master["Category"].unique()
                    if str(c).strip()
                    not in ["","nan"]]),
                key="cat_up")

            section_title(
                "STEP 4 — COMPARISON & VERDICT")

            if not new_services and new_price<=0:
                st.info(
                    "Select services above to compare.")
            else:
                if new_services:
                    candidates=df_exploded[
                        df_exploded["Service"].isin(
                            new_services)].copy()
                else:
                    candidates=df_exploded.copy()

                if cat_filter_up!="All":
                    candidates=candidates[
                        candidates["Category"]==
                        cat_filter_up]

                cand_files=(
                    candidates
                    .drop_duplicates(
                        subset=["File Name","Vendor"])
                    [["File Name","Vendor",
                      "Category","Hyperlink",
                      "Quoted Price"]]
                    .copy())

                if cand_files.empty:
                    st.warning(
                        "No historical quotes found.")
                else:
                    hist_prices=[]
                    vendor_p_map={}
                    for _,r in cand_files.iterrows():
                        qp=_parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp>0:
                            hist_prices.append(qp)
                            vendor_p_map[
                                r["Vendor"]]=qp
                        ck="px_{}".format(str(r.get(
                            "File Name","")).strip())
                        ca=st.session_state.get(ck)
                        if ca and ca.get(
                                "price_num",0)>0:
                            hist_prices.append(
                                ca["price_num"])
                            vendor_p_map[
                                r["Vendor"]]=\
                                ca["price_num"]

                    if new_price>0 and hist_prices:
                        ps,ps_lbl,avg_h,mn_h,mx_h=\
                            price_score(
                                new_price,hist_prices)
                        vt,vt_desc,vt_col=\
                            get_verdict(ps)

                        # Main verdict box
                        vt_css=score_css(ps)
                        st.markdown(
                            "<div class='verdict-{}'>"
                            "<div style='font-size:"
                            "1.2em;margin-bottom:6px'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.88em;font-weight:400'>"
                            "{}</div>"
                            "</div>".format(
                                vt_css,vt,vt_desc),
                            unsafe_allow_html=True)

                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)

                        # Score cards
                        sc1,sc2,sc3,sc4=st.columns(4)
                        sc1.markdown(
                            "<div class='score-card {}'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:{}'>Price Score"
                            "</div>"
                            "<div style='font-size:"
                            "2.0em;font-weight:800;"
                            "color:{}'>{}/100</div>"
                            "<div style='font-size:"
                            "0.75em;color:#555;"
                            "margin-top:4px'>vs {} "
                            "historical</div>"
                            "</div>".format(
                                vt_css,vt_col,vt_col,
                                ps if ps is not None
                                else "N/A",
                                len(hist_prices)),
                            unsafe_allow_html=True)
                        sc2.markdown(
                            "<div class='score-card "
                            "yellow'><div style='"
                            "font-size:0.68em;"
                            "font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:#856404'>Your Price"
                            "</div>"
                            "<div style='font-size:"
                            "1.8em;font-weight:800;"
                            "color:#D04A02'>{}</div>"
                            "</div>".format(
                                _fmt(new_price)),
                            unsafe_allow_html=True)
                        sc3.markdown(
                            "<div class='score-card "
                            "yellow'><div style='"
                            "font-size:0.68em;"
                            "font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:#856404'>"
                            "Historical Avg</div>"
                            "<div style='font-size:"
                            "1.8em;font-weight:800;"
                            "color:#295477'>{}</div>"
                            "<div style='font-size:"
                            "0.72em;color:#555;"
                            "margin-top:4px'>"
                            "min {} · max {}</div>"
                            "</div>".format(
                                _fmt(avg_h),
                                _fmt(mn_h),_fmt(mx_h)),
                            unsafe_allow_html=True)
                        sc4.markdown(
                            "<div class='score-card {}'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:{}'>vs Average</div>"
                            "<div style='font-size:"
                            "1.4em;font-weight:800;"
                            "color:{};margin-top:6px'>"
                            "{}</div>"
                            "</div>".format(
                                vt_css,vt_col,vt_col,
                                ps_lbl),
                            unsafe_allow_html=True)

                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        ai_box(ai_price_insight(
                            new_price,hist_prices,
                            vendor_p_map))

                        # Price chart
                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        section_title(
                            "PRICE POSITIONING CHART")
                        chart_data=[]
                        for _,r in cand_files.iterrows():
                            fn=str(r.get(
                                "File Name","")).strip()
                            ck="px_{}".format(fn)
                            cached=st.session_state.get(
                                ck)
                            qp=_parse_num(str(r.get(
                                "Quoted Price",
                                "")).strip())
                            ep=(cached["price_num"]
                                if cached else 0.0)
                            pval=ep if ep>0 else qp
                            if pval>0:
                                chart_data.append({
                                    "Label":
                                        "{}/{}".format(
                                            r["Vendor"],
                                            fn[:12]),
                                    "Price":pval,
                                    "Type":"Historical",
                                    "Color":
                                        vendor_color_map
                                        .get(r["Vendor"],
                                             "#8C8C8C"),
                                })
                        chart_data.append({
                            "Label":"★ NEW:{}".format(
                                fname_up[:15]),
                            "Price":new_price,
                            "Type":"New Upload",
                            "Color":"#D04A02",
                        })
                        cdf=pd.DataFrame(
                            chart_data).sort_values(
                                "Price")
                        cf=go.Figure()
                        hist_df=cdf[
                            cdf["Type"]=="Historical"]
                        new_df=cdf[
                            cdf["Type"]=="New Upload"]
                        if not hist_df.empty:
                            cf.add_trace(go.Bar(
                                x=hist_df["Label"],
                                y=hist_df["Price"],
                                marker_color=hist_df[
                                    "Color"],
                                marker_line_width=0,
                                name="Historical",
                                text=hist_df[
                                    "Price"].apply(
                                        _fmt),
                                textposition="outside"))
                        if not new_df.empty:
                            cf.add_trace(go.Bar(
                                x=new_df["Label"],
                                y=new_df["Price"],
                                marker_color="#D04A02",
                                marker_line_width=0,
                                name="Your Upload",
                                text=new_df[
                                    "Price"].apply(
                                        _fmt),
                                textposition="outside"))
                        cf.add_hline(
                            y=avg_h,
                            line_dash="dash",
                            line_color="#FFB600",
                            line_width=2,
                            annotation_text="Avg:{}".format(
                                _fmt(avg_h)),
                            annotation_position=
                                "top right")
                        cf.update_layout(
                            height=360,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(
                                l=5,r=10,t=20,b=10),
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
                                x=0,y=1.05),
                            bargap=0.25)
                        st.plotly_chart(
                            cf,use_container_width=True)


# ════════════════════════════════════════════════════════════
# TAB 4 — DATA TABLE
# ════════════════════════════════════════════════════════════
with tab4:
    if NO_DATA:
        st.info("No catalog. Go to 🗂 Upload Catalog.")
    else:
        dm=df_master.copy()
        if selected_cat!="All":
            dm=dm[dm["Category"]==selected_cat]
        if selected_vendor!="All":
            dm=dm[dm["Vendor"]==selected_vendor]
        st.dataframe(
            dm.drop(
                columns=["Services List","Hyperlink"],
                errors="ignore"),
            use_container_width=True,
            height=500)


# ════════════════════════════════════════════════════════════
# TAB 5 — UPLOAD CATALOG
# ════════════════════════════════════════════════════════════
with tab5:
    st.markdown(
        "<div style='background:#2D2D2D;color:white;"
        "padding:20px 28px;border-radius:4px;"
        "border-left:6px solid #D04A02;"
        "margin-bottom:20px'>"
        "<div style='font-size:0.72em;font-weight:700;"
        "letter-spacing:2px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:5px'>"
        "Catalog Management</div>"
        "<h1 style='margin:0;font-size:1.2em;"
        "font-weight:700;color:white'>"
        "Upload Master Catalog</h1>"
        "<p style='margin:6px 0 0;opacity:0.6;"
        "font-size:0.85em'>"
        "Upload any Excel or CSV quotation catalog — "
        "AI auto-detects columns and builds "
        "the dashboard automatically."
        "</p></div>",
        unsafe_allow_html=True)

    if DATA_SOURCE=="uploaded":
        st.success(
            "✅ Using uploaded catalog: "
            "**{}** rows · **{}** vendors · "
            "**{}** services".format(
                len(df_master),
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique()))

    catalog_file=st.file_uploader(
        "Upload Master Catalog",
        type=["xlsx","xls","csv"],
        label_visibility="collapsed",
        key="catalog_upload")

    if catalog_file is not None:
        file_bytes=catalog_file.read()
        fname_cat=catalog_file.name

        with st.spinner("AI analyzing catalog…"):
            df_new,df_exp_new,err=\
                process_uploaded_catalog(
                    file_bytes,fname_cat)

        if err:
            st.error("❌ {}".format(err))
        elif df_new is None:
            st.error("❌ Could not process file.")
        else:
            st.success(
                "✅ Detected **{}** rows · "
                "**{}** vendors · **{}** categories".format(
                    len(df_new),
                    df_new["Vendor"].nunique(),
                    df_new["Category"].nunique()))

            # Preview
            pc1,pc2=st.columns(2)
            spv_new=(df_exp_new.groupby(
                "Vendor")["Service"].nunique()
                .sort_values(ascending=False)
                .reset_index())
            spv_new.columns=["Vendor","Services"]
            with pc1:
                pf1=go.Figure(go.Bar(
                    x=spv_new["Vendor"],
                    y=spv_new["Services"],
                    marker_color=[
                        get_color(i)
                        for i in range(
                            len(spv_new))],
                    marker_line_width=0,
                    text=spv_new["Services"],
                    textposition="outside"))
                pf1.update_layout(
                    title="Services per Vendor",
                    height=300,
                    plot_bgcolor=CBG,
                    paper_bgcolor=CBG,
                    margin=dict(l=5,r=10,t=40,b=10),
                    font=CFONT,
                    yaxis=dict(
                        showgrid=True,
                        gridcolor="#E0E0E0",
                        zeroline=False),
                    xaxis=dict(tickangle=-30),
                    bargap=0.35)
                st.plotly_chart(
                    pf1,use_container_width=True)

            cat_new=(df_new.drop_duplicates(
                subset=["Category","File Name"])
                .groupby("Category").size()
                .reset_index())
            cat_new.columns=["Category","Count"]
            with pc2:
                if not cat_new.empty:
                    pf2=px.pie(
                        cat_new,
                        names="Category",
                        values="Count",
                        hole=0.45,
                        color_discrete_sequence=COLORS)
                    pf2.update_traces(
                        textposition="outside",
                        textinfo="label+percent",
                        textfont_size=10)
                    pf2.update_layout(
                        title="Category Distribution",
                        height=300,
                        margin=dict(
                            l=10,r=10,t=40,b=10),
                        paper_bgcolor=CBG,font=CFONT)
                    st.plotly_chart(
                        pf2,use_container_width=True)

            st.dataframe(
                df_new.drop(
                    columns=["Services List",
                              "Hyperlink"],
                    errors="ignore").head(20),
                use_container_width=True,
                height=280)

            if st.button(
                "✅ Apply This Catalog to Dashboard",
                type="primary",
                use_container_width=False):
                st.session_state[
                    "uploaded_catalog_df"]=df_new
                st.session_state[
                    "uploaded_catalog_exp"]=df_exp_new
                st.success(
                    "✅ Catalog applied! "
                    "Dashboard updated.")
                st.rerun()
