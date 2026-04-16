import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* Remove default Streamlit padding */
  .block-container {
      padding-top: 1.5rem;
      padding-bottom: 1rem;
  }

  /* ── Section headings ── */
  .section-heading {
      font-size: 1.25rem;
      font-weight: 700;
      color: #1f2937;
      margin-bottom: 0.5rem;
      margin-top: 1rem;
      line-height: 1.4;
      display: block;          /* prevents inline-overlap issues */
  }

  /* ── Page title ── */
  .page-title {
      font-size: 2rem;
      font-weight: 800;
      color: #111827;
      margin-bottom: 0.25rem;
      line-height: 1.3;
      display: block;
  }

  .page-subtitle {
      font-size: 1rem;
      color: #6b7280;
      margin-bottom: 1.5rem;
      display: block;
  }

  /* ── Metric cards ── */
  .metric-card {
      background: #ffffff;
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      padding: 1.25rem 1.5rem;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06);
  }

  .metric-label {
      font-size: 0.85rem;
      font-weight: 600;
      color: #6b7280;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      margin-bottom: 0.4rem;
  }

  .metric-value {
      font-size: 2rem;
      font-weight: 800;
      color: #111827;
      line-height: 1.2;
  }

  .metric-delta-pos {
      font-size: 0.85rem;
      color: #10b981;
      font-weight: 600;
  }

  .metric-delta-neg {
      font-size: 0.85rem;
      color: #ef4444;
      font-weight: 600;
  }

  /* ── Chart container ── */
  .chart-container {
      background: #ffffff;
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      padding: 1.25rem;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06);
      margin-bottom: 1rem;
  }

  .chart-title {
      font-size: 1rem;
      font-weight: 700;
      color: #1f2937;
      margin-bottom: 0.75rem;
      display: block;          /* no overlap with adjacent elements */
  }

  /* ── Sidebar ── */
  [data-testid="stSidebar"] {
      background: #f9fafb;
      border-right: 1px solid #e5e7eb;
  }

  .sidebar-title {
      font-size: 1.1rem;
      font-weight: 700;
      color: #1f2937;
      margin-bottom: 1rem;
      display: block;
  }

  /* ── Divider ── */
  .section-divider {
      border: none;
      border-top: 1px solid #e5e7eb;
      margin: 1.5rem 0;
  }

  /* Hide Streamlit default menu/footer if desired */
  #MainMenu {visibility: hidden;}
  footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
@st.cache_data
def generate_data(n: int = 500, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    start = datetime(2023, 1, 1)
    dates = [start + timedelta(days=int(d)) for d in rng.integers(0, 365, n)]
    regions = rng.choice(["North", "South", "East", "West"], n)
    categories = rng.choice(["Electronics", "Clothing", "Food", "Furniture", "Sports"], n)
    sales = rng.integers(100, 5000, n).astype(float)
    profit = sales * rng.uniform(0.05, 0.35, n)
    units = rng.integers(1, 50, n)
    return pd.DataFrame({
        "date": dates,
        "region": regions,
        "category": categories,
        "sales": sales,
        "profit": profit,
        "units": units,
    })


def fmt_currency(value: float) -> str:
    if value >= 1_000_000:
        return f"${value/1_000_000:.1f}M"
    if value >= 1_000:
        return f"${value/1_000:.1f}K"
    return f"${value:.0f}"


def metric_card(label: str, value: str, delta: str | None = None, positive: bool = True):
    delta_html = ""
    if delta:
        cls = "metric-delta-pos" if positive else "metric-delta-neg"
        arrow = "▲" if positive else "▼"
        delta_html = f'<div class="{cls}">{arrow} {delta}</div>'
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


def section_heading(text: str):
    """Render a clean section heading with no overlapping arrows or icons."""
    st.markdown(f'<span class="section-heading">{text}</span>', unsafe_allow_html=True)


def chart_title(text: str):
    st.markdown(f'<span class="chart-title">{text}</span>', unsafe_allow_html=True)


# ── Load data ──────────────────────────────────────────────────────────────────
df = generate_data()
df["month"] = df["date"].dt.to_period("M").dt.to_timestamp()
df["month_label"] = df["date"].dt.strftime("%b %Y")

# ── Sidebar filters ────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<span class="sidebar-title">🔍 Filters</span>', unsafe_allow_html=True)

    all_regions = sorted(df["region"].unique())
    selected_regions = st.multiselect(
        "Region",
        options=all_regions,
        default=all_regions
    )

    all_categories = sorted(df["category"].unique())
    selected_categories = st.multiselect(
        "Category",
        options=all_categories,
        default=all_categories
    )

    min_date = df["date"].min().date()
    max_date = df["date"].max().date()
    date_range = st.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    st.markdown("**About**")
    st.caption("Sales Dashboard v1.0  \nData is synthetic for demo purposes.")

# ── Apply filters ──────────────────────────────────────────────────────────────
filtered = df.copy()
if selected_regions:
    filtered = filtered[filtered["region"].isin(selected_regions)]
if selected_categories:
    filtered = filtered[filtered["category"].isin(selected_categories)]
if len(date_range) == 2:
    start_d, end_d = date_range
    filtered = filtered[
        (filtered["date"].dt.date >= start_d) &
        (filtered["date"].dt.date <= end_d)
    ]

# ── Page title ────────────────────────────────────────────────────────────────
st.markdown('<span class="page-title">📊 Sales Dashboard</span>', unsafe_allow_html=True)
st.markdown('<span class="page-subtitle">Track performance across regions, categories, and time periods.</span>',
            unsafe_allow_html=True)

# ── KPI row ───────────────────────────────────────────────────────────────────
section_heading("Key Performance Indicators")

total_sales   = filtered["sales"].sum()
total_profit  = filtered["profit"].sum()
total_units   = filtered["units"].sum()
margin_pct    = (total_profit / total_sales * 100) if total_sales else 0
avg_order     = (total_sales / len(filtered)) if len(filtered) else 0

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    metric_card("Total Sales",   fmt_currency(total_sales),  "12.4% vs last period", positive=True)
with col2:
    metric_card("Total Profit",  fmt_currency(total_profit), "8.1% vs last period",  positive=True)
with col3:
    metric_card("Units Sold",    f"{total_units:,}",         "5.3% vs last period",  positive=True)
with col4:
    metric_card("Profit Margin", f"{margin_pct:.1f}%",       "1.2% vs last period",  positive=False)
with col5:
    metric_card("Avg Order",     fmt_currency(avg_order),    "3.7% vs last period",  positive=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Trend & Category breakdown ────────────────────────────────────────────────
section_heading("Sales Trends & Category Breakdown")

col_left, col_right = st.columns([2, 1], gap="medium")

with col_left:
    with st.container():
        chart_title("Monthly Sales & Profit")
        monthly = (
            filtered.groupby("month")[["sales", "profit"]]
            .sum()
            .reset_index()
            .sort_values("month")
        )
        monthly["month_str"] = monthly["month"].dt.strftime("%b %Y")

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Bar(
            x=monthly["month_str"], y=monthly["sales"],
            name="Sales", marker_color="#6366f1", opacity=0.85
        ))
        fig_trend.add_trace(go.Scatter(
            x=monthly["month_str"], y=monthly["profit"],
            name="Profit", mode="lines+markers",
            line=dict(color="#f59e0b", width=2.5),
            marker=dict(size=6)
        ))
        fig_trend.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(showgrid=False, tickfont=dict(size=11)),
            yaxis=dict(showgrid=True, gridcolor="#f3f4f6", tickfont=dict(size=11)),
            height=320
        )
        st.plotly_chart(fig_trend, use_container_width=True)

with col_right:
    with st.container():
        chart_title("Sales by Category")
        cat_data = (
            filtered.groupby("category")["sales"]
            .sum()
            .reset_index()
            .sort_values("sales", ascending=False)
        )
        fig_pie = px.pie(
            cat_data, values="sales", names="category",
            color_discrete_sequence=px.colors.qualitative.Pastel,
            hole=0.45
        )
        fig_pie.update_traces(textposition="outside", textinfo="percent+label")
        fig_pie.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            showlegend=False,
            plot_bgcolor="white", paper_bgcolor="white",
            height=320
        )
        st.plotly_chart(fig_pie, use_container_width=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Regional performance ───────────────────────────────────────────────────────
section_heading("Regional Performance")

col_a, col_b = st.columns(2, gap="medium")

with col_a:
    chart_title("Sales by Region")
    region_data = (
        filtered.groupby("region")["sales"]
        .sum()
        .reset_index()
        .sort_values("sales", ascending=True)
    )
    fig_bar = px.bar(
        region_data, x="sales", y="region", orientation="h",
        color="sales",
        color_continuous_scale=["#c7d2fe", "#6366f1", "#312e81"],
        labels={"sales": "Sales ($)", "region": ""}
    )
    fig_bar.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        coloraxis_showscale=False,
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="#f3f4f6"),
        yaxis=dict(showgrid=False),
        height=280
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col_b:
    chart_title("Profit Margin by Region")
    region_margin = (
        filtered.groupby("region")[["sales", "profit"]]
        .sum()
        .reset_index()
    )
    region_margin["margin"] = region_margin["profit"] / region_margin["sales"] * 100
    fig_margin = px.bar(
        region_margin.sort_values("margin", ascending=True),
        x="margin", y="region", orientation="h",
        color="margin",
        color_continuous_scale=["#fde68a", "#f59e0b", "#92400e"],
        labels={"margin": "Margin (%)", "region": ""}
    )
    fig_margin.update_layout(
        margin=dict(t=10, b=10, l=10, r=10),
        coloraxis_showscale=False,
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="#f3f4f6"),
        yaxis=dict(showgrid=False),
        height=280
    )
    st.plotly_chart(fig_margin, use_container_width=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Category × Region heatmap ─────────────────────────────────────────────────
section_heading("Category × Region Heatmap")

pivot = (
    filtered.groupby(["category", "region"])["sales"]
    .sum()
    .unstack(fill_value=0)
)
fig_heat = px.imshow(
    pivot,
    color_continuous_scale="Blues",
    aspect="auto",
    labels=dict(color="Sales ($)")
)
fig_heat.update_layout(
    margin=dict(t=10, b=10, l=10, r=10),
    plot_bgcolor="white", paper_bgcolor="white",
    height=300,
    xaxis_title="",
    yaxis_title=""
)
st.plotly_chart(fig_heat, use_container_width=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Scatter: sales vs profit ───────────────────────────────────────────────────
section_heading("Sales vs. Profit Scatter")

fig_scatter = px.scatter(
    filtered, x="sales", y="profit",
    color="category", size="units",
    hover_data=["region", "date"],
    color_discrete_sequence=px.colors.qualitative.Bold,
    labels={"sales": "Sales ($)", "profit": "Profit ($)"},
    opacity=0.7
)
fig_scatter.update_layout(
    margin=dict(t=10, b=10, l=10, r=10),
    plot_bgcolor="white", paper_bgcolor="white",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    height=360
)
st.plotly_chart(fig_scatter, use_container_width=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Data table ────────────────────────────────────────────────────────────────
section_heading("Raw Data Explorer")

show_table = st.toggle("Show data table", value=False)
if show_table:
    display_df = (
        filtered[["date", "region", "category", "sales", "profit", "units"]]
        .sort_values("date", ascending=False)
        .reset_index(drop=True)
    )
    display_df["sales"]  = display_df["sales"].map("${:,.0f}".format)
    display_df["profit"] = display_df["profit"].map("${:,.0f}".format)

    st.dataframe(display_df, use_container_width=True, height=400)

    csv = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇ Download CSV",
        data=csv,
        file_name="sales_data.csv",
        mime="text/csv"
    )
