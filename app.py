import math

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
from io import BytesIO

st.set_page_config(
    page_title="NG Disposal Cost Dashboard",
    page_icon="📉",
    layout="wide",
    initial_sidebar_state="expanded",
)

TARGET_DEFAULT = 0.005
DEFAULT_FILES = [
    Path("Disposal_cost.xlsb"),
    Path("Disposal_cost Update 20261013 Ver.00 (New)(1).xlsb"),
]

# ---------- Global styling ----------
st.markdown(
    """

<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;600;700;800;900&display=swap');

    :root { --dashboard-font: 'Roboto', sans-serif; }

    html, body, #root, .stApp, .stApp *,
    [class*='st-'], [class*='css-'],
    button, input, textarea, select, option, label, p, span, div,
    h1, h2, h3, h4, h5, h6, a, li, table, th, td {
        font-family: var(--dashboard-font) !important;
    }

    /* Streamlit's generated containers */
    [data-testid='stAppViewContainer'] *,
    [data-testid='stMain'] *,
    [data-testid='stSidebar'] *,
    [data-testid='stHeader'] *,
    [data-testid='stToolbar'] *,
    [data-testid='stMarkdownContainer'] *,
    [data-testid='stCaptionContainer'] *,
    [data-testid='stMetric'] *,
    [data-testid='stExpander'] *,
    [data-testid='stFileUploader'] * {
        font-family: var(--dashboard-font) !important;
    }

    /* Plotly SVG text */
    .js-plotly-plot text,
    .js-plotly-plot .legendtext,
    .js-plotly-plot .gtitle,
    .js-plotly-plot .xtick text,
    .js-plotly-plot .ytick text,
    .js-plotly-plot .pie text {
        font-family: 'Roboto' !important;
    }

    .stApp {
        background: #f3f4f6;
    }

    [data-testid="stHeader"] {
        background: transparent;
    }

    [data-testid="stToolbar"] {
        right: 1rem;
    }

    .block-container {
    max-width: 1600px !important;
    padding-top: 1.2rem !important;
    padding-bottom: 1.2rem !important;
    padding-left: 1.5rem !important;
    padding-right: 1.5rem !important;
}

    /* Header */
    .dashboard-header {
        background: #f3f4f6;
        border-bottom: 1px solid #d9dde3;
        padding: 4px 8px 12px 0;
        margin-bottom: 14px;
    }

    .title-row {
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .title-bar {
        width: 9px;
        height: 56px;
        background: #d71920;
        border-radius: 1px;
        flex: 0 0 auto;
    }

    .dashboard-title {
        font-size: 32px;
        line-height: 1.05;
        font-weight: 800;
        color: #111827;
        margin: 0;
        padding: 0;
    }

    .header-meta {
        text-align: left;
        font-size: 14px;
        line-height: 1.8;
        color: #111827;
        padding-top: 2px;
    }

    .header-meta b {
        display: inline-block;
        min-width: 88px;
    }

    /* KPI cards */
    .kpi-card {
    background: rgba(255, 255, 255, 0.97);
    border: 1px solid #e4e7f0;
    border-radius: 16px;
    min-height: 145px;
    padding: 16px 20px 14px 20px;
    box-shadow: 0 4px 18px rgba(70, 80, 140, 0.08);
    text-align: center;
    box-sizing: border-box;
}

    .kpi-title {
        font-size: 21px;
        line-height: 1.15;
        font-weight: 700;
        color: #111827;
        margin-bottom: 11px;
    }

    .kpi-value {
        font-size: 39px;
        line-height: 1.05;
        font-weight: 800;
        color: #111827;
        margin-bottom: 9px;
    }

    .kpi-sub {
        font-size: 16px;
        color: #222;
    }

    .status-ok {
        display: inline-block;
        background: #12a34a;
        color: #fff !important;
        font-weight: 700;
        padding: 3px 12px;
        margin-left: 4px;
    }

    .status-ng {
        display: inline-block;
        background: #d71920;
        color: #fff !important;
        font-weight: 700;
        padding: 3px 12px;
        margin-left: 4px;
    }

    /* Section cards */
    .warehouse-chart {
        background:#fff;
        min-height:250px;
        height:250px;
        padding:8px 8px 4px 0;
        box-sizing:border-box;
        font-family:Roboto, Arial, sans-serif;
        overflow:visible;
    }

    .warehouse-rows {
        display:flex;
        flex-direction:column;
        gap:8px;
        padding-top:4px;
    }

    .warehouse-row {
        display:grid;
        grid-template-columns:126px minmax(0, 1fr);
        align-items:center;
        min-height:32px;
        column-gap:5px;
    }

    .warehouse-name {
        font-size:11px;
        color:#6b7280;
        text-align:right;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
    }

    .warehouse-track {
        position:relative;
        height:19px;
        background:repeating-linear-gradient(
            to right,
            transparent 0,
            transparent calc(25% - .5px),
            #e5e7eb calc(25% - .5px),
            #e5e7eb calc(25% + .5px)
        );
        overflow:visible;
    }

    .warehouse-bar {
        height:19px;
        min-width:2px;
        border-radius:0 8px 8px 0;
        background:linear-gradient(180deg, #3b82c4 0%, #2f75b5 55%, #245f95 100%);
        box-shadow:inset 0 1px 1px rgba(255,255,255,.35), 0 1px 2px rgba(0,0,0,.12);
    }

    .warehouse-bar-max {
        background:linear-gradient(180deg, #ff4b4b 0%, #f00000 55%, #c40000 100%);
        box-shadow:inset 0 1px 1px rgba(255,255,255,.38), 0 1px 3px rgba(180,0,0,.25);
    }

    .warehouse-value {
        position:absolute;
        top:50%;
        transform:translate(7px,-50%);
        font-size:11px;
        color:#4b5563;
        font-weight:500;
        white-space:nowrap;
        text-align:left;
        z-index:2;
    }

    .warehouse-axis {
        position:relative;
        margin-left:131px;
        margin-right:8px;
        height:25px;
        margin-top:2px;
    }

    .warehouse-axis-line {
        position:absolute;
        left:0;
        right:0;
        top:0;
        border-top:1px solid #e5e7eb;
    }

    .warehouse-tick {
        position:absolute;
        top:5px;
        transform:translateX(-50%);
        font-size:10px;
        color:#6b7280;
        white-space:nowrap;
    }

    .panel {
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid #e3e6f0;
    border-radius: 16px;
    padding: 12px 14px 10px 14px;
    margin-bottom: 8px;
    box-shadow: 0 4px 18px rgba(70, 80, 140, 0.08);
    box-sizing: border-box;
    min-height: 300px;
}

    .panel-title {
        font-size: 20px;
        font-weight: 700;
        color: #111827;
        margin: 0 0 0 0;
    }

    /* Process / Material / P-Material - reference progress ring */
    .mini-label {
        font-size:18px;
        font-weight:700;
        color:#a40000;
        text-align:center;
        margin-top:10px;
        margin-bottom:6px;
    }

    .process-donut-wrap {
    display: flex;
    justify-content: center;
    align-items: center;
    margin-top: 4px;
    margin-bottom: 12px;
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
}
    .process-donut {
    width:170px;
    height:170px;
    position:relative;
    display:flex;
    align-items:center;
    justify-content:center;
    filter:drop-shadow(0 4px 5px rgba(0,0,0,.14));
    z-index:1;
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
}

    .process-donut::before {
        content:"";
        position:absolute;
        inset:-7px;
        border-radius:50%;
        border:1px solid rgba(165,0,0,.28);
        box-shadow:
            0 1px 2px rgba(0,0,0,.08),
            inset 0 0 0 1px rgba(255,255,255,.45);
        pointer-events:none;
        z-index:0;
    }

    .process-donut-inner {
        position:absolute;
        width:118px;
        height:118px;
        left:50%;
        top:50%;
        transform:translate(-50%,-50%);
        border-radius:50%;
        background:linear-gradient(145deg,#fafafa 0%,#f1f2f3 55%,#e6e8ea 100%);
        box-shadow:
            inset 0 1px 3px rgba(255,255,255,.95),
            inset 0 -2px 4px rgba(0,0,0,.10),
            0 1px 2px rgba(0,0,0,.08);
        display:flex;
        flex-direction:column;
        justify-content:center;
        align-items:center;
        z-index:2;
    }

    .process-number {
        font-size:18px;
        line-height:1.1;
        font-weight:800;
        color:#111827;
    }

    .process-unit {
        font-size:14px;
        line-height:1.2;
        font-weight:600;
        color:#111827;
        margin-top:3px;
    }

.process-value {
        font-size: 16px;
        font-weight: 700;
        color: #111827;
    }

    .info-box {
        background: #eaf3ff;
        border: 1px solid #c9def8;
        color: #0f4c81;
        border-radius: 3px;
        padding: 10px 12px;
        font-size: 14px;
    }

    .footer {
        color: #8a8f98;
        font-size: 13px;
        padding-top: 5px;
    }

    /* Bottom dashboard row */
    .trend-panel, .sep-panel {
        background:#fff;
        border:1px solid #cfd6de;
        border-radius:9px;
        overflow:hidden;
        box-shadow:0 3px 9px rgba(0,0,0,.12), inset 0 1px 0 rgba(255,255,255,.95);
    }
    .sep-panel {
        margin-top:0;
    }
    .trend-title, .sep-title {
        font-size:17px;
        font-weight:700;
        color:#111827;
        padding:3px 8px;
        line-height:1.25;
    }
    .sep-title {
        background:linear-gradient(180deg,#245b86 0%,#173f63 52%,#123653 100%);
        color:#fff;
        height:29px;
        box-sizing:border-box;
        padding:4px 9px;
        border-bottom:1px solid #0d2d45;
        text-shadow:0 1px 1px rgba(0,0,0,.25);
    }
    .sep-item {
        min-height:53px;
        padding:4px 8px 5px;
        border-bottom:1px solid #e5e7eb;
        box-sizing:border-box;
        transition:background .15s ease;
    }
    .sep-item:last-child { border-bottom:0; }
    .sep-label {
        font-size:13px;
        line-height:1.15;
        font-weight:700;
        color:#111827;
        margin-bottom:4px;
    }
    .sep-row {
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:8px;
    }
    .sep-value {
        font-size:15px;
        line-height:1;
        font-weight:800;
        color:#f00000;
    }
    .sep-pct {
        position:relative;
        width:112px;
        height:22px;
        flex:0 0 112px;
        box-sizing:border-box;
        border:0;
        border-radius:12px;
        background:transparent;
        box-shadow:none;
        font-size:12px;
        font-weight:700;
        color:#222;
        overflow:visible;
        padding:0;
    }

    .sep-pct::before {
        content:"";
        position:absolute;
        left:-8px;
        right:40px;
        top:50%;
        height:6px;
        transform:translateY(-50%);
        border-radius:6px;
        background:linear-gradient(180deg,#e7e8ee 0%,#d7d9e1 50%,#cdd0d9 100%);
    }

    .sep-pct-fill {
        position:absolute;
        left:-8px;
        top:50%;
        height:5px;
        min-width:0;
        transform:translateY(-50%);
        background:linear-gradient(
            90deg,
            #ff5b62 0%,
            #ff3f48 45%,
            #ff2633 75%,
            #f01827 100%
        );
        border-radius:6px;
        box-shadow:none;
        opacity:1;
    }

    .sep-pct-fill::after {
        content:"";
        position:absolute;
        right:-6px;
        top:50%;
        width:13px;
        height:13px;
        transform:translateY(-50%);
        border-radius:50%;
        background:linear-gradient(145deg,#ff7078 0%,#ff3542 52%,#ed1020 100%);
        box-shadow:0 1px 2px rgba(0,0,0,.20);
    }

    .sep-pct-text {
        position:absolute;
        right:-2px;
        top:50%;
        transform:translateY(-50%);
        z-index:2;
        color:#222;
        font-weight:700;
        text-shadow:none;
        white-space:nowrap;
        min-width:31px;
        text-align:right;
        margin:0;
    }

    .sep-pct-text.light {
        color:#222;
        text-shadow:none;
    }

    .sep-item:nth-child(2) { background:#fff; }
    .sep-item:nth-child(3) { background:#fff8dc; }
    .sep-item:nth-child(4) { background:#dff2fb; }
    .sep-item:nth-child(5) { background:#dff2fb; }
    .sep-item:nth-child(6) { background:#fff; }

    /* Streamlit chart / spacing */
    div[data-testid="stVerticalBlock"] > div:has(> div.stPlotlyChart) {
        margin-bottom: 0 !important;
    }

    div[data-testid="stHorizontalBlock"] {
        gap: 0.7rem;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #ffffff;
    }

     /* Hide the default divider created by st.divider when not needed */
    hr {
        border-color: #d9dde3 !important;
    }


    /* =========================================================
       DASHBOARD LAYOUT
       ========================================================= */

    .stApp {
    background: #f1f1f1 !important;
}

    .block-container {
        max-width: 1500px !important;
        padding: 24px 24px 20px 24px !important;
    }


    /* ---------- MAIN CARDS ---------- */

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card,
    .st-key-trend_card,
    .st-key-october_card {

        background: #ffffff !important;

        border: 1px solid #e7e9f1 !important;

        border-radius: 16px !important;

        box-shadow:
            0 5px 18px rgba(80,90,150,.08) !important;

        box-sizing: border-box !important;

        overflow: hidden !important;
    }


    /* ---------- MIDDLE ROW ---------- */

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {

        height: 330px !important;

        min-height: 330px !important;

        padding: 14px 14px 8px 14px !important;
    }


    /* ---------- BOTTOM ROW ---------- */

.st-key-trend_card,
.st-key-october_card {

    height: 350px !important;
    min-height: 350px !important;
    max-height: 350px !important;

    padding: 14px 14px 8px 14px !important;

    box-sizing: border-box !important;

    align-self: stretch !important;

    display: flex !important;
    flex-direction: column !important;

    overflow: hidden !important;
}


    /* ---------- CARD TITLE ---------- */

    .reference-card-title {

        font-size: 30px;

        line-height: 1.2;

        font-weight: 700;

        color: #10183d;

        margin: 0 0 8px 0;

        padding: 0;
    }


    /* ---------- PROCESS LABEL ---------- */

    .reference-process-label {

        font-size: 15px;

        font-weight: 700;

        color: #c40000;

        text-align: center;

        margin-bottom: 2px;
    }


    /* ---------- PLOTLY ---------- */

    .st-key-process_card .stPlotlyChart,
    .st-key-warehouse_card .stPlotlyChart,
    .st-key-model_card .stPlotlyChart,
    .st-key-trend_card .stPlotlyChart {

        margin: 0 !important;

        padding: 0 !important;
    }


    /* ---------- OCTOBER ---------- */

    .reference-oct-title {

        font-size: 19px;

        font-weight: 700;

        color: #10183d;

        margin-bottom: 18px;
    }


    .reference-oct-item {

        margin-bottom: 15px;
    }


    .reference-oct-name {

        font-size: 13px;

        font-weight: 600;

        color: #10183d;

        margin-bottom: 5px;
    }


    .reference-oct-row {

        display: flex;

        align-items: center;

        gap: 10px;
    }


    .reference-oct-value {

        width: 125px;

        flex-shrink: 0;

        font-size: 15px;

        font-weight: 800;

        color: #ef233c;
    }


    .reference-oct-bar {

        position: relative;

        flex: 1;

        height: 8px;

        background: #e5e7ef;

        border-radius: 10px;
    }


    .reference-oct-fill {

        position: absolute;

        left: 0;

        top: 0;

        height: 8px;

        background: #ff3344;

        border-radius: 10px;
    }


    .reference-oct-dot {

        position: absolute;

        top: 50%;

        width: 14px;

        height: 14px;

        background: #ff3344;

        border-radius: 50%;

        transform: translate(-50%,-50%);
    }


    .reference-oct-pct {

        width: 38px;

        text-align: right;

        font-size: 13px;

        color: #111827;

        font-weight: 600;
    }


    /* ---------- COLUMN GAP ---------- */

    div[data-testid="stHorizontalBlock"] {
    gap: 14px !important;
}

/* =========================================================
   KPI CARDS - SAME SIZE
   ========================================================= */

.st-key-k1,
.st-key-k2,
.st-key-k3,
.st-key-k4 {
    height: 116px !important;
    min-height: 116px !important;
    max-height: 116px !important;

    box-sizing: border-box !important;

    display: flex !important;
    flex-direction: column !important;

    justify-content: center !important;
}

    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("""<style>
:root { --nav-width: 78px; --nav-blue:#3559e6; }
.stApp { background:#f4f6fb !important; }
section[data-testid="stSidebar"] { display:none !important; }
[data-testid="stAppViewContainer"] { margin-left:var(--nav-width) !important; }
.block-container { max-width:1500px !important; padding:22px 26px 18px 26px !important; }
.reference-nav { position:fixed; left:0; top:0; bottom:0; width:var(--nav-width); background:rgba(255,255,255,.96); border-right:1px solid #e9ecf5; z-index:999999; display:flex; flex-direction:column; align-items:center; padding-top:22px; box-sizing:border-box; }
.reference-nav-logo { width:38px; height:38px; border-radius:11px; background:linear-gradient(145deg,#6678f2,#3559e6); color:#fff; display:flex; align-items:center; justify-content:center; font-size:21px; font-weight:800; margin-bottom:34px; box-shadow:0 5px 14px rgba(53,89,230,.22); }
.reference-nav-item { width:44px; height:44px; border-radius:12px; display:flex; align-items:center; justify-content:center; color:#7180a8; font-size:21px; margin-bottom:13px; box-sizing:border-box; }
.reference-nav-item.active { background:#eef2ff; color:var(--nav-blue); }
.reference-nav-spacer { flex:1; }
.dashboard-header { background:transparent !important; border-bottom:1px solid #dfe3ec !important; padding-bottom:13px !important; margin-bottom:18px !important; }
.dashboard-title { font-size:31px !important; font-weight:800 !important; color:#0d1735 !important; }
.header-meta { font-size:12px !important; line-height:1.65 !important; color:#4b5876 !important; }
.st-key-k1,.st-key-k2,.st-key-k3,.st-key-k4 { height:148px !important; min-height:148px !important; max-height:148px !important; }
.kpi-card { min-height:148px !important; border-radius:16px !important; padding:17px 18px 14px !important; box-shadow:0 7px 20px rgba(70,80,140,.07) !important; }
.kpi-title { font-size:17px !important; }
.kpi-value { font-size:34px !important; }
.kpi-sub { font-size:13px !important; }
.st-key-process_card,.st-key-warehouse_card,.st-key-model_card { height:304px !important; min-height:304px !important; max-height:304px !important; border-radius:16px !important; padding:16px 16px 8px !important; box-shadow:0 7px 20px rgba(70,80,140,.07) !important; }
.reference-card-title { font-size:21px !important; color:#0d1735 !important; margin-bottom:7px !important; }
.st-key-process_card .stPlotlyChart,.st-key-warehouse_card .stPlotlyChart,.st-key-model_card .stPlotlyChart,.st-key-trend_card .stPlotlyChart { margin:0 !important; padding:0 !important; }
.st-key-trend_card,.st-key-october_card { height:350px !important; min-height:350px !important; max-height:350px !important; border-radius:16px !important; box-shadow:0 7px 20px rgba(70,80,140,.07) !important; }
.footer { color:#7a849d !important; font-size:12px !important; }
</style>""", unsafe_allow_html=True)
st.markdown("""<div class="reference-nav"><div class="reference-nav-logo">✓</div><div class="reference-nav-item active">⌂</div><div class="reference-nav-item">▥</div><div class="reference-nav-item">▤</div><div class="reference-nav-item">◫</div><div class="reference-nav-item">▣</div><div class="reference-nav-spacer"></div><div class="reference-nav-item">⚙</div></div>""", unsafe_allow_html=True)


def num(v):
    if v is None:
        return 0.0
    try:
        if pd.isna(v):
            return 0.0
    except Exception:
        pass
    try:
        if isinstance(v, str):
            v = v.replace(",", "").replace("฿", "").replace("%", "").strip()
            if not v:
                return 0.0
        return float(v)
    except Exception:
        return 0.0


def money(v):
    return f"฿ {num(v):,.2f}"


def qty(v):
    return f"{num(v):,.2f} EA"


def pct(v):
    return f"{num(v) * 100:.2f}%"


def cell(df, r, c, default=0):
    try:
        return df.iloc[r, c]
    except Exception:
        return default


def excel_date(v):
    try:
        if pd.isna(v):
            return ""
        if isinstance(v, (int, float)):
            return pd.to_datetime(v, unit="D", origin="1899-12-30").strftime("%d/%m/%Y %H:%M")
        return pd.to_datetime(v).strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(v)


@st.cache_data(show_spinner=False)
def load_excel(file_bytes):
    x = pd.ExcelFile(BytesIO(file_bytes), engine="pyxlsb")
    sheets = x.sheet_names
    out = {}
    for s in ["Monitor", "GI Disposal", "Report"]:
        out[s] = (
            pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=s,
                engine="pyxlsb",
                header=None,
            )
            if s in sheets
            else pd.DataFrame()
        )
    return out["Monitor"], out["GI Disposal"], out["Report"], sheets


# ---------- Sidebar ----------
with st.sidebar:
    st.header("📁 Data Source")
    uploaded = st.file_uploader("Upload Disposal Cost Excel", type=["xlsb"])
    st.caption("Required sheets: Monitor, GI Disposal, Report")
    st.divider()
    st.markdown("### 📌 Supported Sheets")
    st.markdown("- Monitor\n- GI Disposal\n- Report")


file_bytes = None
file_name = ""

if uploaded is not None:
    file_bytes = uploaded.getvalue()
    file_name = uploaded.name
else:
    for p in DEFAULT_FILES:
        if p.exists():
            file_bytes = p.read_bytes()
            file_name = p.name
            break


# ---------- Empty state ----------
if file_bytes is None:
    st.markdown(
        """
        <div class="dashboard-header">
            <div class="title-row">
                <div class="title-bar"></div>
                <div class="dashboard-title">📉 NG Disposal Cost Dashboard</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("กรุณา Upload Disposal_cost.xlsb ทางด้านซ้าย")
    st.stop()


# ---------- Load data ----------
try:
    monitor, gi_raw, report, sheets = load_excel(file_bytes)
except Exception as e:
    st.error("❌ ไม่สามารถอ่าน Excel ได้")
    st.code(str(e))
    st.stop()

missing = [s for s in ["Monitor", "GI Disposal", "Report"] if s not in sheets]
if missing:
    st.warning("ไม่พบ Sheet: " + ", ".join(missing))


# ---------- KPI values ----------
total_cost = num(cell(monitor, 5, 2))
total_qty = num(cell(monitor, 5, 10))
ng_rate = num(cell(monitor, 5, 18))
target = num(cell(monitor, 5, 25))

if target <= 0:
    target = TARGET_DEFAULT
elif target > 1:
    target /= 100

if ng_rate < 0:
    ng_rate = 0.0
elif ng_rate > 1:
    ng_rate = ng_rate / 100 if ng_rate <= 100 else 0.0

process_cost = num(cell(monitor, 14, 3))
material_cost = num(cell(monitor, 14, 8))
pmaterial_cost = num(cell(monitor, 14, 13))

period = str(cell(monitor, 1, 26, "")).strip()
if not period or period.lower() == "nan":
    period = "N/A"

last_update = excel_date(cell(monitor, 2, 26, ""))
achieved = ng_rate <= target
difference = abs(target - ng_rate)


# ---------- Warehouse / model data ----------
warehouse_data = pd.DataFrame(columns=["Warehouse", "Cost"])
model_data = pd.DataFrame(columns=["Model", "Cost"])

if not gi_raw.empty and gi_raw.shape[1] >= 27:
    wh = gi_raw.iloc[3:, [22, 23]].copy()
    wh.columns = ["Warehouse", "Cost"]
    wh["Warehouse"] = wh["Warehouse"].fillna("Other").astype(str).str.strip()
    wh["Cost"] = pd.to_numeric(wh["Cost"], errors="coerce").fillna(0)
    wh = wh[
        (wh["Warehouse"] != "")
        & (wh["Warehouse"].str.lower() != "nan")
        & (wh["Cost"] != 0)
    ]
    if not wh.empty:
        warehouse_data = (
            wh.groupby("Warehouse", as_index=False)["Cost"]
            .sum()
            .sort_values("Cost", ascending=True)
        )

    md = gi_raw.iloc[3:, [25, 26]].copy()
    md.columns = ["Model", "Cost"]
    md["Model"] = md["Model"].fillna("Other").astype(str).str.strip()
    md["Cost"] = pd.to_numeric(md["Cost"], errors="coerce").fillna(0)
    md = md[
        (md["Model"] != "")
        & (md["Model"].str.lower() != "nan")
        & (md["Cost"] != 0)
    ]
    if not md.empty:
        model_data = (
            md.groupby("Model", as_index=False)["Cost"]
            .sum()
            .sort_values("Cost", ascending=False)
        )


# ---------- Header ----------
st.markdown(
    f"""
    <div class="dashboard-header">
        <div class="title-row">
            <div class="title-bar"></div>
            <div style="flex:1;">
                <div class="dashboard-title">📉 NG Disposal Cost Dashboard</div>
            </div>
            <div class="header-meta">
                <div><b>📅 Period :</b> {period}</div>
                <div><b>🔄 Last Update :</b> {last_update if last_update else "N/A"}</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------- KPI cards ----------
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Total NG Disposal Cost</div>
            <div class="kpi-value">{money(total_cost)}</div>
            <div class="kpi-sub">Excluding NG Material</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Total NG Disposal Q'ty</div>
            <div class="kpi-value">{qty(total_qty)}</div>
            <div class="kpi-sub">Excluding NG Material</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k3:
    color = "#16803c" if achieved else "#c62828"
    arrow = "▼" if achieved else "▲"
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">NG Rate</div>
            <div class="kpi-value" style="color:{color};">{pct(ng_rate)}</div>
            <div class="kpi-sub" style="color:{color};">
                {arrow} {difference * 100:.2f}% vs Target
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k4:
    cls = "status-ok" if achieved else "status-ng"
    text = "Achieved" if achieved else "Not Achieved"
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Target</div>
            <div class="kpi-value">≤ {pct(target)}</div>
            <div class="kpi-sub">
                Status: <span class="{cls}">{text}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# PROCESS + WAREHOUSE + MODEL
# =========================================================

# เว้นระยะระหว่าง KPI กับ Card แถวกลาง
st.markdown(
    "<div style='height:24px;'></div>",
    unsafe_allow_html=True
)

pcol, wcol, mcol = st.columns(
    [1.15, 1.00, 1.05],
    gap="small"
)


# =========================================================
# PROCESS CARD
# =========================================================

with pcol:

    with st.container(
        border=True,
        key="process_card"
    ):

        st.markdown(
            """
            <div class="reference-card-title">
                NG Cost by Process Type
            </div>
            """,
            unsafe_allow_html=True
        )

        pc1, pc2, pc3 = st.columns(
            3,
            gap="small"
        )

        process_items = [
            (
                pc1,
                "Process",
                process_cost
            ),
            (
                pc2,
                "Material",
                material_cost
            ),
            (
                pc3,
                "P-Material",
                pmaterial_cost
            ),
        ]

        process_max = max(
            process_cost,
            material_cost,
            pmaterial_cost,
            1
        )

        for idx, (
            col,
            name,
            value
        ) in enumerate(process_items):

            with col:

                st.markdown(
                    f"""
                    <div class="reference-process-label">
                        {name}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                remaining = max(
                    process_max - value,
                    0
                )

                fig_process = go.Figure(
                    go.Pie(
                        labels=[
                            "Cost",
                            "Remaining"
                        ],

                        values=[
                            max(value, 0),
                            remaining
                        ],

                        hole=0.68,

                        sort=False,

                        direction="clockwise",

                        textinfo="none",

                        marker=dict(
                            colors=[
                                "#ef233c",
                                "#e8e9ef"
                            ],

                            line=dict(
                                color="#ffffff",
                                width=1
                            )
                        ),

                        hovertemplate=(
                            f"<b>{name}</b><br>"
                            "฿ %{value:,.2f}"
                            "<extra></extra>"
                        )
                    )
                )

                fig_process.update_layout(

                    height=205,

                    margin=dict(
                        l=0,
                        r=0,
                        t=0,
                        b=0
                    ),

                    showlegend=False,

                    paper_bgcolor=(
                        "rgba(0,0,0,0)"
                    ),

                    plot_bgcolor=(
                        "rgba(0,0,0,0)"
                    ),

                    annotations=[
                        dict(
                            text=(
                                f"<b>"
                                f"{value:,.1f}"
                                f"</b>"
                                "<br>"
                                "<span "
                                "style='font-size:12px'>"
                                "THB"
                                "</span>"
                            ),

                            x=0.5,

                            y=0.5,

                            showarrow=False,

                            font=dict(
                                family="Roboto",

                                size=15,

                                color="#111827"
                            )
                        )
                    ]
                )

                st.plotly_chart(
                    fig_process,

                    use_container_width=True,

                    config={
                        "displayModeBar": False
                    },

                    key=f"process_donut_{idx}"
                )


# =========================================================
# WAREHOUSE CARD
# =========================================================

with wcol:

    with st.container(
        border=True,
        key="warehouse_card"
    ):

        st.markdown(
            """
            <div class="reference-card-title">
                Disposal Cost By Warehouse
            </div>
            """,
            unsafe_allow_html=True
        )

        if warehouse_data.empty:

            st.info(
                "No warehouse summary data found."
            )

        else:

            wh = warehouse_data.copy()

            wh["Cost"] = pd.to_numeric(
                wh["Cost"],
                errors="coerce"
            ).fillna(0)

            # -----------------------------------------
            # Keep original warehouse list
            # -----------------------------------------

            warehouse_list = [
                "Work defect Warehouse",
                "QC Warehouse",
                "OS&D Warehouse",
                "Rework Warehouse",
                "R&D Warehouse",
            ]

            wh = (
                wh.set_index("Warehouse")
                .reindex(
                    warehouse_list,
                    fill_value=0
                )
                .reset_index()
            )

            wh["Cost"] = pd.to_numeric(
                wh["Cost"],
                errors="coerce"
            ).fillna(0)

            wh["Warehouse_Display"] = (
                wh["Warehouse"]
                .astype(str)
                .str.replace(
                    "Work defect Warehouse",
                    "Work<br>defect Warehouse",
                    regex=False
                )
                .str.replace(
                    "QC Warehouse",
                    "QC<br>Warehouse",
                    regex=False
                )
                .str.replace(
                    "OS&D Warehouse",
                    "OS&D<br>Warehouse",
                    regex=False
                )
                .str.replace(
                    "Rework Warehouse",
                    "Rework<br>Warehouse",
                    regex=False
                )
                .str.replace(
                    "R&D Warehouse",
                    "R&D<br>Warehouse",
                    regex=False
                )
            )

            max_cost = wh["Cost"].max()

            bar_colors = [
                "#ff0000"
                if value == max_cost
                and max_cost > 0
                else "#b7b7b7"
                for value in wh["Cost"]
            ]

            fig_wh = go.Figure()

            fig_wh.add_trace(
                go.Bar(

                    x=wh["Warehouse_Display"],

                    y=wh["Cost"],

                    marker=dict(
                        color=bar_colors
                    ),

                    text=[
                        f"{value:,.0f}"
                        for value
                        in wh["Cost"]
                    ],

                    textposition="outside",

                    textfont=dict(
                        size=11,
                        color="#555555"
                    ),

                    hovertemplate=(
                        "<b>%{x}</b><br>"
                        "Cost: %{y:,.0f} THB"
                        "<extra></extra>"
                    )
                )
            )

            fig_wh.update_layout(

                height=240,

                margin=dict(
                    l=35,
                    r=10,
                    t=8,
                    b=42
                ),

                paper_bgcolor=(
                    "rgba(0,0,0,0)"
                ),

                plot_bgcolor=(
                    "rgba(0,0,0,0)"
                ),

                showlegend=False,

                font=dict(
                    family="Roboto"
                ),

                xaxis=dict(

                    tickfont=dict(
                        size=12,
                        color="#333333"
                    ),

                    showgrid=False,

                    zeroline=False,

                    automargin=True
                ),

                yaxis=dict(

                    tickfont=dict(
                        size=12,
                        color="#555555"
                    ),

                    tickformat=",",

                    dtick=100000,

                    showgrid=True,

                    gridcolor="#e5e7eb",

                    zeroline=False
                )
            )

            st.plotly_chart(
                fig_wh,

                use_container_width=True,

                config={
                    "displayModeBar": False
                },

                key="warehouse_chart"
            )


# =========================================================
# MODEL CARD
# =========================================================

with mcol:

    with st.container(
        border=True,
        key="model_card"
    ):

        st.markdown(
            """
            <div class="reference-card-title">
                Disposal Cost By Model
            </div>
            """,
            unsafe_allow_html=True
        )

        if model_data.empty:

            st.info(
                "No model summary data found."
            )

        else:

            model_values = (
                model_data
                .copy()
            )

            model_values["Cost"] = (
                pd.to_numeric(
                    model_values["Cost"],
                    errors="coerce"
                )
                .fillna(0)
            )

            fig_model = go.Figure(
                go.Pie(

                    labels=model_values["Model"],

                    values=model_values["Cost"],

                    textinfo="percent",

                    textposition="inside",

                    textfont=dict(
                        family="Roboto",
                        size=18,
                        color="white"
                    ),

                    hole=0,

                    sort=False,

                    marker=dict(
                        line=dict(
                            color="white",
                            width=1
                        )
                    ),

                    hovertemplate=(
                        "<b>%{label}</b><br>"
                        "฿ %{value:,.2f}<br>"
                        "%{percent}"
                        "<extra></extra>"
                    )
                )
            )

            model_legend = [
                f"{name}    {num(value):,.2f}"
                for name, value in zip(
                    model_values["Model"],
                    model_values["Cost"]
                )
            ]

            fig_model.update_traces(
                labels=model_legend,
                domain=dict(
                    x=[0.00, 0.58],
                    y=[0.08, 0.92]
                )
            )

            fig_model.update_layout(
                height=235,
                margin=dict(l=0, r=0, t=0, b=0),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                showlegend=True,
                font=dict(family="Roboto"),
                legend=dict(
                    orientation="v",
                    x=0.62,
                    xanchor="left",
                    y=0.50,
                    yanchor="middle",
                    bgcolor="rgba(255,255,255,0)",
                    borderwidth=0,
                    font=dict(family="Roboto", size=13, color="#111827")
                )
            )

            st.plotly_chart(
                fig_model,
                use_container_width=True,
                config={"displayModeBar": False},
                key="model_chart"
            )


# =========================================================
# TREND + OCTOBER
# =========================================================

st.markdown(
    "<div style='height:16px'></div>",
    unsafe_allow_html=True
)

trend_col, sep_col = st.columns(
    [3.35, 1.00],
    gap="small"
)


# =========================================================
# TREND CARD
# =========================================================

with trend_col:

    with st.container(
        border=True,
        key="trend_card"
    ):

        st.markdown(
            """
            <div class="reference-card-title">
                ↗ Actual Disposal Cost Trend 2026
            </div>
            """,
            unsafe_allow_html=True
        )

        months = [
            "JAN", "FEB", "MAR", "APR",
            "MAY", "JUN", "JUL", "AUG",
            "SEP", "OCT", "NOV", "DEC"
        ]

        actual = (
            pd.to_numeric(
                report.iloc[5, 13:25],
                errors="coerce"
            )
            .fillna(0)
            .tolist()
            if len(report) > 5
            and report.shape[1] >= 25
            else [0] * 12
        )

        target_monthly = (
            pd.to_numeric(
                report.iloc[4, 13:25],
                errors="coerce"
            )
            .fillna(0)
            .tolist()
            if len(report) > 4
            and report.shape[1] >= 25
            else [0] * 12
        )

        trend_fig = go.Figure()


        # Actual
        trend_fig.add_trace(
            go.Scatter(

                x=months,

                y=actual,

                name="Disposal Cost (THB)",

                mode="lines+markers",

                line=dict(
                    color="#4c7df0",
                    width=3
                ),

                marker=dict(
                    color="#4c7df0",
                    size=7
                ),

                fill="tozeroy",

                fillcolor=(
                    "rgba(76,125,240,0.15)"
                ),

                hovertemplate=(
                    "%{x}<br>"
                    "฿ %{y:,.2f}"
                    "<extra></extra>"
                )
            )
        )


        # Target
        trend_fig.add_trace(
            go.Scatter(

                x=months,

                y=target_monthly,

                name="Target Disposal Cost",

                mode="lines+markers",

                line=dict(
                    color="#ef233c",
                    width=2
                ),

                marker=dict(
                    color="#ef233c",
                    size=6
                ),

                fill="tozeroy",

                fillcolor=(
                    "rgba(239,35,60,0.08)"
                ),

                hovertemplate=(
                    "%{x}<br>"
                    "Target ฿ %{y:,.2f}"
                    "<extra></extra>"
                )
            )
        )


        trend_fig.update_layout(
            height=295,
            margin=dict(l=55, r=10, t=8, b=30),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(
                family="Roboto",
                size=14,
                color="#10183d"
            ),
            showlegend=True,
            legend=dict(
                orientation="h",
                x=0.98,
                y=1.08,
                xanchor="right",
                yanchor="bottom",
                font=dict(size=13)
            ),
            xaxis=dict(
                categoryorder="array",
                categoryarray=months,
                showgrid=True,
                gridcolor="#edf0f5",
                zeroline=False,
                tickfont=dict(size=13)
            ),
            yaxis=dict(
                showgrid=True,
                gridcolor="#e6e9f0",
                zeroline=False,
                tickformat=",.0f",
                tickfont=dict(size=12)
            ),
            hovermode="x unified"
        )

        st.plotly_chart(
            trend_fig,

            use_container_width=True,

            config={
                "displayModeBar": False
            },

            key="actual_disposal_trend_2026"
        )

# =========================================================
# OCTOBER DISPOSAL COST
# =========================================================

with sep_col:

    # ---------- Card ----------
    with st.container(
        border=True,
        key="october_card"
    ):

        st.markdown(
            """
            <div class="reference-oct-title">
                October Disposal Cost
            </div>
            """,
            unsafe_allow_html=True
        )

        # ---------- Warehouse Lookup ----------
        wh_lookup = (
            dict(
                zip(
                    warehouse_data["Warehouse"].astype(str),
                    warehouse_data["Cost"]
                )
            )
            if not warehouse_data.empty
            else {}
        )

        # ---------- Total ----------
        sep_total = num(total_cost)

        # ---------- October Items ----------
        sep_items = [
            (
                "Vendor Reject Claim",
                "OS&D Warehouse"
            ),
            (
                "NG Return from Customer",
                "QC Warehouse"
            ),
            (
                "R&D Sample",
                "R&D Warehouse"
            ),
            (
                "NG In process",
                "Work defect Warehouse"
            ),
            (
                "Rework NG",
                "Rework Warehouse"
            )
        ]

        # ---------- Items ----------
        for label, wh_name in sep_items:

            value = num(
                wh_lookup.get(
                    wh_name,
                    0
                )
            )

            ratio = (
                value / sep_total * 100
                if sep_total
                else 0
            )

            ratio_fill = max(
                0,
                min(
                    ratio,
                    100
                )
            )

            value_text = (
                f"฿ {value:,.2f}"
                if value
                else
                "฿ -"
            )

            text_class = (
                "light"
                if ratio >= 60
                else ""
            )

            html = (
                '<div class="sep-item">'

                '<div class="sep-label">'
                f'{label}'
                '</div>'

                '<div class="sep-row">'

                '<div class="sep-value">'
                f'{value_text}'
                '</div>'

                '<div class="sep-pct">'

                '<div '
                'class="sep-pct-fill" '
                f'style="width:calc({ratio_fill:.1f}% * 0.69);">'
                '</div>'

                f'<span class="sep-pct-text {text_class}">'
                f'{ratio:.1f}%'
                '</span>'

                '</div>'

                '</div>'

                '</div>'
            )

            st.markdown(
                html,
                unsafe_allow_html=True
            )

# =========================================================
# DATA CHECK
# =========================================================

with st.expander("🔎 Data Check / ตรวจสอบข้อมูล"):
    check = pd.DataFrame(
        {
            "Item": [
                "Monitor sheet",
                "GI Disposal sheet",
                "Report sheet",
                "Total NG Disposal Cost",
                "Total NG Disposal Qty",
                "NG Rate",
                "Target",
                "Warehouse summary rows",
                "Model summary rows",
            ],
            "Value": [
                "OK" if "Monitor" in sheets else "Missing",
                "OK" if "GI Disposal" in sheets else "Missing",
                "OK" if "Report" in sheets else "Missing",
                money(total_cost),
                qty(total_qty),
                pct(ng_rate),
                pct(target),
                len(warehouse_data),
                len(model_data),
            ],
        }
    )
    st.write("**File:**", file_name)
    st.write("**Sheets found:**", ", ".join(sheets))
    st.dataframe(check, use_container_width=True, hide_index=True)

    if not warehouse_data.empty:
        st.write("**Warehouse summary**")
        st.dataframe(warehouse_data, use_container_width=True, hide_index=True)

    if not model_data.empty:
        st.write("**Model summary**")
        st.dataframe(model_data, use_container_width=True, hide_index=True)


st.markdown(
    '<div class="footer">Source: ERP System | Unit: THB | Prepared by: PI Team</div>',
    unsafe_allow_html=True,
)
