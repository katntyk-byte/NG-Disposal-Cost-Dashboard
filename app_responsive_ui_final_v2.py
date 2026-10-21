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
    initial_sidebar_state="collapsed",
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
    padding-top: 0.6rem !important;
    padding-bottom: 0.8rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
}

    .dashboard-header {
    background:transparent !important;
    border-bottom:1px solid #dfe3ec !important;
    padding-top: 0px !important;
    padding-bottom: 13px !important;
    margin-top: 0 !important;
    margin-bottom: 18px !important;
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

    .header-meta,
.header-meta * {
    font-size: 15px !important;
}
    line-height: 1.7 !important;
    color: #4b5876 !important;
}

.header-meta b {
    font-weight: 700 !important;
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

/* =========================================
   OCTOBER DISPOSAL COST
   ========================================= */

.reference-oct-title {
    background: linear-gradient(180deg, #245b86 0%, #173f63 100%);
    color: #ffffff !important;
    font-size: 16px !important;
    font-weight: 800 !important;
    padding: 7px 10px !important;
    margin: 0 0 10px 0 !important;
    border-radius: 8px 8px 0 0 !important;
    line-height: 1.2 !important;
}

.sep-item {
    width: 100%;
    margin: 0 0 10px 0;
    padding: 0 4px;
    box-sizing: border-box;
}

.sep-info {
    width: 100%;
    margin: 0 0 4px 0;
}

.sep-label {
    font-size: 14px !important;
    font-weight: 700 !important;
    color: #111827 !important;

    white-space: nowrap;
}

.sep-bar {
    position: relative;
    width: calc(100% - 42px);
    height: 25px;
    background: #edf0f3;
    overflow: hidden;
    display: inline-block;
    vertical-align: middle;
    box-sizing: border-box;
}

.sep-fill {
    position: absolute;
    left: 0;
    top: 0;
    height: 100%;
    background: linear-gradient(90deg, #ffb3b8 0%, #ff6b75 45%, #e51f32 100%);
    transition: width 0.3s ease;
}

.sep-value {
    position: absolute;
    left: 50%;
    top: 50%;
    transform: translate(-50%, -50%);
    width: 100%;
    text-align: center;
    font-size: 14px;
    font-weight: 800 !important;
    color: #111827 !important;
    white-space: nowrap;
    z-index: 5;
    text-shadow: 0 1px 2px rgba(0, 0, 0, 0.18);
}

.sep-percent {
    display: inline-flex;
    width: 36px;
    height: 25px;
    margin-left: 6px;
    align-items: center;
    justify-content: flex-end;
    vertical-align: middle;
    box-sizing: border-box;
    font-size: 15px !important;
    font-weight: 800 !important;
    color: #111827 !important;
}

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
    height: 270px !important;
    min-height: 270px !important;
    max-height: 270px !important;

    border-radius: 16px !important;
    padding: 10px 16px !important;

    box-shadow: 0 7px 20px rgba(70,80,140,.07) !important;

    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
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

        font-size: 16px;

        line-height: 1.2;

        font-weight: 700;

        color: #10183d;

        margin: 0 0 8px 0;

        padding: 0;
    }


    /* ---------- PROCESS LABEL ---------- */

    .reference-process-label {
    font-size: 15px !important;
    font-weight: 700 !important;
    color: #c40000 !important;
    text-align: center !important;

    margin-top: 8px !important;
    margin-bottom: 0 !important;
    line-height: 1.2 !important;
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

        background: linear-gradient(180deg, #245b86 0%, #173f63 100%) !important;

        color: #ffffff !important;

        -webkit-text-fill-color: #ffffff !important;

        font-size: 21px !important;

        font-weight: 700 !important;

        line-height: 1.2 !important;

        padding: 7px 10px !important;

        margin: 0 0 10px 0 !important;

        border-radius: 8px 8px 0 0 !important;
    }

    .reference-oct-title * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
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
    height:90px !important;
    min-height:90px !important;
    max-height:90px !important;

    padding:8px 12px !important;
    box-sizing:border-box !important;

    display:flex !important;
    flex-direction:column !important;
    justify-content:center !important;
}

    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("""<style>
/* =========================
   LEFT SIDEBAR
   ========================= */

.stApp { background:#EEF0F3 !important; }

section[data-testid="stSidebar"] {
    background:#ffffff !important;
    border-right:1px solid #e8ebf3 !important;
    min-width:230px !important;
    max-width:230px !important;
}

section[data-testid="stSidebar"] > div:first-child {
    padding-top:24px !important;
    padding-left:18px !important;
    padding-right:18px !important;
}

section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
    padding-top:0 !important;
}

section[data-testid="stSidebar"] h2 {
    color:#17213d !important;
    font-size:18px !important;
    font-weight:800 !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
    background:#f7f8fc !important;
    border:1px dashed #cbd3ea !important;
    border-radius:12px !important;
    padding:12px !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background:transparent !important;
    border:none !important;
}

section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
    border-radius:9px !important;
    border:1px solid #d9dff0 !important;
    background:#ffffff !important;
    color:#3559e6 !important;
    font-weight:600 !important;
}

section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
    color:#8a93aa !important;
    font-size:11px !important;
    line-height:1.6 !important;
}

/* =========================
   SIDEBAR COLLAPSE BUTTON
   Keep the real Streamlit button clickable,
   but replace the Material Icon text with »
   ========================= */

button[data-testid="stSidebarCollapseButton"] {
    position: relative !important;
    font-size: 0 !important;
    color: transparent !important;
}

button[data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"],
button[data-testid="stSidebarCollapseButton"] .material-symbols-rounded,
button[data-testid="stSidebarCollapseButton"] .material-icons {
    font-size: 0 !important;
    color: transparent !important;
}

button[data-testid="stSidebarCollapseButton"] [data-testid="stIconMaterial"]::before,
button[data-testid="stSidebarCollapseButton"] .material-symbols-rounded::before,
button[data-testid="stSidebarCollapseButton"] .material-icons::before {
    content: "»" !important;
    display: inline-block !important;
    font-family: Arial, sans-serif !important;
    font-size: 26px !important;
    font-weight: 400 !important;
    line-height: 1 !important;
    color: #4b5876 !important;
    pointer-events: none !important;
}

/* Make sure the replacement icon never blocks the real button click */
button[data-testid="stSidebarCollapseButton"]::before,
button[data-testid="stSidebarCollapseButton"]::after {
    pointer-events: none !important;
}

/* Keep dashboard content at normal position */
[data-testid="stAppViewContainer"] {
    margin-left:0 !important;
}

.dashboard-header {
    background:transparent !important;
    border-bottom:1px solid #dfe3ec !important;
    padding-bottom:13px !important;
    margin-bottom:18px !important;
}

.dashboard-title {
    font-size:31px !important;
    font-weight:800 !important;
    color:#0d1735 !important;
}

.header-meta {
    font-size:12px !important;
    line-height:1.65 !important;
    color:#4b5876 !important;
}

.st-key-k1,.st-key-k2,.st-key-k3,.st-key-k4 {
    height:90px !important;
    min-height:90px !important;
    max-height:90px !important;
}

.kpi-card {
    min-height: 148px !important;
    height: 148px !important;

    border-radius: 16px !important;
    padding: 14px 18px !important;

    box-shadow: 0 7px 20px rgba(70,80,140,.07) !important;

    display: flex !important;
    flex-direction: column !important;

    align-items: center !important;
    justify-content: center !important;

    text-align: center !important;

    box-sizing: border-box !important;
}

/* เลื่อนข้อความทั้งชุดลงเล็กน้อย */
.kpi-card > div {
    transform: translateY(7px);
}

/* หัวข้อ */
.kpi-card .kpi-title {
    text-align: center !important;
}

/* ตัวเลข */
.kpi-card .kpi-value {
    text-align: center !important;
}

/* ข้อความด้านล่าง */
.kpi-card .kpi-sub {
    text-align: center !important;
}

.kpi-title { font-size:20px !important; }
.kpi-value { font-size:34px !important; }
.kpi-sub {
    font-size: 15px !important;
    font-weight: 500 !important;
}

.st-key-process_card,.st-key-warehouse_card,.st-key-model_card {
    height:270px !important;
    min-height:270px !important;
    max-height:270px !important;
    border-radius:16px !important;
    padding:16px 16px 8px !important;
    box-shadow:0 7px 20px rgba(70,80,140,.07) !important;
}

/* ---------- CENTER CONTENT IN MIDDLE CARDS ---------- */

.st-key-process_card,
.st-key-warehouse_card,
.st-key-model_card {
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
}

.reference-card-title {
    font-size:21px !important;
    color:#0d1735 !important;
    margin-bottom:7px !important;
}

.st-key-process_card .stPlotlyChart,
.st-key-warehouse_card .stPlotlyChart,
.st-key-model_card .stPlotlyChart,
.st-key-trend_card .stPlotlyChart {
    margin:0 !important;
    padding:0 !important;
}

/* =========================================================
   BOTTOM CARDS
   ========================================================= */

.st-key-trend_card,
.st-key-october_card {
    height:330px !important;
    min-height:330px !important;
    max-height:330px !important;
    padding:8px 16px 4px !important;
    box-sizing:border-box !important;
    border-radius:16px !important;
    box-shadow:0 7px 20px rgba(70,80,140,.07) !important;
}

/* ลดพื้นที่กราฟด้านใน */
.st-key-trend_card .stPlotlyChart {
    height:225px !important;
    max-height:225px !important;
    margin-top:-10px !important;
}

/* ---------- OCTOBER DISPOSAL COST ---------- */

.st-key-october_card {
    height:330px !important;
    min-height:330px !important;
    max-height:330px !important;
    overflow:hidden !important;
    padding:8px 12px 4px !important;
    box-sizing:border-box !important;
}

/* ลดระยะห่างของข้อมูลแต่ละรายการ */
.st-key-october_card p {
    margin-top:2px !important;
    margin-bottom:2px !important;
    line-height:1.15 !important;
}

/* ลดความสูงของ Progress Bar */
.st-key-october_card [data-testid="stProgress"] {
    margin-top:1px !important;
    margin-bottom:2px !important;
}

/* ---------- REDUCE OCTOBER ROW SPACING ---------- */

.st-key-october_card [data-testid="stVerticalBlock"] {
    gap: 0px !important;
}

.st-key-october_card p {
    margin-top: 0px !important;
    margin-bottom: 1px !important;
    line-height: 1.0 !important;
}

.st-key-october_card [data-testid="stProgress"] {
    margin-top: 0px !important;
    margin-bottom: 1px !important;
}

.footer { color:#7a849d !important; font-size:14px !important; }
</style>""", unsafe_allow_html=True)



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
with st.container(key="kpi_row"):
    k1, k2, k3, k4 = st.columns(4, gap="small")

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
    "<div style='height:16px;'></div>",
    unsafe_allow_html=True
)

with st.container(key="middle_row"):
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

                        height=190,

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
                            "displayModeBar": False,
                            "responsive": True
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

                    height=245,

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
                        "displayModeBar": False,
                        "responsive": True
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
                    height=245,
                    margin=dict(l=0, r=0, t=0, b=0),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    showlegend=True,
                    font=dict(family="Roboto"),
                    legend=dict(
                        orientation="v",
                        x=0.54,
                        xanchor="left",
                        y=0.50,
                        yanchor="middle",
                        bgcolor="rgba(255,255,255,0)",
                        borderwidth=0,
                        font=dict(family="Roboto", size=16, color="#111827")
                    )
                )

                st.plotly_chart(
                    fig_model,
                    use_container_width=True,
                    config={
                    "displayModeBar": False,
                    "responsive": True
                },
                    key="model_chart"
                )
# =========================================================
# TREND + OCTOBER
# =========================================================

with st.container(key="bottom_row"):
    trend_col, sep_col = st.columns(
        [3.5, 1.0],
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
                    Actual Disposal Cost Trend 2026
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
                height=300,
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
                    font=dict(size=15)
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
                    "displayModeBar": False,
                    "responsive": True
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
                    f"{value:,.0f} THB"
                    if value
                    else
                    "0 THB"
                )

                text_class = ""

                html = (
        '<div class="sep-item">'

        '<div class="sep-info">'
            f'<div class="sep-label">{label}</div>'
        '</div>'

        '<div class="sep-bar">'

            f'<div class="sep-fill" style="width:{ratio_fill:.1f}%"></div>'

            f'<div class="sep-value">{value_text}</div>'

        '</div>'

        f'<div class="sep-percent">{ratio:.0f}%</div>'

        '</div>'
    )

                st.markdown(
                    html,
                    unsafe_allow_html=True
                )

    st.markdown(
        '<div class="footer">Source: ERP System | Unit: THB | Prepared by: PI Team</div>',
        unsafe_allow_html=True,
    )



# =========================================================
# RESPONSIVE UI LAYOUT
# Component-level responsive behavior only.
# No page zoom / fixed canvas scaling.
# =========================================================
st.markdown(r"""
<style>

/* ---------- Fluid page shell ---------- */
[data-testid="stMainBlockContainer"],
section.main > div.block-container,
.block-container {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 1600px !important;
    margin: 0 auto !important;
    box-sizing: border-box !important;
    padding-left: clamp(10px, 1.4vw, 24px) !important;
    padding-right: clamp(10px, 1.4vw, 24px) !important;
}

[data-testid="stMain"],
[data-testid="stAppViewContainer"] {
    overflow-x: hidden !important;
}

/* ---------- All Streamlit rows may wrap ---------- */
div[data-testid="stHorizontalBlock"] {
    width: 100% !important;
    max-width: 100% !important;
    flex-wrap: wrap !important;
    align-items: stretch !important;
    box-sizing: border-box !important;
    gap: clamp(8px, 1vw, 14px) !important;
}

div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
    min-width: 0 !important;
    box-sizing: border-box !important;
    flex-shrink: 1 !important;
}

/* ---------- Header ---------- */
.dashboard-header {
    width: 100% !important;
    box-sizing: border-box !important;
}

.title-row {
    width: 100% !important;
    min-width: 0 !important;
    display: flex !important;
    align-items: center !important;
    gap: clamp(8px, 1vw, 16px) !important;
    flex-wrap: wrap !important;
}

.dashboard-title {
    min-width: 0 !important;
    font-size: clamp(20px, 2.1vw, 32px) !important;
    line-height: 1.1 !important;
    overflow-wrap: anywhere !important;
}

.header-meta {
    min-width: 0 !important;
    max-width: 100% !important;
    font-size: clamp(9px, .9vw, 14px) !important;
    line-height: 1.25 !important;
    overflow-wrap: anywhere !important;
}

/* ---------- KPI cards ---------- */
.st-key-kpi_row {
    width: 100% !important;
    max-width: 100% !important;
}

.kpi-card {
    width: 100% !important;
    min-height: 90px !important;
    height: auto !important;
    box-sizing: border-box !important;
    padding: clamp(10px, 1vw, 16px) !important;
}

.kpi-title {
    font-size: clamp(13px, 1.25vw, 20px) !important;
}

.kpi-value {
    font-size: clamp(21px, 2.25vw, 36px) !important;
    line-height: 1.05 !important;
}

.kpi-sub {
    font-size: clamp(10px, .95vw, 15px) !important;
}

/* ---------- Middle cards ---------- */
.st-key-process_card,
.st-key-warehouse_card,
.st-key-model_card {
    width: 100% !important;
    min-width: 0 !important;
    min-height: 270px !important;
    height: 270px !important;
    max-height: 270px !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
}

.st-key-process_card .stPlotlyChart,
.st-key-warehouse_card .stPlotlyChart,
.st-key-model_card .stPlotlyChart,
.st-key-trend_card .stPlotlyChart {
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    margin: 0 !important;
    padding: 0 !important;
}

/* Process donut columns can wrap only when the card itself becomes narrow. */
.st-key-process_card div[data-testid="stHorizontalBlock"] {
    width: 100% !important;
    flex-wrap: nowrap !important;
}

.st-key-process_card div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
    min-width: 0 !important;
}

/* ---------- Bottom cards ---------- */
.st-key-trend_card,
.st-key-october_card {
    width: 100% !important;
    min-width: 0 !important;
    min-height: 350px !important;
    height: 350px !important;
    max-height: 350px !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
}

.reference-card-title {
    font-size: clamp(14px, 1.15vw, 18px) !important;
    line-height: 1.2 !important;
    overflow-wrap: anywhere !important;
}

/* ---------- October card ---------- */
.reference-oct-item {
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
}

.reference-oct-row {
    width: 100% !important;
    min-width: 0 !important;
    display: flex !important;
    align-items: center !important;
    gap: 6px !important;
}

.reference-oct-value {
    width: clamp(72px, 9vw, 125px) !important;
    min-width: 0 !important;
    flex: 0 1 auto !important;
    font-size: clamp(10px, .9vw, 15px) !important;
}

.reference-oct-bar {
    min-width: 0 !important;
    flex: 1 1 auto !important;
}

.reference-oct-pct {
    width: 32px !important;
    min-width: 32px !important;
    flex: 0 0 32px !important;
    font-size: clamp(10px, .85vw, 14px) !important;
}

/* =========================================================
   TABLET
   ========================================================= */
@media (max-width: 1100px) {
    /* KPI: 2 x 2 */
    .st-key-kpi_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 calc(50% - 7px) !important;
        width: calc(50% - 7px) !important;
    }

    /* Middle cards: stack safely instead of squeezing charts. */
    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        min-height: 300px !important;
        height: 300px !important;
        max-height: 300px !important;
    }

    /* Bottom cards: stack safely. */
    .st-key-bottom_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
    }

    .st-key-trend_card,
    .st-key-october_card {
        min-height: 330px !important;
        height: 330px !important;
        max-height: 330px !important;
    }
}

/* =========================================================
   PHONE
   ========================================================= */
@media (max-width: 650px) {
    [data-testid="stMainBlockContainer"],
    section.main > div.block-container,
    .block-container {
        padding-left: 8px !important;
        padding-right: 8px !important;
    }

    .dashboard-title {
        font-size: clamp(18px, 6vw, 24px) !important;
    }

    .header-meta {
        width: 100% !important;
        margin-left: 0 !important;
    }

    /* KPI: one card per row */
    .st-key-kpi_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
    }

    /* One donut per row: prevents labels/plots from touching. */
    .st-key-process_card div[data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
    }

    .st-key-process_card div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        min-height: 290px !important;
        height: auto !important;
        max-height: none !important;
    }

    .st-key-trend_card,
    .st-key-october_card {
        min-height: 300px !important;
        height: auto !important;
        max-height: none !important;
    }

    .reference-card-title {
        font-size: 16px !important;
    }

    .reference-oct-value {
        width: 80px !important;
        font-size: 10px !important;
    }

    .reference-oct-pct {
        width: 28px !important;
        min-width: 28px !important;
        flex-basis: 28px !important;
        font-size: 10px !important;
    }
}

/* Very narrow phone */
@media (max-width: 420px) {
    .kpi-card {
        min-height: 88px !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        min-height: 270px !important;
    }
}

</style>
""", unsafe_allow_html=True)



st.markdown("""
<style>
/* =========================================================
   FINAL UI OVERRIDES
   Component responsive only - NO page zoom / NO fixed canvas
   ========================================================= */

[data-testid="stMainBlockContainer"],
section.main > div.block-container,
.block-container {
    width: 100% !important;
    max-width: 1600px !important;
    min-width: 0 !important;
    margin: 0 auto !important;
    box-sizing: border-box !important;
}

[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    overflow-x: hidden !important;
}

@media (min-width: 1201px) {
    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],
    .st-key-bottom_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        min-width: 0 !important;
    }
}

@media (min-width: 901px) and (max-width: 1200px) {
    .st-key-kpi_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 calc(50% - 7px) !important;
        width: calc(50% - 7px) !important;
        min-width: 0 !important;
    }

    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        min-width: 0 !important;
    }

    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(1) {
        flex: 1 1 100% !important;
        width: 100% !important;
    }

    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2),
    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(3) {
        flex: 1 1 calc(50% - 7px) !important;
        width: calc(50% - 7px) !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        height: 285px !important;
        min-height: 285px !important;
        max-height: 285px !important;
    }

    .st-key-bottom_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(1) {
        flex: 1 1 calc(70% - 7px) !important;
        width: calc(70% - 7px) !important;
    }

    .st-key-bottom_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2) {
        flex: 1 1 calc(30% - 7px) !important;
        width: calc(30% - 7px) !important;
    }
}

@media (min-width: 651px) and (max-width: 900px) {
    .st-key-kpi_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 calc(50% - 7px) !important;
        width: calc(50% - 7px) !important;
        min-width: 0 !important;
    }

    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],
    .st-key-bottom_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
        min-width: 0 !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        height: 300px !important;
        min-height: 300px !important;
        max-height: 300px !important;
    }

    .st-key-trend_card,
    .st-key-october_card {
        height: 330px !important;
        min-height: 330px !important;
        max-height: 330px !important;
    }
}

@media (max-width: 650px) {
    [data-testid="stMainBlockContainer"],
    section.main > div.block-container,
    .block-container {
        padding-left: 8px !important;
        padding-right: 8px !important;
        padding-top: 12px !important;
    }

    .title-row {
        align-items: flex-start !important;
    }

    .dashboard-title {
        font-size: clamp(18px, 6vw, 24px) !important;
    }

    .header-meta {
        width: 100% !important;
        flex: 1 1 100% !important;
        font-size: 10px !important;
    }

    .st-key-kpi_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],
    .st-key-middle_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],
    .st-key-bottom_row > div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
        min-width: 0 !important;
    }

    .kpi-card {
        min-height: 92px !important;
        height: auto !important;
        padding: 10px 12px !important;
    }

    .kpi-title {
        font-size: 14px !important;
    }

    .kpi-value {
        font-size: clamp(22px, 7vw, 32px) !important;
    }

    .kpi-sub {
        font-size: 11px !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        height: auto !important;
        min-height: 280px !important;
        max-height: none !important;
    }

    .st-key-process_card div[data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
    }

    .st-key-process_card div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        flex: 1 1 100% !important;
        width: 100% !important;
        min-width: 0 !important;
    }

    .st-key-trend_card,
    .st-key-october_card {
        height: auto !important;
        min-height: 300px !important;
        max-height: none !important;
    }

    .st-key-trend_card .stPlotlyChart {
        height: auto !important;
        max-height: none !important;
    }

    .reference-oct-value {
        width: 82px !important;
        font-size: 10px !important;
    }

    .reference-oct-pct {
        width: 28px !important;
        min-width: 28px !important;
        flex-basis: 28px !important;
        font-size: 10px !important;
    }
}

@media (max-width: 420px) {
    .reference-card-title {
        font-size: 15px !important;
    }

    .st-key-process_card,
    .st-key-warehouse_card,
    .st-key-model_card {
        min-height: 260px !important;
    }
}
</style>
""", unsafe_allow_html=True)
