"""Shared responsive page chrome for the hosted Streamlit website."""
from __future__ import annotations

import base64
import functools
import html
from contextlib import contextmanager
from pathlib import Path

import streamlit as st
import streamlit.components.v2 as components_v2

ROOT = Path(__file__).resolve().parent
APP_NAME = "Matalan Carton Supply Report"
PACD_LOGO = str(ROOT / "assets" / "logo_left.png")
MATALAN_LOGO = str(ROOT / "assets" / "logo_right.png")
FAVICON = str(ROOT / "assets" / "favicon.png")


_DASHBOARD_CARD_COMPONENT = components_v2.component(
    "dashboard_metric_cards",
    html="<div class='dashboard-metrics' role='group' aria-label='Dashboard view'></div>",
    css="""
      .dashboard-metrics { display:grid; grid-template-columns:repeat(var(--metric-columns,4),minmax(0,1fr));
        gap:.625rem; width:100%; padding:3px 0 .625rem; box-sizing:border-box; }
      .metric { min-width:0; min-height:112px; display:flex; flex-direction:column;
        justify-content:space-between; padding:.85rem .95rem; overflow:hidden; position:relative;
        border:1px solid #DDE4EE; border-radius:14px; background:#FFFFFF; color:#10294F;
        box-shadow:0 4px 18px rgba(8,26,54,.055); text-align:left; cursor:pointer;
        font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
        transition:transform .16s ease,border-color .16s ease,box-shadow .16s ease; }
      .metric::after { content:""; position:absolute; width:72px; height:72px; right:-25px;
        bottom:-32px; border-radius:50%; background:rgba(20,47,87,.06); }
      .metric:hover { border-color:#B9C5D5; box-shadow:0 8px 22px rgba(8,26,54,.11);
        transform:translateY(-2px); }
      .metric.active { border-color:#E22B54;
        box-shadow:0 0 0 3px rgba(226,43,84,.22),0 8px 22px rgba(8,26,54,.12); }
      .metric.navy { color:#FFFFFF; border-color:#142F57;
        background:linear-gradient(138deg,#091D3C,#173B6D); }
      .metric.navy-mid { color:#FFFFFF; border-color:#244B7A;
        background:linear-gradient(138deg,#17345E,#315D91); }
      .metric.navy-soft { color:#FFFFFF; border-color:#3A5B86;
        background:linear-gradient(138deg,#294A75,#5579A4); }
      .metric.red { color:#FFFFFF; border-color:#A81737;
        background:linear-gradient(138deg,#7C102A,#C32148); }
      .metric.red-mid { color:#FFFFFF; border-color:#8F2440;
        background:linear-gradient(138deg,#6D1A36,#A93654); }
      .metric.red-soft { color:#FFFFFF; border-color:#A84961;
        background:linear-gradient(138deg,#8D3850,#C05B73); }
      .metric.matte-navy-deep { color:#FFFFFF; border-color:#112E55; background:#112E55; }
      .metric.matte-crimson-deep { color:#FFFFFF; border-color:#7D1936; background:#7D1936; }
      .metric.matte-navy-mid { color:#FFFFFF; border-color:#274D78; background:#274D78; }
      .metric.matte-crimson-mid { color:#FFFFFF; border-color:#942E4B; background:#942E4B; }
      .metric.matte-navy-soft { color:#FFFFFF; border-color:#42658E; background:#42658E; }
      .metric.matte-crimson-soft { color:#FFFFFF; border-color:#A84662; background:#A84662; }
      .metric.navy::after,.metric.navy-mid::after,.metric.navy-soft::after,
      .metric.red::after,.metric.red-mid::after,.metric.red-soft::after,
      .metric[class*="matte-"]::after {
        background:rgba(255,255,255,.09); }
      .label { position:relative; z-index:1; color:#65758B; font-size:.72rem; line-height:1.2;
        font-weight:700; letter-spacing:.055em; text-transform:uppercase; }
      .value { position:relative; z-index:1; color:#10294F; font-size:1.85rem; line-height:1;
        font-weight:750; letter-spacing:-.04em; font-variant-numeric:tabular-nums; }
      .note { position:relative; z-index:1; color:#78889D; font-size:.72rem; line-height:1.2; }
      .metric.navy .label,.metric.navy-mid .label,.metric.navy-soft .label,
      .metric.red .label,.metric.red-mid .label,.metric.red-soft .label {
        color:rgba(255,255,255,.76); }
      .metric[class*="matte-"] .label { color:rgba(255,255,255,.78); }
      .metric.navy .value,.metric.navy-mid .value,.metric.navy-soft .value,
      .metric.red .value,.metric.red-mid .value,.metric.red-soft .value { color:#FFFFFF; }
      .metric[class*="matte-"] .value { color:#FFFFFF; }
      .metric.navy .note,.metric.navy-mid .note,.metric.navy-soft .note,
      .metric.red .note,.metric.red-mid .note,.metric.red-soft .note {
        color:rgba(255,255,255,.72); }
      .metric[class*="matte-"] .note { color:rgba(255,255,255,.74); }
      .metric.selected-now { animation:metric-select .22s cubic-bezier(.22,1,.36,1); }
      @keyframes metric-select { 50% { transform:translateY(-2px) scale(.985); } }
      @media (max-width:1100px) { .dashboard-metrics { grid-template-columns:repeat(2,minmax(0,1fr)); } }
      @media (max-width:640px) { .dashboard-metrics { gap:.5rem; }
        .metric { min-height:102px; padding:.75rem; }
        .value { font-size:1.5rem; } }
      @media (prefers-reduced-motion:reduce) { .metric { transition:none; }
        .metric.selected-now { animation:none; } }
    """,
    js="""
      export default function(component) {
        const { data, parentElement, setStateValue } = component;
        const root = parentElement.querySelector('.dashboard-metrics');
        root.style.setProperty('--metric-columns', String(data.columns || 4));

        const wantedIds = new Set(data.tiles.map((tile) => String(tile.id)));
        root.querySelectorAll('.metric').forEach((card) => {
          if (!wantedIds.has(card.dataset.id)) card.remove();
        });

        data.tiles.forEach((tile) => {
          let card = Array.from(root.querySelectorAll('.metric')).find(
            (item) => item.dataset.id === String(tile.id)
          );
          if (!card) {
            card = document.createElement('button');
            card.type = 'button';
            card.dataset.id = String(tile.id);
            card.append(
              Object.assign(document.createElement('span'), { className: 'label' }),
              Object.assign(document.createElement('span'), { className: 'value' }),
              Object.assign(document.createElement('span'), { className: 'note' }),
            );
            root.appendChild(card);
          }
          card.className = `metric ${tile.tone || ''} ${tile.id === data.active ? 'active' : ''}`;
          card.setAttribute('aria-pressed', tile.id === data.active ? 'true' : 'false');
          card.querySelector('.label').textContent = tile.label;
          card.querySelector('.value').textContent = tile.value;
          card.querySelector('.note').textContent = tile.note || '\u00a0';

          card.onclick = () => {
            root.querySelectorAll('.metric').forEach((item) => {
              const selected = item === card;
              item.classList.toggle('active', selected);
              item.setAttribute('aria-pressed', selected ? 'true' : 'false');
            });
            card.classList.remove('selected-now');
            void card.offsetWidth;
            card.classList.add('selected-now');
            setStateValue('selected', tile.id);
          };
        });
      }
    """,
)


@functools.lru_cache(maxsize=4)
def _data_uri(path: str) -> str:
    raw = Path(path).read_bytes()
    return "data:image/png;base64," + base64.b64encode(raw).decode("ascii")


THEME = """
:root {
  --mc-navy: #002B61; --mc-sky: #00AEEF; --mc-red: #B61922;
  --mc-ink: #0F1B2D; --mc-muted: #64748B; --mc-line: #E2E8F0;
  --mc-surface: #FFFFFF; --mc-canvas: #F5F7FA; --mc-warn: #B45309;
  --mc-warn-bg: #FFF7E6; --mc-ok: #15803D; --mc-radius: 12px;
  --mc-shadow: 0 1px 2px rgba(15,27,45,.04), 0 1px 3px rgba(15,27,45,.06);
  --mc-page-top: 1rem; --mc-page-inline: 1.5rem; --mc-page-bottom: 1.5rem;
}
.stApp { background: var(--mc-canvas); }
[data-testid="stHeader"] { height: 68px; background: rgba(245,247,250,.96);
  border-bottom: 1px solid var(--mc-line); backdrop-filter: blur(10px); }
[data-testid="stHeader"]::before {
  content: "CARTON SUPPLY REPORT"; position: absolute; top: 0; right: 114px;
  display: flex; align-items: center; height: 68px; padding-right: 14px;
  border-right: 1px solid var(--mc-line); color: var(--mc-navy);
  font-size: .75rem; font-weight: 600; letter-spacing: .08em; white-space: nowrap;
}
[data-testid="stHeader"]::after {
  content: ""; position: absolute; top: 14px; right: 20px; width: 74px; height: 40px;
  background: url("__MATALAN_LOGO__") center / contain no-repeat;
}
[data-testid="stDecoration"], [data-testid="stHeaderActionElements"] { display: none; }
[data-testid="stMainBlockContainer"] { width: 100%; max-width: 1560px;
  padding: var(--mc-page-top) var(--mc-page-inline) var(--mc-page-bottom); }
[data-testid="stSidebar"] { box-shadow: none; }
[data-testid="stSidebarHeader"] { padding-top: 1.25rem; padding-bottom: .5rem; }
[data-testid="stSidebarHeader"] img, [data-testid="stHeaderLogo"] img,
[data-testid="stLogo"] { height: 24px !important; width: auto !important; }
[data-testid="stSidebarNav"] [data-testid="stNavSectionHeader"] {
  font-size: .6875rem; font-weight: 600; letter-spacing: .08em;
  text-transform: uppercase; color: var(--mc-muted); }
[data-testid="stSidebarNavLink"] { border-radius: 8px; }
[data-testid="stSidebarNavLink"] span { font-weight: 500; }
h1, h2, h3, h4 { color: var(--mc-navy); letter-spacing: -.015em; }
[data-testid="stCaptionContainer"] { color: var(--mc-muted); }
.mc-topbrand-spacer { display: block; width: 100%; height: 68px; flex: 0 0 68px; }
.mc-header { display: flex; flex-wrap: wrap; align-items: flex-end; justify-content: space-between;
  gap: .625rem 2rem; padding-bottom: .75rem; margin-bottom: .5rem;
  border-bottom: 1px solid var(--mc-line); }
.mc-header__main { min-width: 0; flex: 1 1 22rem; }
.mc-eyebrow { margin-bottom: .4rem; font-size: .6875rem; font-weight: 600;
  letter-spacing: .08em; text-transform: uppercase; color: var(--mc-muted); }
.mc-title { margin: 0; font-size: 1.75rem; line-height: 1.2; font-weight: 600;
  letter-spacing: -.02em; color: var(--mc-navy); }
.mc-desc { margin: .25rem 0 0; max-width: 50rem; font-size: .9rem;
  line-height: 1.5; color: var(--mc-muted); }
.mc-status { display: flex; flex-direction: column; align-items: flex-end; gap: .125rem;
  padding: .5rem .875rem; background: var(--mc-surface); border: 1px solid var(--mc-line);
  border-radius: 10px; box-shadow: var(--mc-shadow); }
.mc-status__label { font-size: .6875rem; font-weight: 600; letter-spacing: .06em;
  text-transform: uppercase; color: var(--mc-muted); }
.mc-status__value { font-size: .9375rem; font-weight: 600; color: var(--mc-navy); white-space: nowrap; }
.mc-status__value::before { content: ""; display: inline-block; width: 7px; height: 7px;
  margin-right: .4rem; border-radius: 50%; background: var(--mc-sky); }
.mc-kpis { display: grid; grid-template-columns: repeat(var(--mc-kpi-count, 4),minmax(0,1fr));
  gap: .625rem; margin: .125rem 0 .625rem; }
.mc-kpi { min-width: 0; min-height: 82px; padding: .75rem .875rem; background: var(--mc-surface);
  border: 1px solid var(--mc-line); border-radius: var(--mc-radius); box-shadow: var(--mc-shadow); }
.mc-kpi__label { font-size: .8125rem; font-weight: 500; color: var(--mc-muted); }
.mc-kpi__value { margin-top: .175rem; font-size: 1.4rem; line-height: 1.2; font-weight: 600;
  letter-spacing: -.02em; font-variant-numeric: tabular-nums; color: var(--mc-navy); }
.mc-kpi__note { margin-top: .2rem; font-size: .72rem; color: var(--mc-muted); }
.mc-kpi--warn { border-color: #F5D9A8; background: var(--mc-warn-bg); }
.mc-kpi--warn .mc-kpi__value { color: var(--mc-warn); }
.mc-kpi--ok .mc-kpi__value { color: var(--mc-ok); }
[class*="st-key-card_"] { background: var(--mc-surface); border-radius: var(--mc-radius) !important;
  box-shadow: var(--mc-shadow); }
[class*="st-key-card_"] h3 { margin: 0 0 .25rem; font-size: 1.05rem; line-height: 1.25; }
[class*="st-key-card_"] [data-testid="stVerticalBlock"] { gap: .45rem; }
[class*="st-key-card_"] [data-testid="stVegaLiteChart"] { margin-top: -.35rem; }
[class*="st-key-kpi_card_"] button {
  width: 100%; min-height: 88px; padding: .7rem .8rem; white-space: pre-line;
  justify-content: flex-start; text-align: left; line-height: 1.45;
  background: var(--mc-surface); border-color: var(--mc-line); color: var(--mc-navy);
  box-shadow: var(--mc-shadow); font-size: .86rem;
}
[class*="st-key-kpi_card_"] button p::first-line { color: var(--mc-muted); }
[class*="st-key-kpi_card_"] button[kind="primary"] {
  border-color: var(--mc-sky); background: #EAF7FD; box-shadow: 0 0 0 1px var(--mc-sky);
}
[class*="st-key-manage_row_"] { padding: .55rem 0; border-bottom: 1px solid var(--mc-line); }
[data-testid="stMultiSelect"] [data-baseweb="tag"] {
  background: #EAF6FD; border: 1px solid #B9DFF3; border-radius: 6px;
  color: var(--mc-navy); font-weight: 500; }
[data-testid="stMultiSelect"] [data-baseweb="tag"] svg { color: #475569; }
[data-testid="stExpander"] details { background: var(--mc-surface); border-radius: var(--mc-radius); }
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
  font-weight: 500; min-height: 2.5rem; }
[data-testid="stFileUploaderDropzone"] { background: var(--mc-surface); }
[data-testid="stAlert"] { border-radius: 10px; }
.mc-header, .mc-kpis, [class*="st-key-card_"] {
  animation: mc-enter .16s cubic-bezier(.22,1,.36,1) both;
}

/* Dashboard-only visual system. Other pages keep the shared application theme. */
.mc-dashboard-header { display: flex; flex-direction: column; justify-content: center;
  min-height: 58px; margin: 0 0 .625rem; padding: .15rem 0 .15rem 1rem;
  position: relative; }
.mc-dashboard-header::before { content: ""; position: absolute; left: 0; top: 5px; bottom: 5px;
  width: 4px; border-radius: 999px; background: linear-gradient(180deg,#173B6D,#B6193B); }
.mc-dashboard-product { color: #6C7D94; font-size: .7rem; line-height: 1.2;
  font-weight: 750; letter-spacing: .09em; text-transform: uppercase; }
.mc-dashboard-title { margin: .16rem 0 0; color: #10294F; font-size: 1.55rem;
  line-height: 1.1; font-weight: 750; letter-spacing: -.025em; }
.mc-reports-header { display: flex; align-items: center; justify-content: space-between; gap: 1rem;
  min-height: 58px; margin: 0 0 .625rem; padding: .15rem 0 .15rem 1rem; position: relative; }
.mc-reports-header::before { content: ""; position: absolute; left: 0; top: 5px; bottom: 5px;
  width: 4px; border-radius: 999px; background: linear-gradient(180deg,#173B6D,#B6193B); }
.mc-reports-product { color: #6C7D94; font-size: .7rem; line-height: 1.2; font-weight: 750;
  letter-spacing: .09em; text-transform: uppercase; }
.mc-reports-title { margin: .16rem 0 0; color: #10294F; font-size: 1.55rem; line-height: 1.1;
  font-weight: 750; letter-spacing: -.025em; }
.mc-reports-count { min-width: 118px; display: flex; align-items: center; justify-content: space-between;
  gap: .65rem; padding: .42rem .6rem; border: 1px solid rgba(255,255,255,.16);
  border-radius: 10px; background: linear-gradient(135deg,#102E59,#6D1835); color: #FFFFFF;
  box-shadow: 0 3px 10px rgba(8,26,54,.10); }
.mc-reports-count span { color: rgba(255,255,255,.72); font-size: .68rem; line-height: 1.15;
  font-weight: 700; letter-spacing: .045em; text-transform: uppercase; }
.mc-reports-count strong { font-size: 1.15rem; line-height: 1; font-weight: 750;
  font-variant-numeric: tabular-nums; }
[class*="st-key-reports_panel_"] { padding: .72rem .8rem .62rem;
  border: 1px solid #DDE4EE !important; border-radius: 14px !important; background: #FFFFFF;
  box-shadow: 0 3px 12px rgba(8,26,54,.045);
  animation: mc-dashboard-panel-enter .2s cubic-bezier(.22,1,.36,1) both; }
[class*="st-key-reports_panel_"] [data-testid="stVerticalBlock"] { gap: .32rem; }
[class*="st-key-reports_panel_"] h3 { margin: 0; color: #10294F; font-size: .96rem;
  line-height: 1.25; font-weight: 700; letter-spacing: -.01em; }
[class*="st-key-reports_panel_"] label p { color: #52647D; font-size: .75rem; font-weight: 650; }
[class*="st-key-reports_panel_selection"] [data-baseweb="select"] > div {
  min-height: 38px; border-color: #D3DCE8; border-radius: 9px; background: #F9FBFD; }
[class*="st-key-reports_panel_selection"] [data-baseweb="select"] > div:focus-within {
  border-color: #B6193B; box-shadow: 0 0 0 3px rgba(182,25,59,.10); }
[class*="st-key-reports_panel_selection"] [data-baseweb="tag"] { border: 0; border-radius: 7px;
  color: #FFFFFF; font-weight: 600; background: linear-gradient(112deg,#142F57,#A6193A);
  box-shadow: 0 2px 6px rgba(20,47,87,.18); }
[class*="st-key-reports_panel_selection"] [data-baseweb="tag"] svg { color: #FFFFFF; }
[class*="st-key-reports_panel_selection"] [data-testid="stPills"] button {
  min-height: 32px; padding: .28rem .72rem; border: 1px solid #D7E0EB; border-radius: 999px;
  background: #F7F9FC; color: #52647D; font-size: .74rem; font-weight: 650; }
[class*="st-key-reports_panel_selection"] [data-testid="stPills"] button[aria-pressed="true"] {
  border-color: #173B6D; background: #EAF0F8; color: #173B6D; }
[class*="st-key-reports_selection_header"] { width: 100%; }
[class*="st-key-reports_selection_header"] [data-testid="stHorizontalBlock"] {
  width: 100%; flex-wrap: nowrap !important; gap: .6rem; }
[class*="st-key-reports_selection_header"] [data-testid="stHorizontalBlock"] > :first-child {
  flex: 1 1 auto !important; min-width: 0; margin-right: auto; }
[class*="st-key-reports_selection_header"] h3 { white-space: nowrap; }
[class*="st-key-reports_selection_header"] [data-testid="stCheckbox"] label { justify-content: flex-end; }
[class*="st-key-reports_file_header"] { width: 100%; }
[class*="st-key-reports_file_header"] [data-testid="stHorizontalBlock"] {
  width: 100%; flex-wrap: nowrap !important; gap: .45rem; }
[class*="st-key-reports_file_header"] [data-testid="stHorizontalBlock"] > :first-child {
  flex: 1 1 auto !important; min-width: 0; margin-right: auto; }
[class*="st-key-reports_file_header"] h3 { margin-right: auto; white-space: nowrap; }
[class*="st-key-report_build_button"] button { min-height: 38px; padding: .45rem .85rem; border: 0;
  background: linear-gradient(118deg,#7C102A,#C32148) !important; color: #FFFFFF !important;
  font-weight: 700; box-shadow: 0 3px 10px rgba(124,16,42,.16); }
[class*="st-key-report_save_button"] button { min-height: 38px; padding: .45rem .85rem;
  border: 1px solid #CBD5E2; background: #FFFFFF !important; color: #334A68 !important;
  font-weight: 650; }
[class*="st-key-report_download_all"] button { min-height: 38px; padding: .45rem .85rem;
  border: 0; background: linear-gradient(118deg,#0B2347,#1D4B83) !important;
  color: #FFFFFF !important; font-weight: 700; box-shadow: 0 3px 10px rgba(16,46,89,.14); }
[class*="st-key-report_download_all"] button:disabled { opacity: .48; box-shadow: none; }
[class*="st-key-reports_panel_downloads"] [data-testid="stExpander"] [data-testid="stDownloadButton"] button {
  min-height: 34px; padding: .35rem .65rem; justify-content: flex-start; border: 1px solid #E0E6EF;
  background: #FFFFFF !important; color: #334A68 !important; font-size: .76rem;
  font-weight: 550; box-shadow: none; }
[class*="st-key-reports_individual_files"] [data-testid="stDownloadButton"] button {
  min-height: 34px; padding: .35rem .65rem; justify-content: flex-start;
  border: 1px solid #6E88AD !important; background: #FFFFFF !important;
  color: #173B6D !important; font-size: .76rem; font-weight: 550; text-align: left !important;
  box-shadow: none !important; }
[class*="st-key-reports_individual_files"] [data-testid="stDownloadButton"] button p,
[class*="st-key-reports_individual_files"] [data-testid="stDownloadButton"] button svg {
  color: #173B6D !important; text-align: left !important; }
.mc-reports-section-label { margin: .15rem 0 .45rem; color: #7A899D; font-size: .69rem;
  line-height: 1; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
[class*="st-key-manage_filter_"] { padding: .7rem .8rem .45rem; margin-bottom: .625rem;
  border: 1px solid #DDE4EE; border-radius: 14px; background: rgba(255,255,255,.96);
  box-shadow: 0 3px 12px rgba(8,26,54,.045); }
[class*="st-key-manage_filter_"] [data-testid="stVerticalBlock"] { gap: .35rem; }
[class*="st-key-manage_filter_"] [data-testid="stHorizontalBlock"] {
  align-items: flex-end; gap: .65rem; }
[class*="st-key-manage_filter_"] label p { color:#52647D; font-size:.75rem; font-weight:650; }
[class*="st-key-manage_filter_"] [data-baseweb="select"] > div,
[class*="st-key-manage_filter_"] [data-baseweb="input"] {
  min-height: 40px; border-color:#D3DCE8; border-radius:9px; background:#F9FBFD; }
[class*="st-key-manage_filter_"] [data-baseweb="select"] > div:focus-within,
[class*="st-key-manage_filter_"] [data-baseweb="input"]:focus-within {
  border-color:#B6193B; box-shadow:0 0 0 3px rgba(182,25,59,.10); }
[class*="st-key-manage_filter_"] [data-testid="stSegmentedControl"] button {
  min-height: 36px; font-size:.76rem; }
[class*="st-key-manage_panel_"] { overflow:hidden; padding:.78rem .85rem .7rem;
  margin-bottom:.625rem; border:1px solid #DDE4EE; border-radius:14px;
  background:#FFFFFF; box-shadow:0 3px 12px rgba(8,26,54,.045); }
[class*="st-key-manage_panel_"] [data-testid="stVerticalBlock"] { gap:.38rem; }
[class*="st-key-manage_panel_"] h3 { margin:0; color:#10294F; font-size:.96rem;
  line-height:1.25; font-weight:700; letter-spacing:-.01em; }
[class*="st-key-manage_panel_"] [data-testid="stCaptionContainer"] p {
  margin:0; color:#718198; font-size:.72rem; line-height:1.35; }
[class*="st-key-manage_panel_"] [data-testid="stDataFrame"],
[class*="st-key-manage_panel_"] [data-testid="stDataEditor"] {
  margin-top:.15rem; border-radius:10px; overflow:hidden; }
[class*="st-key-manage_panel_"] [data-testid="stFileUploaderDropzone"] {
  min-height:86px; padding:.75rem; border:1px dashed #AEBED2; border-radius:11px;
  background:#F8FAFD; }
[class*="st-key-manage_panel_"] [data-testid="stExpander"] details {
  border:1px solid #DDE4EE; background:#F9FBFD; }
[class*="st-key-manage_actions_"] { margin:.1rem 0 .625rem; }
[class*="st-key-manage_actions_"] [data-testid="stHorizontalBlock"] {
  display:flex !important; flex-wrap:wrap !important; gap:.5rem !important; }
[class*="st-key-manage_primary_"] button { min-height:38px; padding:.45rem .85rem;
  border:1px solid #173B6D !important; background:#173B6D !important; color:#FFFFFF !important;
  font-weight:700; box-shadow:0 3px 10px rgba(16,46,89,.14); }
[class*="st-key-manage_primary_"] button:hover { background:#214A7C !important;
  border-color:#214A7C !important; }
[class*="st-key-manage_primary_"] button:disabled { opacity:.48; box-shadow:none; }
[class*="st-key-manage_danger_"] button { min-height:38px; padding:.45rem .85rem;
  border:1px solid #B95B71 !important; background:#FFFFFF !important; color:#8F2440 !important;
  font-weight:650; box-shadow:none !important; }
[class*="st-key-manage_danger_"] button:disabled { opacity:.45; }
[class*="st-key-manage_import_sources"] { margin-bottom:.625rem; }
[class*="st-key-manage_import_sources"] [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"]) {
  display:grid !important; grid-template-columns:repeat(2,minmax(0,1fr)) !important;
  align-items:stretch; gap:.625rem !important; }
[class*="st-key-manage_import_sources"] [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"]) > [data-testid="stColumn"] {
  width:auto !important; min-width:0 !important; flex:none !important;
  padding:.72rem .8rem .62rem; border:1px solid #DDE4EE;
  border-radius:14px; background:#FFFFFF; box-shadow:0 3px 12px rgba(8,26,54,.045); }
[class*="st-key-manage_import_sources"] [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"])
  > [data-testid="stColumn"] > [data-testid="stVerticalBlock"] { gap:.42rem; height:100%; }
[class*="st-key-manage_import_sources"] h3 { margin:0; color:#10294F; font-size:.94rem;
  line-height:1.25; font-weight:700; }
[class*="st-key-manage_import_sources"] [data-testid="stCaptionContainer"] {
  position:relative; z-index:1; margin:0 0 .18rem !important; }
[class*="st-key-manage_import_sources"] [data-testid="stCaptionContainer"] p {
  margin:0; color:#718198; font-size:.7rem; line-height:1.3; }
[class*="st-key-manage_import_sources"] [data-testid="stFileUploader"] {
  min-width:0; margin-top:.12rem; }
[class*="st-key-manage_import_sources"] [data-testid="stFileUploaderDropzone"] {
  display:flex !important; flex-direction:row !important; align-items:center !important;
  justify-content:flex-start !important; gap:.85rem !important;
  min-height:70px; padding:.55rem .65rem; border:1px dashed #9FB2CA;
  border-radius:10px; background:#F8FAFD; }
[class*="st-key-manage_import_sources"] [data-testid="stFileUploaderDropzone"] button {
  flex:0 0 auto; align-self:center !important; min-height:34px; margin:0 !important;
  padding:.35rem .65rem; border-color:#CBD6E4; color:#173B6D; }
[class*="st-key-manage_import_sources"] [data-testid="stFileUploaderDropzone"] > div,
[class*="st-key-manage_import_sources"] [data-testid="stFileUploaderDropzoneInstructions"] {
  align-self:center !important; margin:0 !important; }
[class*="st-key-manage_import_sources"] [data-testid="stFileUploaderDropzone"] small {
  display:block; margin:0 !important; line-height:1.25; }
[class*="st-key-manage_panel_tracker_import"] [data-testid="stStatusWidget"],
[class*="st-key-manage_panel_ebiz_import"] [data-testid="stStatusWidget"] {
  border-radius:10px; }
[class*="st-key-cross_reference_filter_bar"],
[class*="st-key-discrepancy_filter_bar"],
[class*="st-key-rebate_filter_bar"] { padding: .7rem .8rem .35rem;
  margin-bottom: .625rem; border: 1px solid #DDE4EE; border-radius: 14px;
  background: rgba(255,255,255,.95); box-shadow: 0 3px 12px rgba(8,26,54,.045); }
[class*="st-key-cross_reference_filter_bar"] label p,
[class*="st-key-discrepancy_filter_bar"] label p,
[class*="st-key-rebate_filter_bar"] label p { color: #52647D;
  font-size: .75rem; font-weight: 650; }
[class*="st-key-cross_reference_filter_bar"] [data-baseweb="select"] > div,
[class*="st-key-discrepancy_filter_bar"] [data-baseweb="select"] > div,
[class*="st-key-rebate_filter_bar"] [data-baseweb="select"] > div {
  min-height: 40px; border-color: #D3DCE8; border-radius: 9px; background: #F9FBFD; }
[class*="st-key-cross_reference_filter_bar"] [data-baseweb="select"] > div:focus-within,
[class*="st-key-discrepancy_filter_bar"] [data-baseweb="select"] > div:focus-within,
[class*="st-key-rebate_filter_bar"] [data-baseweb="select"] > div:focus-within {
  border-color: #B6193B; box-shadow: 0 0 0 3px rgba(182,25,59,.10); }
[class*="st-key-rebate_secondary_filters"] { margin-top: .15rem; }
[class*="st-key-rebate_secondary_filters"] [data-testid="stHorizontalBlock"] {
  align-items: flex-end; gap: .75rem; }
[class*="st-key-rebate_secondary_filters"] [data-testid="stPills"] button {
  min-height: 32px; padding: .28rem .72rem; white-space: nowrap; }
[class*="st-key-rebate_secondary_filters"] [data-testid="stRadio"] [role="radiogroup"] {
  display: flex; flex-wrap: wrap; gap: .35rem .8rem; }
[class*="st-key-cross_reference_panel_"] { overflow: hidden; padding: .78rem .85rem .65rem;
  border: 1px solid #DDE4EE !important; border-radius: 14px !important; background: #FFFFFF;
  box-shadow: 0 3px 12px rgba(8,26,54,.045);
  animation: mc-dashboard-panel-enter .2s cubic-bezier(.22,1,.36,1) both; }
[class*="st-key-cross_reference_panel_"] [data-testid="stVerticalBlock"] { gap: .32rem; }
[class*="st-key-cross_reference_panel_"] h3 { margin: 0; color: #10294F;
  font-size: .96rem; line-height: 1.25; font-weight: 700; }
[class*="st-key-cross_reference_panel_"] [data-testid="stTextInput"] input {
  min-height: 36px; border-radius: 9px; background: #F9FBFD; }
[class*="st-key-cross_reference_download"] button { min-height: 38px; padding: .45rem .85rem;
  border: 0; background: linear-gradient(118deg,#0B2347,#1D4B83) !important;
  color: #FFFFFF !important; font-weight: 700; box-shadow: 0 3px 10px rgba(16,46,89,.14); }
.mc-cross-reference-section-label { margin: .15rem 0 .45rem; color: #7A899D;
  font-size: .69rem; line-height: 1; font-weight: 700; letter-spacing: .08em;
  text-transform: uppercase; }
[class*="st-key-dashboard_filter_bar"] { padding: .75rem .85rem .35rem;
  margin-bottom: .625rem; border: 1px solid #DDE4EE; border-radius: 14px;
  background: rgba(255,255,255,.95); box-shadow: 0 4px 18px rgba(8,26,54,.055); }
[class*="st-key-dashboard_filter_bar"] [data-testid="stVerticalBlock"] { gap: .35rem; }
[class*="st-key-dashboard_filter_bar"] label p { color: #52647D; font-size: .75rem;
  font-weight: 650; letter-spacing: .015em; }
[class*="st-key-dashboard_filter_bar"] [data-baseweb="select"] > div {
  min-height: 42px; border-color: #D3DCE8; border-radius: 10px; background: #F9FBFD; }
[class*="st-key-dashboard_filter_bar"] [data-baseweb="select"] > div:focus-within {
  border-color: #B6193B; box-shadow: 0 0 0 3px rgba(182,25,59,.10); }
[class*="st-key-dashboard_filter_bar"] [data-baseweb="tag"] {
  border: 0; border-radius: 7px; color: #FFFFFF; font-weight: 600;
  background: linear-gradient(112deg,#142F57,#A6193A); box-shadow: 0 2px 6px rgba(20,47,87,.18); }
[class*="st-key-dashboard_filter_bar"] [data-baseweb="tag"] svg { color: #FFFFFF; }
.mc-dashboard-metrics { display: grid;
  grid-template-columns: repeat(var(--mc-metric-columns,4),minmax(0,1fr));
  gap: .625rem; margin: 0 0 .625rem; }
.mc-dashboard-metric { min-width: 0; min-height: 112px; display: flex; flex-direction: column;
  justify-content: space-between; padding: .85rem .95rem; overflow: hidden; position: relative;
  border: 1px solid #DDE4EE; border-radius: 14px; background: #FFFFFF;
  box-shadow: 0 4px 18px rgba(8,26,54,.055); color: inherit; text-decoration: none;
  cursor: default; }
.mc-dashboard-metric--active { border-color: #B6193B;
  box-shadow: 0 0 0 3px rgba(182,25,59,.13), 0 8px 22px rgba(8,26,54,.12); }
.mc-dashboard-metric--navy.mc-dashboard-metric--active,
.mc-dashboard-metric--red.mc-dashboard-metric--active { border-color: rgba(255,255,255,.82); }
.mc-dashboard-metric::after { content: ""; position: absolute; width: 72px; height: 72px;
  right: -25px; bottom: -32px; border-radius: 50%; background: rgba(20,47,87,.06); }
.mc-dashboard-metric--navy { color: #FFFFFF; border-color: #142F57;
  background: linear-gradient(138deg,#091D3C,#173B6D); }
.mc-dashboard-metric--red { color: #FFFFFF; border-color: #A81737;
  background: linear-gradient(138deg,#7C102A,#C32148); }
.mc-dashboard-metric--matte-navy-deep { color:#FFFFFF; border-color:#112E55; background:#112E55; }
.mc-dashboard-metric--matte-crimson-deep { color:#FFFFFF; border-color:#7D1936; background:#7D1936; }
.mc-dashboard-metric--matte-navy-mid { color:#FFFFFF; border-color:#274D78; background:#274D78; }
.mc-dashboard-metric--matte-crimson-mid { color:#FFFFFF; border-color:#942E4B; background:#942E4B; }
.mc-dashboard-metric--matte-navy-soft { color:#FFFFFF; border-color:#42658E; background:#42658E; }
.mc-dashboard-metric--matte-crimson-soft { color:#FFFFFF; border-color:#A84662; background:#A84662; }
.mc-dashboard-metric--navy::after, .mc-dashboard-metric--red::after {
  background: rgba(255,255,255,.09); }
.mc-dashboard-metric[class*="--matte-"]::after { background:rgba(255,255,255,.09); }
.mc-dashboard-metric__label { position: relative; z-index: 1; color: #65758B;
  font-size: .72rem; line-height: 1.2; font-weight: 700; letter-spacing: .055em;
  text-transform: uppercase; }
.mc-dashboard-metric--navy .mc-dashboard-metric__label,
.mc-dashboard-metric--red .mc-dashboard-metric__label { color: rgba(255,255,255,.72); }
.mc-dashboard-metric[class*="--matte-"] .mc-dashboard-metric__label {
  color:rgba(255,255,255,.78); }
.mc-dashboard-metric__value { position: relative; z-index: 1; color: #10294F;
  font-size: 1.85rem; line-height: 1; font-weight: 750; letter-spacing: -.04em;
  font-variant-numeric: tabular-nums; }
.mc-dashboard-metric--navy .mc-dashboard-metric__value,
.mc-dashboard-metric--red .mc-dashboard-metric__value { color: #FFFFFF; }
.mc-dashboard-metric[class*="--matte-"] .mc-dashboard-metric__value { color:#FFFFFF; }
.mc-dashboard-metric__note { position: relative; z-index: 1; color: #78889D;
  font-size: .72rem; line-height: 1.2; }
.mc-dashboard-metric--navy .mc-dashboard-metric__note,
.mc-dashboard-metric--red .mc-dashboard-metric__note { color: rgba(255,255,255,.68); }
.mc-dashboard-metric[class*="--matte-"] .mc-dashboard-metric__note {
  color:rgba(255,255,255,.74); }
[class*="st-key-dashboard_panel_"] { overflow: hidden; padding: .82rem .9rem .65rem;
  border: 1px solid #DDE4EE !important; border-radius: 14px !important;
  background: #FFFFFF; box-shadow: 0 4px 18px rgba(8,26,54,.055); }
[class*="st-key-dashboard_panel_"] [data-testid="stVerticalBlock"] { gap: .3rem; }
[class*="st-key-dashboard_panel_"] h3 { margin: 0; color: #10294F;
  font-size: .96rem; line-height: 1.25; font-weight: 700; letter-spacing: -.01em; }
[class*="st-key-dashboard_panel_"] [data-testid="stCaptionContainer"] p {
  margin: 0; font-size: .72rem; line-height: 1.3; }
[class*="st-key-dashboard_panel_"] [data-testid="stVegaLiteChart"] { margin-top: -.2rem; }
[class*="st-key-dashboard_panel_"] [data-testid="stDataFrame"] { margin-top: .15rem; }
[class*="st-key-dashboard_panel_"] [data-testid="stTextInput"] input {
  min-height: 38px; border-radius: 9px; background: #F9FBFD; }
[class*="st-key-dashboard_panel_"] { animation: mc-dashboard-panel-enter .2s cubic-bezier(.22,1,.36,1) both; }
.mc-dashboard-section-label { margin: .15rem 0 .45rem; color: #7A899D;
  font-size: .69rem; line-height: 1; font-weight: 700; letter-spacing: .08em;
  text-transform: uppercase; }
[data-testid="stSidebarNavLink"], .stButton > button, .stDownloadButton > button,
.stFormSubmitButton > button, .mc-kpi, .mc-status {
  transition: background-color .16s ease, border-color .16s ease,
    box-shadow .16s ease, color .16s ease, transform .16s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover,
.stFormSubmitButton > button:hover { transform: translateY(-1px); }
@keyframes mc-enter {
  from { opacity: .72; transform: translateY(3px); }
  to { opacity: 1; transform: translateY(0); }
}
@keyframes mc-dashboard-panel-enter {
  from { opacity: .58; transform: translateY(4px); }
  to { opacity: 1; transform: translateY(0); }
}
@media (prefers-reduced-motion: reduce) {
  .mc-header, .mc-kpis, [class*="st-key-card_"] { animation: none; }
  [class*="st-key-dashboard_panel_"] { animation: none; }
  [class*="st-key-reports_panel_"] { animation: none; }
  [class*="st-key-cross_reference_panel_"] { animation: none; }
  [data-testid="stSidebarNavLink"], .stButton > button, .stDownloadButton > button,
  .stFormSubmitButton > button, .mc-kpi, .mc-status { transition: none; }
}
@media (max-width: 1100px) { .mc-kpis { grid-template-columns: repeat(2,minmax(0,1fr)); } }
@media (max-width: 1100px) { .mc-dashboard-metrics { grid-template-columns: repeat(2,minmax(0,1fr)); } }
@media (max-width: 1024px) { :root { --mc-page-inline: 1.25rem; } }
@media (max-width: 900px) {
  [class*="st-key-rebate_secondary_filters"] [data-testid="stHorizontalBlock"] {
    display: grid !important; grid-template-columns: minmax(130px,.65fr) minmax(260px,1.35fr) !important;
    gap: .5rem .75rem !important; }
  [class*="st-key-rebate_secondary_filters"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
    width: auto !important; min-width: 0 !important; flex: none !important; }
  [class*="st-key-rebate_secondary_filters"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(3) {
    grid-column: 1 / -1; }
  [class*="st-key-manage_filter_"] [data-testid="stHorizontalBlock"] {
    gap:.5rem !important; }
}
@media (max-width: 640px) {
  :root { --mc-page-top: .75rem; --mc-page-inline: .75rem; --mc-page-bottom: 1.25rem; }
  [data-testid="stHeader"]::before { display: none; }
  [data-testid="stHeader"]::after { right: 12px; width: 68px; }
  .mc-topbrand-spacer { height: 60px; flex-basis: 60px; }
  .mc-title { font-size: 1.375rem; } .mc-desc { font-size: .875rem; }
  .mc-status { align-items: flex-start; }
  .mc-kpis { grid-template-columns: repeat(2,minmax(0,1fr)); gap: .5rem; }
  .mc-kpi { padding: .75rem; } .mc-kpi__value { font-size: 1.25rem; }
  .mc-dashboard-header { min-height: 52px; padding-left: .8rem; }
  .mc-dashboard-title { font-size: 1.3rem; }
  .mc-dashboard-product { font-size: .62rem; }
  .mc-reports-header { min-height: 52px; padding-left: .8rem; }
  .mc-reports-title { font-size: 1.3rem; }
  .mc-reports-product { font-size: .62rem; }
  .mc-reports-count { min-width: 120px; padding: .48rem .6rem; }
  .mc-reports-count span { font-size: .6rem; }
  .mc-reports-count strong { font-size: 1.15rem; }
  [class*="st-key-rebate_secondary_filters"] [data-testid="stHorizontalBlock"] {
    grid-template-columns: minmax(0,1fr) !important; }
  [class*="st-key-rebate_secondary_filters"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]:nth-child(3) {
    grid-column: auto; }
  [class*="st-key-manage_filter_"] { padding:.65rem; }
  [class*="st-key-manage_filter_"] [data-testid="stHorizontalBlock"] {
    display:grid !important; grid-template-columns:minmax(0,1fr) !important; }
  [class*="st-key-manage_filter_"] [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
    width:auto !important; min-width:0 !important; flex:none !important; }
  [class*="st-key-manage_panel_"] { padding:.65rem; }
  [class*="st-key-manage_actions_"] [data-testid="stHorizontalBlock"] {
    display:grid !important; grid-template-columns:minmax(0,1fr) minmax(0,1fr) !important; }
  [class*="st-key-manage_actions_"] button { width:100%; }
  [class*="st-key-manage_import_sources"] [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"]) {
    display:grid !important;
    grid-template-columns:minmax(0,1fr) !important; }
  [class*="st-key-manage_import_sources"] [data-testid="stHorizontalBlock"]:has(> [data-testid="stColumn"]) > [data-testid="stColumn"] {
    width:auto !important; min-width:0 !important; flex:none !important; padding:.65rem; }
  [class*="st-key-reports_selection_header"] [data-testid="stHorizontalBlock"] {
    display: flex !important; flex-wrap: nowrap !important; gap: .35rem !important; }
  [class*="st-key-reports_selection_header"] h3 { font-size: .88rem; }
  [class*="st-key-reports_selection_header"] [data-testid="stCheckbox"] p { font-size: .66rem; }
  [class*="st-key-reports_file_header"] [data-testid="stHorizontalBlock"] {
    display: flex !important; flex-wrap: nowrap !important; gap: .35rem !important; }
  [class*="st-key-reports_file_header"] h3 { font-size: .88rem; }
  [class*="st-key-report_build_button"] button,
  [class*="st-key-report_download_all"] button,
  [class*="st-key-report_save_button"] button {
    display: inline-flex !important; align-items: center !important; justify-content: center !important;
    position: relative !important; gap: 0 !important;
    width: 36px !important; min-width: 36px !important; max-width: 36px !important;
    height: 36px; min-height: 36px; padding: 0 !important; border-radius: 9px; }
  [class*="st-key-report_build_button"] button p,
  [class*="st-key-report_download_all"] button p,
  [class*="st-key-report_save_button"] button p { display: none !important; }
  [class*="st-key-report_build_button"] button svg,
  [class*="st-key-report_download_all"] button svg,
  [class*="st-key-report_save_button"] button svg,
  [class*="st-key-report_build_button"] button [data-testid="stIconMaterial"],
  [class*="st-key-report_download_all"] button [data-testid="stIconMaterial"],
  [class*="st-key-report_save_button"] button [data-testid="stIconMaterial"] {
    position: absolute !important; left: 50% !important; top: 50% !important;
    margin: 0 !important; transform: translate(-50%,-50%) !important; }
  [class*="st-key-report_build_button"] button > svg,
  [class*="st-key-report_download_all"] button > svg,
  [class*="st-key-report_save_button"] button > svg {
    position: absolute !important; left: 50% !important; top: 50% !important;
    margin: 0 !important; transform: translate(-50%,-50%) !important; }
  [class*="st-key-report_build_button"] button [data-testid="stIconMaterial"] svg,
  [class*="st-key-report_download_all"] button [data-testid="stIconMaterial"] svg,
  [class*="st-key-report_save_button"] button [data-testid="stIconMaterial"] svg {
    position: static !important; transform: none !important; }
  [class*="st-key-reports_individual_files"] [data-testid="stDownloadButton"] button,
  [class*="st-key-reports_individual_files"] [data-testid="stDownloadButton"] button p {
    font-size: .67rem !important; line-height: 1.35; text-align: left !important; }
  .mc-dashboard-metrics { grid-template-columns: repeat(2,minmax(0,1fr)); gap: .5rem; }
  .mc-dashboard-metric { min-height: 102px; padding: .75rem; }
  .mc-dashboard-metric__value { font-size: 1.5rem; }
  [class*="st-key-dashboard_panel_"] { padding: .75rem; }
  [class*="st-key-reports_panel_"] { padding: .65rem; }
  [class*="st-key-cross_reference_panel_"] { padding: .65rem; }
}
"""


def apply_theme() -> None:
    css = THEME.replace("__MATALAN_LOGO__", _data_uri(MATALAN_LOGO))
    st.html(f"<style>{css}</style>")


def app_chrome() -> None:
    """Render stable app chrome before the selected page does any work."""
    apply_theme()
    st.html("<div class='mc-topbrand-spacer' aria-hidden='true'></div>")


def header(title: str, description: str | None = None,
           status_label: str | None = None, status_value: str | None = None) -> None:
    desc = f"<p class='mc-desc'>{html.escape(description)}</p>" if description else ""
    status = ""
    if status_label and status_value:
        status = ("<div class='mc-status'>"
                  f"<span class='mc-status__label'>{html.escape(status_label)}</span>"
                  f"<span class='mc-status__value'>{html.escape(status_value)}</span></div>")
    st.html(
        "<div class='mc-header'><div class='mc-header__main'>"
        "<div class='mc-eyebrow'>PACD · Matalan</div>"
        f"<div class='mc-title' role='heading' aria-level='1'>{html.escape(title)}</div>{desc}"
        f"</div>{status}</div>")


def dashboard_header() -> None:
    """Render the compact, branding-free Dashboard heading."""
    st.html(
        "<section class='mc-dashboard-header'>"
        "<div class='mc-dashboard-product'>Matalan Carton Supply Report</div>"
        "<h1 class='mc-dashboard-title'>Dashboard</h1>"
        "</section>"
    )


def reports_header(available_suppliers: int) -> None:
    """Render the compact Build reports heading and supplier-count badge."""
    st.html(
        "<section class='mc-reports-header'>"
        "<div><div class='mc-reports-product'>Matalan Carton Supply Report</div>"
        "<h1 class='mc-reports-title'>Build reports</h1></div>"
        "<div class='mc-reports-count' aria-label='Available suppliers'>"
        "<span>Available<br>suppliers</span>"
        f"<strong>{int(available_suppliers):,}</strong></div>"
        "</section>"
    )


def cross_reference_header(ebiz_rows: int | None) -> None:
    """Render the compact Cross reference heading and eBiz status badge."""
    report_page_header("Cross reference", "eBiz<br>rows", f"{int(ebiz_rows or 0):,}")


def report_page_header(title: str, badge_label: str, badge_value: str) -> None:
    """Render the shared compact report heading and status badge."""
    st.html(
        "<section class='mc-reports-header'>"
        "<div><div class='mc-reports-product'>Matalan Carton Supply Report</div>"
        f"<h1 class='mc-reports-title'>{html.escape(title)}</h1></div>"
        f"<div class='mc-reports-count' aria-label='{html.escape(badge_label.replace('<br>', ' '))}'>"
        f"<span>{badge_label}</span>"
        f"<strong>{html.escape(str(badge_value))}</strong></div>"
        "</section>"
    )


def dashboard_metric_grid(tiles: list[dict], columns: int | None = None) -> None:
    """Render dense responsive Dashboard metrics without Streamlit column stacking."""
    cards = []
    for tile in tiles:
        tone = f" mc-dashboard-metric--{tile['tone']}" if tile.get("tone") else ""
        cards.append(
            f"<article class='mc-dashboard-metric{tone}'>"
            f"<div class='mc-dashboard-metric__label'>{html.escape(str(tile['label']))}</div>"
            f"<div class='mc-dashboard-metric__value'>{html.escape(str(tile['value']))}</div>"
            f"<div class='mc-dashboard-metric__note'>{html.escape(str(tile['note']))}</div>"
            "</article>"
        )
    count = max(1, min(int(columns or len(tiles) or 1), 6))
    st.html(
        f"<section class='mc-dashboard-metrics' style='--mc-metric-columns:{count}'>"
        + "".join(cards) + "</section>"
    )


def dashboard_interactive_metric_grid(
        tiles: list[dict], key: str, default: str = "all", columns: int = 4) -> str:
    """Render stable metric cards that rerun only their containing fragment."""
    valid = {str(tile.get("id", index)) for index, tile in enumerate(tiles)}
    state_key = f"{key}_active"
    active = st.session_state.get(state_key, default)
    if active not in valid:
        active = default if default in valid else next(iter(valid))
        st.session_state[state_key] = active

    payload = []
    for index, tile in enumerate(tiles):
        payload.append({
            "id": str(tile.get("id", index)),
            "label": str(tile["label"]),
            "value": str(tile["value"]),
            "note": str(tile.get("note", "")),
            "tone": str(tile.get("tone", "")),
        })
    component_key = f"{key}_cards"

    def sync_selected_card() -> None:
        component_state = st.session_state.get(component_key, {})
        selected = component_state.get("selected") if isinstance(component_state, dict) else None
        if selected in valid:
            st.session_state[state_key] = selected

    result = _DASHBOARD_CARD_COMPONENT(
        key=component_key,
        data={"tiles": payload, "active": active, "columns": max(1, min(int(columns), 6))},
        default={"selected": active},
        on_selected_change=sync_selected_card,
    )
    selected = getattr(result, "selected", None)
    if selected in valid:
        active = selected
        st.session_state[state_key] = active
    return active


def kpi_grid(tiles: list[dict]) -> None:
    cards = []
    for tile in tiles:
        tone = f" mc-kpi--{tile['tone']}" if tile.get("tone") else ""
        note = (f"<div class='mc-kpi__note'>{html.escape(str(tile['note']))}</div>"
                if tile.get("note") else "")
        cards.append(
            f"<div class='mc-kpi{tone}'><div class='mc-kpi__label'>"
            f"{html.escape(str(tile['label']))}</div><div class='mc-kpi__value'>"
            f"{html.escape(str(tile['value']))}</div>{note}</div>")
    st.html(f"<div class='mc-kpis' style='--mc-kpi-count:{min(len(tiles), 6)}'>"
            + "".join(cards) + "</div>")


def _select_kpi(state_key: str, tile_id: str) -> None:
    """Apply a card selection before Streamlit renders the next script run."""
    st.session_state[state_key] = tile_id


def interactive_kpi_grid(tiles: list[dict], key: str, default: str = "all") -> str:
    """Render KPI cards as real buttons and return the active tile id."""
    state_key = f"{key}_active"
    valid = {str(tile.get("id", index)) for index, tile in enumerate(tiles)}
    if st.session_state.get(state_key) not in valid:
        st.session_state[state_key] = default if default in valid else next(iter(valid))
    columns = st.columns(len(tiles), gap="small")
    for index, (column, tile) in enumerate(zip(columns, tiles)):
        tile_id = str(tile.get("id", index))
        label = f"{tile['label']}\n{tile['value']}"
        if tile.get("note"):
            label += f"\n{tile['note']}"
        column.button(
            label,
            key=f"kpi_card_{key}_{tile_id}",
            use_container_width=True,
            type="primary" if st.session_state[state_key] == tile_id else "secondary",
            help=tile.get("help"),
            on_click=_select_kpi,
            args=(state_key, tile_id),
        )
    return st.session_state[state_key]


def search_frame(frame, query: str):
    """Case-insensitive contains search across every visible dataframe column."""
    query = (query or "").strip().casefold()
    if not query or frame.empty:
        return frame
    text = frame.fillna("").astype(str)
    mask = text.apply(lambda column: column.str.casefold().str.contains(
        query, regex=False, na=False)).any(axis=1)
    return frame.loc[mask]


def table_search(label: str, key: str, placeholder: str = "Search all columns") -> str:
    return st.text_input(label, key=key, placeholder=placeholder,
                         icon=":material/search:", label_visibility="collapsed")


def clear_built_reports() -> None:
    """Drop prepared downloads after the data behind them changes."""
    for key in ("built_reports", "live_discrepancy_pdf", "live_rebate_pdf"):
        st.session_state.pop(key, None)


@contextmanager
def operation_status(label: str, complete: str):
    """Show one consistent animated status while a server operation is running."""
    status = st.status(label, state="running", expanded=False, type="compact")
    try:
        yield
    except Exception:
        status.update(label=f"{label.rstrip('…')} failed", state="error", expanded=False)
        raise
    else:
        status.update(label=complete, state="complete", expanded=False)


def download_started(file_name: str) -> None:
    """Confirm that a prepared file was handed to the browser."""
    st.toast(f"Download started: {file_name}", icon=":material/download:")


def page(fn):
    """Show a circular spinner once and replace raw page exceptions with a useful panel."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        first_load_key = f"page_loaded_{fn.__module__}.{fn.__name__}"
        first_load = not st.session_state.get(first_load_key, False)
        st.session_state[first_load_key] = True
        try:
            if first_load:
                with st.spinner("Loading…"):
                    return fn(*args, **kwargs)
            return fn(*args, **kwargs)
        except Exception as exc:
            st.error("This page could not be loaded. Try again, then contact support if it continues.")
            with st.expander("Technical details"):
                st.exception(exc)
            st.stop()
    return wrapper
