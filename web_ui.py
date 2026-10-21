"""Shared responsive page chrome for the hosted Streamlit website."""
from __future__ import annotations

import base64
import functools
import html
import traceback
from datetime import date
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent
APP_NAME = "Matalan Carton Supply Report"
PACD_LOGO = str(ROOT / "assets" / "logo_left.png")
MATALAN_LOGO = str(ROOT / "assets" / "logo_right.png")
FAVICON = str(ROOT / "assets" / "favicon.png")


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
[data-testid="stMainBlockContainer"] { max-width: 1560px; padding: 1rem 1.5rem 1.5rem; }
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
.mc-footer-chip { position: fixed !important; right: 12px !important; bottom: 10px !important;
  z-index: 1000000;
  padding: .42rem .75rem; border: 1px solid rgba(226,232,240,.95); border-radius: 999px;
  background: rgba(255,255,255,.94); box-shadow: 0 4px 14px rgba(15,27,45,.10);
  backdrop-filter: blur(10px); white-space: nowrap; text-align: center; width: auto !important;
  font-size: .72rem; line-height: 1; color: #7C8CA3; }
.mc-footer-chip a { color: var(--mc-navy); font-weight: 600; text-decoration: none; }
.mc-footer-chip a:hover { color: var(--mc-sky); text-decoration: underline; }
.mc-header, .mc-kpis, [class*="st-key-card_"] {
  animation: mc-enter .16s cubic-bezier(.22,1,.36,1) both;
}
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
@media (prefers-reduced-motion: reduce) {
  .mc-header, .mc-kpis, [class*="st-key-card_"] { animation: none; }
  [data-testid="stSidebarNavLink"], .stButton > button, .stDownloadButton > button,
  .stFormSubmitButton > button, .mc-kpi, .mc-status { transition: none; }
}
@media (max-width: 1100px) { .mc-kpis { grid-template-columns: repeat(2,minmax(0,1fr)); } }
@media (max-width: 1024px) { [data-testid="stMainBlockContainer"] { padding: 1rem 1.25rem 1.5rem; } }
@media (max-width: 640px) {
  [data-testid="stMainBlockContainer"] { padding: .75rem .75rem 1.25rem; }
  [data-testid="stHeader"]::before { display: none; }
  [data-testid="stHeader"]::after { right: 12px; width: 68px; }
  .mc-topbrand-spacer { height: 60px; flex-basis: 60px; }
  .mc-title { font-size: 1.375rem; } .mc-desc { font-size: .875rem; }
  .mc-status { align-items: flex-start; }
  .mc-kpis { grid-template-columns: repeat(2,minmax(0,1fr)); gap: .5rem; }
  .mc-kpi { padding: .75rem; } .mc-kpi__value { font-size: 1.25rem; }
  .mc-footer-chip { right: 8px !important; bottom: 8px !important;
    padding: .4rem .65rem; font-size: .68rem; }
}
"""


def apply_theme() -> None:
    css = THEME.replace("__MATALAN_LOGO__", _data_uri(MATALAN_LOGO))
    st.html(f"<style>{css}</style>")


def _footer_chip_html() -> str:
    year = date.today().year
    chip_style = (
        "position:fixed!important;right:12px!important;bottom:10px!important;"
        "z-index:1000000!important;display:inline-flex!important;align-items:center!important;"
        "width:auto!important;margin:0!important;padding:7px 12px!important;"
        "border:1px solid #e2e8f0!important;border-radius:999px!important;"
        "background:rgba(255,255,255,.96)!important;box-shadow:0 4px 14px rgba(15,27,45,.10)!important;"
        "white-space:nowrap!important;font-size:11.5px!important;line-height:1!important;color:#7c8ca3!important"
    )
    return (
        f"<div class='mc-footer-chip' style='{chip_style}'>&copy; {year} PACD&nbsp; | &nbsp;Developed by&nbsp;"
        "<a href='https://github.com/RecurSID' target='_blank' "
        "rel='noopener noreferrer' style='color:#002b61;font-weight:700;text-decoration:none'>"
        "recurSID</a></div>"
    )


def app_chrome() -> None:
    """Render stable app chrome before the selected page does any work."""
    apply_theme()
    st.html(
        "<div class='mc-topbrand-spacer' aria-hidden='true'></div>"
        f"{_footer_chip_html()}"
    )


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


def page(title: str, description: str | None = None):
    """Replace raw page exceptions with a useful panel."""
    def decorate(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                result = fn(*args, **kwargs)
                return result
            except Exception as exc:
                st.error("This page could not be loaded. Try again, then contact support if it continues.")
                with st.expander("Technical details"):
                    st.code(f"{exc}\n\n{traceback.format_exc()}", language="text")
                st.stop()
        return wrapper
    return decorate
