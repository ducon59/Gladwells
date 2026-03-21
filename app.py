import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import msoffcrypto
import openpyxl
import io
import os
import tempfile
import gdown
from datetime import datetime
import time

# ─── Config ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gladwells Management Accounts",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

FILE_PASSWORD = "Dunnock057"
GDRIVE_FILE_ID = os.environ.get("GDRIVE_FILE_ID", "")   # Set via env var or secrets

# ─── Colours ──────────────────────────────────────────────────────────────────
PRIMARY   = "#2C5F2E"   # deep forest green
ACCENT    = "#97BC62"   # lime green
ORANGE    = "#E8623A"   # warm orange for negative
BG        = "#F5F7F2"
CARD_BG   = "#FFFFFF"
TEXT_DARK = "#1A2E1A"

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; background: {BG}; }}
  .main {{ background: {BG}; }}
  .stSidebar {{ background: {PRIMARY}; }}
  .stSidebar * {{ color: white !important; }}
  .stSidebar .stSelectbox label, .stSidebar .stRadio label {{ color: white !important; }}

  .metric-card {{
    background: {CARD_BG};
    border-radius: 12px;
    padding: 20px 24px;
    border-left: 4px solid {ACCENT};
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    margin-bottom: 4px;
  }}
  .metric-title {{ font-size: 12px; font-weight: 500; color: #888; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }}
  .metric-value {{ font-size: 28px; font-weight: 700; color: {TEXT_DARK}; line-height: 1.2; }}
  .metric-delta {{ font-size: 13px; font-weight: 500; margin-top: 4px; }}
  .delta-pos {{ color: {ACCENT}; }}
  .delta-neg {{ color: {ORANGE}; }}
  .section-header {{
    font-size: 18px; font-weight: 700; color: {TEXT_DARK};
    border-bottom: 2px solid {ACCENT}; padding-bottom: 8px;
    margin: 28px 0 16px 0;
  }}
  .narrative-box {{
    background: white; border-radius: 10px; padding: 20px 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    border-left: 4px solid {PRIMARY};
    font-size: 14px; line-height: 1.8; color: {TEXT_DARK};
  }}
  .stPlotlyChart {{ border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
</style>
""", unsafe_allow_html=True)

# ─── Data loading ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def load_from_gdrive(file_id: str):
    """Download from Google Drive and return decrypted BytesIO."""
    url = f"https://drive.google.com/uc?id={file_id}"
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp_path = tmp.name
    gdown.download(url, tmp_path, quiet=True)
    with open(tmp_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=FILE_PASSWORD)
        buf = io.BytesIO()
        office_file.decrypt(buf)
    os.unlink(tmp_path)
    buf.seek(0)
    return buf


def decrypt_local(uploaded_file) -> io.BytesIO:
    office_file = msoffcrypto.OfficeFile(uploaded_file)
    office_file.load_key(password=FILE_PASSWORD)
    buf = io.BytesIO()
    office_file.decrypt(buf)
    buf.seek(0)
    return buf


@st.cache_data(ttl=3600)
def parse_workbook(file_bytes: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    data = {}

    # ── P&L data ──────────────────────────────────────────────────────────────
    ws = wb["P&L_Data"]
    rows = list(ws.iter_rows(values_only=True))
    header_row = rows[4]   # row 5: Account, Feb 2026, Jan 2026 …
    months = [c for c in header_row[2:14] if c is not None]

    def get_row(label):
        for r in rows:
            if r[0] == label or (len(r) > 1 and r[1] == label):
                return r
        return None

    def row_values(label):
        r = get_row(label)
        if r is None:
            return [0] * len(months)
        vals = list(r[2:2+len(months)])
        return [v if isinstance(v, (int, float)) else 0 for v in vals]

    data["months"] = list(months)
    data["revenue"]  = row_values("Total Income")
    data["cogs"]     = row_values("Total Less Cost of Sales")
    data["gross_profit"] = row_values("Gross Profit")   # the line item
    data["total_opex"]   = row_values("Total Less Operating Expenses")
    data["op_profit"]    = row_values("Operating Profit")
    data["net_profit"]   = row_values("Net Profit")
    data["staff_costs"]  = row_values("Total 0 Staff Costs")
    data["rent_util"]    = row_values("Total 1 Rent, Rates, Utilities")
    data["maintenance"]  = row_values("Total 2 Shop Maintenance")
    data["other_oh"]     = row_values("Total 3 Other overheads")
    data["depreciation"] = row_values("Depreciation")

    # Revenue breakdown
    income_lines = [
        "Coffee and Tea Sales", "Fruit and Veg Sales", "Grocery  Sales",
        "In-House Food Sales", "Turner & George sales",
        "Wine, Beer and Spirits Sales", "Homeware and Gifting Sales",
        "Deli Sales", "Cafe Sales", "Plants and Flowers sales"
    ]
    data["revenue_breakdown"] = {k: row_values(k) for k in income_lines}

    # ── Stats sheet ───────────────────────────────────────────────────────────
    ws2 = wb["Stats"]
    srows = list(ws2.iter_rows(min_row=1, values_only=True))
    stat_months, turnover, gp_pct, overheads, cash = [], [], [], [], []
    for r in srows:
        if r[0] == "Turnover":
            turnover = [v if isinstance(v, (int, float)) else 0 for v in r[1:13]]
        elif r[0] == "Gross profit":
            gp_pct = [v if isinstance(v, (int, float)) else 0 for v in r[1:13]]
        elif r[0] == "Overheads":
            overheads = [v if isinstance(v, (int, float)) else 0 for v in r[1:13]]
        elif r[0] == "Cash balance each month":
            cash = [v if isinstance(v, (int, float)) else 0 for v in r[1:13]]
        elif r[0] is None and isinstance(r[1], str) and r[1].startswith("Mar"):
            stat_months = list(r[1:13])
    data["stat_months"] = stat_months
    data["stat_turnover"]  = turnover
    data["stat_gp_pct"]    = gp_pct
    data["stat_overheads"] = overheads
    data["stat_cash"]      = cash

    # ── Manual / KPIs ─────────────────────────────────────────────────────────
    ws3 = wb["Manual"]
    mrows = list(ws3.iter_rows(values_only=True))
    for r in mrows:
        if len(r) > 3 and r[1] == "Number of Transactions p.w.":
            data["txn_curr"] = r[2]
            data["txn_prev"] = r[3]
        elif len(r) > 3 and r[1] == "Average Spend Per Head (ex. VAT)":
            data["asp_curr"] = r[2]
            data["asp_prev"] = r[3]

    # ── Overview ──────────────────────────────────────────────────────────────
    ws4 = wb["Overview"]
    orows = list(ws4.iter_rows(min_row=5, values_only=True))
    for r in orows:
        lbl = r[0]
        if lbl == "Sales":
            data["ov_sales_curr"] = r[1]; data["ov_sales_prev_mth"] = r[2]
            data["ov_sales_prev_yr"] = r[5]; data["ov_ytd"] = r[8]; data["ov_ytd_ly"] = r[9]
        elif lbl == "Gross profit Total":
            data["ov_gp_curr"] = r[1]; data["ov_gp_prev_mth"] = r[2]
        elif lbl == "Gross Profit":
            data["ov_gp_pct_curr"] = r[1]; data["ov_gp_pct_prev_mth"] = r[2]
        elif lbl == "Staff Costs":
            data["ov_staff_curr"] = r[1]; data["ov_staff_prev_mth"] = r[2]
        elif lbl == "Total Staff / Sales":
            data["ov_staff_pct_curr"] = r[1]; data["ov_staff_pct_prev_mth"] = r[2]

    # ── Narrative text ────────────────────────────────────────────────────────
    ws5 = wb["Text"]
    narrative_lines = []
    for r in ws5.iter_rows(values_only=True):
        for cell in r:
            if cell and isinstance(cell, str) and len(cell.strip()) > 5:
                narrative_lines.append(cell.strip())
    data["narrative"] = narrative_lines

    # ── Period label ──────────────────────────────────────────────────────────
    ws6 = wb["P&L_Data"]
    prows = list(ws6.iter_rows(max_row=5, values_only=True))
    data["period_label"] = prows[2][0] if prows[2][0] else "Management Accounts"

    # ── BS data ───────────────────────────────────────────────────────────────
    ws7 = wb["BS_Data"]
    brows = list(ws7.iter_rows(values_only=True))
    bs_header = brows[3]
    bs_months = [c for c in bs_header[2:14] if c is not None]
    data["bs_months"] = [str(m)[:10] if isinstance(m, datetime) else str(m) for m in bs_months]
    for r in brows:
        if r[0] == "Total Fixed Assets":
            data["bs_fixed"] = [v if isinstance(v, (int, float)) else 0 for v in r[2:14]]
        elif r[0] == "Total Current Assets":
            data["bs_current"] = [v if isinstance(v, (int, float)) else 0 for v in r[2:14]]
        elif r[0] == "Total Current Liabilities":
            data["bs_curr_liab"] = [v if isinstance(v, (int, float)) else 0 for v in r[2:14]]
        elif r[0] == "Net Assets":
            data["bs_net_assets"] = [v if isinstance(v, (int, float)) else 0 for v in r[2:14]]
        elif r[0] == "Cash":
            data["bs_cash_bal"] = [v if isinstance(v, (int, float)) else 0 for v in r[2:14]]

    return data


# ─── Helper functions ─────────────────────────────────────────────────────────
def fmt_k(v):
    if v is None: return "—"
    return f"£{v/1000:,.1f}k"

def fmt_pct(v):
    if v is None: return "—"
    return f"{v*100:.1f}%"

def delta_html(curr, prev, is_pct=False, invert=False):
    """Return coloured delta string."""
    if curr is None or prev is None: return ""
    diff = curr - prev
    positive = diff > 0
    if invert: positive = not positive
    cls = "delta-pos" if positive else "delta-neg"
    sign = "▲" if diff > 0 else "▼"
    if is_pct:
        label = f"{sign} {abs(diff)*100:.1f}pp vs prior month"
    else:
        label = f"{sign} {fmt_k(abs(diff))} vs prior month"
    return f'<span class="{cls}">{label}</span>'


def plotly_theme():
    return dict(
        paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="Inter, sans-serif", size=12, color=TEXT_DARK),
        margin=dict(l=40, r=20, t=40, b=40),
    )


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Gladwells")
    st.markdown("**Management Accounts**")
    st.divider()

    source = st.radio("Data source", ["☁️ Google Drive (auto)", "📁 Upload file"])

    if source == "☁️ Google Drive (auto)":
        gdrive_id = st.text_input(
            "Google Drive File ID",
            value=GDRIVE_FILE_ID,
            help="Paste the ID from the Drive share link: drive.google.com/file/d/**THIS_PART**/view"
        )
        if st.button("🔄 Refresh data", use_container_width=True):
            st.cache_data.clear()
        uploaded = None
    else:
        gdrive_id = ""
        uploaded = st.file_uploader("Upload .xlsm file", type=["xlsm", "xlsx"])

    st.divider()
    st.caption("Password protected file is decrypted automatically.")
    st.caption("Place the new monthly file in the same Google Drive location and click Refresh.")


# ─── Load data ────────────────────────────────────────────────────────────────
buf = None
if source == "☁️ Google Drive (auto)" and gdrive_id:
    try:
        with st.spinner("Downloading from Google Drive…"):
            buf = load_from_gdrive(gdrive_id)
    except Exception as e:
        st.error(f"Could not fetch from Google Drive: {e}")
elif uploaded:
    buf = decrypt_local(uploaded)

if buf is None:
    st.markdown(f"""
    <div style="text-align:center; padding: 60px 20px; color: #888;">
      <div style="font-size:60px">📊</div>
      <h2 style="color:{PRIMARY}">Gladwells Dashboard</h2>
      <p>Connect your Google Drive file or upload the monthly .xlsm file to get started.</p>
      <p><b>Google Drive setup:</b><br>1. Upload your management accounts file to Google Drive<br>
      2. Set sharing to "Anyone with the link can view"<br>
      3. Paste the file ID in the sidebar</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

try:
    data = parse_workbook(buf.read())
except Exception as e:
    st.error(f"Failed to parse workbook: {e}")
    st.stop()

months    = data["months"]
curr_mth  = months[0] if months else "Current"
prev_mth  = months[1] if len(months) > 1 else "Prior"

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="display:flex; align-items:center; gap:16px; margin-bottom:8px;">
  <div style="font-size:36px">📊</div>
  <div>
    <h1 style="margin:0; color:{PRIMARY}; font-size:26px; font-weight:700;">Gladwells Camberwell Limited</h1>
    <p style="margin:0; color:#888; font-size:14px;">{data.get('period_label','')}</p>
  </div>
</div>
""", unsafe_allow_html=True)
st.divider()

# ─── Page tabs ────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["🏠 Overview", "📈 P&L Detail", "💰 Cash & Balance Sheet", "🛒 Revenue Mix"])

# ════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ════════════════════════════════════════════════════════════════════════
with tab1:
    # ── KPI row 1 ──────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    sales_curr = data.get("ov_sales_curr"); sales_prev = data.get("ov_sales_prev_mth")
    gp_curr    = data.get("ov_gp_curr");    gp_prev    = data.get("ov_gp_prev_mth")
    gp_pct_c   = data.get("ov_gp_pct_curr"); gp_pct_p  = data.get("ov_gp_pct_prev_mth")
    staff_c    = data.get("ov_staff_curr"); staff_p    = data.get("ov_staff_prev_mth")
    np_curr    = data["net_profit"][0] if data["net_profit"] else None
    np_prev    = data["net_profit"][1] if len(data["net_profit"]) > 1 else None

    for col, title, curr, prev, is_pct, inv in [
        (c1, "Revenue",     sales_curr, sales_prev, False, False),
        (c2, "Gross Profit",gp_curr,    gp_prev,    False, False),
        (c3, "GP %",        gp_pct_c,   gp_pct_p,   True,  False),
        (c4, "Net Profit",  np_curr,    np_prev,    False, False),
    ]:
        val_str = fmt_pct(curr) if is_pct else fmt_k(curr)
        delta   = delta_html(curr, prev, is_pct=is_pct, invert=inv)
        col.markdown(f"""
        <div class="metric-card">
          <div class="metric-title">{title} — {curr_mth}</div>
          <div class="metric-value">{val_str}</div>
          <div class="metric-delta">{delta}</div>
        </div>""", unsafe_allow_html=True)

    # ── KPI row 2 ──────────────────────────────────────────────────────
    c5, c6, c7, c8 = st.columns(4)
    txn_c  = data.get("txn_curr"); txn_p = data.get("txn_prev")
    asp_c  = data.get("asp_curr"); asp_p = data.get("asp_prev")
    ytd    = data.get("ov_ytd");   ytd_ly = data.get("ov_ytd_ly")
    cash_c = data["stat_cash"][-1] if data.get("stat_cash") else None
    cash_p = data["stat_cash"][-2] if data.get("stat_cash") and len(data["stat_cash"]) > 1 else None

    for col, title, curr, prev, fmt_fn, inv in [
        (c5, "Transactions p.w.", txn_c,  txn_p,  lambda v: f"{v:,.0f}" if v else "—", False),
        (c6, "Avg Spend / Head",  asp_c,  asp_p,  lambda v: f"£{v:.2f}" if v else "—", False),
        (c7, "YTD Revenue",       ytd,    ytd_ly, fmt_k, False),
        (c8, "Cash Balance",      cash_c, cash_p, fmt_k, False),
    ]:
        val_str = fmt_fn(curr)
        if curr is not None and prev is not None:
            diff = curr - prev
            sign = "▲" if diff > 0 else "▼"
            cls  = "delta-pos" if diff > 0 else "delta-neg"
            delta = f'<span class="{cls}">{sign} {abs(diff):,.0f} vs prior</span>' if title == "Transactions p.w." else \
                    f'<span class="{cls}">{sign} £{abs(diff):,.2f} vs prior</span>' if title == "Avg Spend / Head" else \
                    f'<span class="{"delta-pos" if diff > 0 else "delta-neg"}">{sign} {fmt_k(abs(diff))} vs prior</span>'
        else:
            delta = ""
        col.markdown(f"""
        <div class="metric-card">
          <div class="metric-title">{title}</div>
          <div class="metric-value">{val_str}</div>
          <div class="metric-delta">{delta}</div>
        </div>""", unsafe_allow_html=True)

    # ── Narrative ──────────────────────────────────────────────────────
    if data.get("narrative"):
        st.markdown('<div class="section-header">📝 Management Commentary</div>', unsafe_allow_html=True)
        lines = data["narrative"]
        salutation = lines[0] if lines else ""
        body = lines[1:] if len(lines) > 1 else []
        st.markdown(f"""
        <div class="narrative-box">
          <em>{salutation}</em><br><br>
          {"<br>".join(f"• {l}" for l in body if l != salutation)}
        </div>""", unsafe_allow_html=True)

    # ── 12-month trend chart ───────────────────────────────────────────
    st.markdown('<div class="section-header">📅 12-Month Trend</div>', unsafe_allow_html=True)
    sm = data.get("stat_months", [])
    st_turn = data.get("stat_turnover", [])
    st_gp   = [g*100 for g in data.get("stat_gp_pct", [])]
    st_cash = data.get("stat_cash", [])

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=sm, y=st_turn, name="Revenue", marker_color=PRIMARY, opacity=0.85,
        text=[f"£{v/1000:.0f}k" for v in st_turn], textposition="outside", textfont=dict(size=9)
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=sm, y=st_gp, name="GP %", mode="lines+markers",
        line=dict(color=ACCENT, width=2.5), marker=dict(size=7)
    ), secondary_y=True)
    fig.add_trace(go.Scatter(
        x=sm, y=[v/1000 for v in st_cash], name="Cash (£k)", mode="lines+markers",
        line=dict(color=ORANGE, width=2, dash="dot"), marker=dict(size=6)
    ), secondary_y=False)
    fig.update_layout(
        **plotly_theme(), legend=dict(orientation="h", y=1.12, x=0),
        yaxis=dict(title="£ (000s)", tickformat=",.0f"),
        yaxis2=dict(title="GP %", tickformat=".0f", ticksuffix="%"),
        barmode="group",
    )
    st.plotly_chart(fig, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════
# TAB 2 — P&L DETAIL
# ════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">Profit & Loss — Monthly Waterfall</div>', unsafe_allow_html=True)

    # Waterfall for current month
    rev = data["revenue"][0] if data["revenue"] else 0
    cgs = -abs(data["cogs"][0]) if data["cogs"] else 0
    gp  = data["gross_profit"][0] if data["gross_profit"] else 0
    stf = -abs(data["staff_costs"][0]) if data["staff_costs"] else 0
    rnt = -abs(data["rent_util"][0]) if data["rent_util"] else 0
    mnt = -abs(data["maintenance"][0]) if data["maintenance"] else 0
    ooh = -abs(data["other_oh"][0]) if data["other_oh"] else 0
    dep = -abs(data["depreciation"][0]) if data["depreciation"] else 0
    np  = data["net_profit"][0] if data["net_profit"] else 0

    wf_labels  = ["Revenue","Cost of Sales","Gross Profit","Staff","Rent & Utilities","Maintenance","Other Overheads","Depreciation","Net Profit"]
    wf_values  = [rev, cgs, gp, stf, rnt, mnt, ooh, dep, np]
    wf_measure = ["absolute","relative","total","relative","relative","relative","relative","relative","total"]
    wf_colours = [PRIMARY, ORANGE, ACCENT, ORANGE, ORANGE, ORANGE, ORANGE, ORANGE, PRIMARY if np >= 0 else ORANGE]

    fig2 = go.Figure(go.Waterfall(
        orientation="v", measure=wf_measure,
        x=wf_labels, y=wf_values,
        connector=dict(line=dict(color="#ccc", width=1)),
        increasing=dict(marker_color=PRIMARY),
        decreasing=dict(marker_color=ORANGE),
        totals=dict(marker_color=ACCENT),
        text=[fmt_k(abs(v)) for v in wf_values],
        textposition="outside",
    ))
    fig2.update_layout(**plotly_theme(), yaxis=dict(tickformat=",.0f", title="£"), showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

    # ── Monthly P&L table ──────────────────────────────────────────────
    st.markdown('<div class="section-header">Monthly P&L Summary (£000s)</div>', unsafe_allow_html=True)
    n = min(len(months), 13)
    pnl_rows = {
        "Revenue": data["revenue"][:n],
        "Cost of Sales": [-abs(v) for v in data["cogs"][:n]],
        "Gross Profit": data["gross_profit"][:n],
        "GP %": [f"{g/r*100:.1f}%" if r else "—" for g, r in zip(data["gross_profit"][:n], data["revenue"][:n])],
        "Staff Costs": [-abs(v) for v in data["staff_costs"][:n]],
        "Rent & Utilities": [-abs(v) for v in data["rent_util"][:n]],
        "Maintenance": [-abs(v) for v in data["maintenance"][:n]],
        "Other Overheads": [-abs(v) for v in data["other_oh"][:n]],
        "Total Overheads": [-abs(v) for v in data["total_opex"][:n]],
        "Operating Profit": data["op_profit"][:n],
        "Depreciation": [-abs(v) for v in data["depreciation"][:n]],
        "Net Profit": data["net_profit"][:n],
    }
    col_headers = months[:n]
    table_data = {}
    for k, vals in pnl_rows.items():
        row_vals = []
        for v in vals:
            if isinstance(v, str): row_vals.append(v)
            elif isinstance(v, float) or isinstance(v, int): row_vals.append(f"£{v/1000:,.1f}k")
            else: row_vals.append("—")
        table_data[k] = row_vals

    df_pnl = pd.DataFrame(table_data, index=col_headers).T
    st.dataframe(df_pnl, use_container_width=True, height=430)

    # ── Overhead breakdown chart ────────────────────────────────────────
    st.markdown('<div class="section-header">Overhead Breakdown — Last 12 Months</div>', unsafe_allow_html=True)
    sm2 = list(reversed(months[:12]))
    fig3 = go.Figure()
    oh_items = {
        "Staff": list(reversed(data["staff_costs"][:12])),
        "Rent & Utilities": list(reversed(data["rent_util"][:12])),
        "Maintenance": list(reversed(data["maintenance"][:12])),
        "Other": list(reversed(data["other_oh"][:12])),
    }
    colours3 = [PRIMARY, ACCENT, ORANGE, "#B5D4A8"]
    for (lbl, vals), col in zip(oh_items.items(), colours3):
        fig3.add_trace(go.Bar(x=sm2, y=vals, name=lbl, marker_color=col))
    fig3.update_layout(**plotly_theme(), barmode="stack",
                       yaxis=dict(title="£", tickformat=",.0f"),
                       legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig3, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════
# TAB 3 — CASH & BALANCE SHEET
# ════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">Cash Position — 12 Months</div>', unsafe_allow_html=True)

    sm3 = data.get("stat_months", [])
    cash_vals = data.get("stat_cash", [])
    cf_vals   = data.get("stat_monthly_cf", [])
    mcf = [data["stat_cash"][i] - data["stat_cash"][i-1] if i > 0 else 0 for i in range(len(data.get("stat_cash", [])))]

    fig4 = make_subplots(specs=[[{"secondary_y": True}]])
    fig4.add_trace(go.Bar(
        x=sm3, y=mcf, name="Monthly Cash Flow",
        marker_color=[PRIMARY if v >= 0 else ORANGE for v in mcf],
        opacity=0.75,
    ), secondary_y=False)
    fig4.add_trace(go.Scatter(
        x=sm3, y=cash_vals, name="Cash Balance", mode="lines+markers",
        line=dict(color=ACCENT, width=3), marker=dict(size=8, color=ACCENT),
        fill="tozeroy", fillcolor="rgba(151,188,98,0.12)"
    ), secondary_y=True)
    fig4.update_layout(
        **plotly_theme(),
        yaxis=dict(title="Monthly Cash Flow (£)", tickformat=",.0f"),
        yaxis2=dict(title="Cash Balance (£)", tickformat=",.0f"),
        legend=dict(orientation="h", y=1.1)
    )
    st.plotly_chart(fig4, use_container_width=True)

    # ── KPIs ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Balance Sheet Snapshot</div>', unsafe_allow_html=True)
    bca = data.get("bs_cash_bal", [None])[0]
    bca_p = data.get("bs_cash_bal", [None, None])[1] if len(data.get("bs_cash_bal", [])) > 1 else None
    bfa = data.get("bs_fixed", [None])[0]
    bcu = data.get("bs_current", [None])[0]
    bcl = data.get("bs_curr_liab", [None])[0]
    bna = data.get("bs_net_assets", [None])[0]

    cc1, cc2, cc3, cc4 = st.columns(4)
    for col, title, curr, prev in [
        (cc1, "Cash & Bank", bca, bca_p),
        (cc2, "Net Assets",  bna, None),
        (cc3, "Current Assets", bcu, None),
        (cc4, "Current Liabilities", bcl, None),
    ]:
        val_str = fmt_k(curr)
        delta = delta_html(curr, prev) if prev is not None else ""
        col.markdown(f"""
        <div class="metric-card">
          <div class="metric-title">{title}</div>
          <div class="metric-value">{val_str}</div>
          <div class="metric-delta">{delta}</div>
        </div>""", unsafe_allow_html=True)

    # ── Balance sheet trend ────────────────────────────────────────────
    bs_m = data.get("bs_months", [])
    bs_ca = data.get("bs_current", [])
    bs_cl = [abs(v) for v in data.get("bs_curr_liab", [])]
    bs_na = data.get("bs_net_assets", [])
    bs_cash = data.get("bs_cash_bal", [])

    if bs_m:
        n_bs = min(len(bs_m), len(bs_ca), len(bs_cl), len(bs_na))
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(x=bs_m[:n_bs], y=bs_ca[:n_bs], name="Current Assets", mode="lines+markers", line=dict(color=PRIMARY, width=2)))
        fig5.add_trace(go.Scatter(x=bs_m[:n_bs], y=bs_cl[:n_bs], name="Current Liabilities", mode="lines+markers", line=dict(color=ORANGE, width=2, dash="dash")))
        fig5.add_trace(go.Scatter(x=bs_m[:n_bs], y=bs_cash[:n_bs], name="Cash", mode="lines+markers", line=dict(color=ACCENT, width=2)))
        fig5.add_trace(go.Scatter(x=bs_m[:n_bs], y=bs_na[:n_bs], name="Net Assets", mode="lines+markers", line=dict(color="#7A5C58", width=2, dash="dot")))
        fig5.update_layout(**plotly_theme(), yaxis=dict(title="£", tickformat=",.0f"),
                           legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig5, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════
# TAB 4 — REVENUE MIX
# ════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">Revenue Mix — Current Month</div>', unsafe_allow_html=True)

    rb = data.get("revenue_breakdown", {})
    labels_pie = [k for k, v in rb.items() if v and v[0] and v[0] > 0]
    vals_pie   = [rb[k][0] for k in labels_pie]

    colour_seq = [PRIMARY, ACCENT, ORANGE, "#B5D4A8", "#7A5C58", "#D4A853", "#5C8A7A", "#C47A5A", "#6B8CBF", "#A87A5C"]

    col_a, col_b = st.columns([1, 1])
    with col_a:
        fig6 = go.Figure(go.Pie(
            labels=labels_pie, values=vals_pie,
            hole=0.5, marker=dict(colors=colour_seq),
            textinfo="label+percent", textfont=dict(size=11),
        ))
        fig6.add_annotation(text=f"<b>Total</b><br>{fmt_k(sum(vals_pie))}", x=0.5, y=0.5,
                            font_size=13, showarrow=False)
        theme6 = plotly_theme()
        theme6["margin"] = dict(l=20, r=20, t=20, b=20)
        fig6.update_layout(**theme6, showlegend=False)
        st.plotly_chart(fig6, use_container_width=True)

    with col_b:
        df_rb = pd.DataFrame({
            "Category": labels_pie,
            f"{curr_mth}": [fmt_k(rb[k][0]) for k in labels_pie],
            f"{prev_mth}": [fmt_k(rb[k][1]) if len(rb[k]) > 1 else "—" for k in labels_pie],
            "Change": [
                f"{'▲' if rb[k][0] > rb[k][1] else '▼'} {fmt_k(abs(rb[k][0]-rb[k][1]))}"
                if len(rb[k]) > 1 and rb[k][1] else "—"
                for k in labels_pie
            ]
        })
        st.dataframe(df_rb, use_container_width=True, hide_index=True, height=340)

    # ── Revenue trend by category ──────────────────────────────────────
    st.markdown('<div class="section-header">Revenue Trend by Category — Last 12 Months</div>', unsafe_allow_html=True)
    top_cats = sorted(labels_pie, key=lambda k: rb[k][0], reverse=True)[:5]
    sm4 = list(reversed(months[:12]))
    fig7 = go.Figure()
    for cat, col in zip(top_cats, colour_seq):
        vals = list(reversed(rb[cat][:12]))
        fig7.add_trace(go.Bar(x=sm4, y=vals, name=cat, marker_color=col))
    fig7.update_layout(**plotly_theme(), barmode="stack",
                       yaxis=dict(title="£", tickformat=",.0f"),
                       legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig7, use_container_width=True)

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center; color:#bbb; font-size:12px; margin-top:40px; padding:20px;">
  Gladwells Management Accounts Dashboard · Auto-refreshes from Google Drive every hour<br>
  Last loaded: {datetime.now().strftime("%d %b %Y %H:%M")}
</div>""", unsafe_allow_html=True)
