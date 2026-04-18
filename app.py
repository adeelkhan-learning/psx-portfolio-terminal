"""
╔══════════════════════════════════════════════════════════════════╗
║          PSX PORTFOLIO DASHBOARD  –  app.py                      ║
║  Pakistan Stock Exchange · Personal Investment Tracker           ║
╚══════════════════════════════════════════════════════════════════╝

Run:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PSX Portfolio Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── GLOBAL STYLES ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Fonts ── */
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Sora:wght@300;400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Sora', sans-serif; }

/* ── Background ── */
.stApp { background: #0d0f14; color: #e2e8f0; }
[data-testid="stSidebar"] { background: #111318 !important; border-right: 1px solid #1e2130; }

/* ── Metric Cards ── */
.kpi-card {
    background: linear-gradient(135deg,#151820 0%,#1a1d26 100%);
    border: 1px solid #252836;
    border-radius: 12px;
    padding: 20px 22px;
    position: relative;
    overflow: hidden;
}
.kpi-card::before {
    content:'';
    position:absolute;top:0;left:0;right:0;height:3px;
    background: var(--accent);
    border-radius:3px 3px 0 0;
}
.kpi-label  { font-size:11px; letter-spacing:1.5px; text-transform:uppercase; color:#6b7280; margin-bottom:6px; }
.kpi-value  { font-family:'IBM Plex Mono',monospace; font-size:22px; font-weight:600; color:#f1f5f9; }
.kpi-delta  { font-size:12px; margin-top:4px; }
.kpi-pos    { color:#22c55e; }
.kpi-neg    { color:#ef4444; }
.kpi-neu    { color:#94a3b8; }

/* ── Ticker bar ── */
.ticker-wrap {
    background:#111318;
    border: 1px solid #1e2130;
    border-radius: 10px;
    padding: 10px 0;
    overflow: hidden;
    white-space: nowrap;
    margin-bottom: 18px;
}
.ticker-inner {
    display:inline-block;
    animation: scroll-left 60s linear infinite;
}
.ticker-inner:hover { animation-play-state: paused; }
.ticker-item {
    display:inline-block;
    margin: 0 28px;
    font-family:'IBM Plex Mono',monospace;
    font-size:13px;
}
.t-sym  { color:#94a3b8; font-weight:600; margin-right:6px; }
.t-price{ color:#f1f5f9; }
.t-pos  { color:#22c55e; }
.t-neg  { color:#ef4444; }
.sep    { color:#2d3148; margin: 0 8px; }
@keyframes scroll-left {
    0%   { transform: translateX(0); }
    100% { transform: translateX(-50%); }
}

/* ── Section titles ── */
.sec-title {
    font-size:13px; letter-spacing:2px; text-transform:uppercase;
    color:#475569; margin:28px 0 14px; border-bottom:1px solid #1e2130;
    padding-bottom:8px;
}

/* ── Table ── */
.stock-table { font-size:13px; }

/* ── Sidebar labels ── */
.sidebar-label { font-size:11px; letter-spacing:1px; color:#6b7280; text-transform:uppercase; margin-bottom:4px; }

/* ── Plotly chart bg ── */
.js-plotly-plot { border-radius:10px; }

/* ── Hide streamlit default header ── */
#MainMenu, footer, header { visibility:hidden; }

/* ── Best/Worst badges ── */
.badge {
    display:inline-block; padding:3px 10px; border-radius:20px;
    font-size:11px; font-weight:600; letter-spacing:1px;
}
.badge-green { background:#14532d; color:#4ade80; }
.badge-red   { background:#450a0a; color:#f87171; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def load_data(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    trades    = xl.parse("Trades")
    dividends = xl.parse("Dividends")
    funds     = xl.parse("Funds")

    # ── normalise dates ──
    for df, col in [(trades,"Trade Date"),(dividends,"Payment Date"),(funds,"Date")]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # ── normalise numerics ──
    num_cols = {
        "trades":    ["Quantity","Price","Commission","Taxes and Fees","Net Total Value"],
        "dividends": ["Net Amount Paid","Gross Dividend","Zakat Deducted","Tax Deducted"],
        "funds":     ["Amount Deposit","Amount Hold Against Charges / Dues",
                      "Amount Transferred To Exposure"],
    }
    for col in num_cols["trades"]:
        trades[col] = pd.to_numeric(trades[col], errors="coerce").fillna(0)
    for col in num_cols["dividends"]:
        dividends[col] = pd.to_numeric(dividends[col], errors="coerce").fillna(0)
    for col in num_cols["funds"]:
        funds[col] = pd.to_numeric(funds[col], errors="coerce").fillna(0)

    # ── tickers ──
    trades["Ticker"]    = trades["Ticker"].str.strip().str.upper()
    dividends["Ticker"] = dividends["Ticker"].str.strip().str.upper()

    return trades, dividends, funds


def compute_holdings(trades):
    """
    FIFO-based holdings: returns a dict ticker → {qty, avg_price, total_invested}
    """
    holdings = {}
    for _, row in trades.sort_values("Trade Date").iterrows():
        t = row["Ticker"]
        qty   = row["Quantity"]
        price = row["Price"]
        txn   = str(row["Transaction Type"]).strip().lower()

        if t not in holdings:
            holdings[t] = {"qty": 0, "cost": 0.0}

        if txn in ("buy", "ipo", "allotment", "ipo allotment"):   # IPO = shares allotted, treat like buy
            holdings[t]["qty"]  += qty
            holdings[t]["cost"] += qty * price
        elif txn == "sell":
            if holdings[t]["qty"] > 0:
                avg = holdings[t]["cost"] / holdings[t]["qty"] if holdings[t]["qty"] else 0
                sell_qty = min(qty, holdings[t]["qty"])
                holdings[t]["qty"]  -= sell_qty
                holdings[t]["cost"] -= sell_qty * avg

    result = {}
    for t, d in holdings.items():
        if d["qty"] > 0:
            avg = d["cost"] / d["qty"] if d["qty"] else 0
            result[t] = {
                "qty":           d["qty"],
                "avg_price":     round(avg, 2),
                "total_invested": round(d["cost"], 2),
            }
    return result


# ── Known ETF tickers (use /etf/ URL path instead of /company/) ────────────
_PSX_ETF_TICKERS = {
    "MIIETF", "MZNPETF", "OCTOPUS", "NAFA", "UBL-ETFS",
    "ABL-ETFS", "HBL-ETFS", "MCB-ETFS", "ALFAETF",
    "PAKETF", "KAFETF", "NIFTYETF",
}


def _is_etf(ticker: str) -> bool:
    """Return True if ticker is a known ETF/fund on PSX."""
    return ticker.upper() in _PSX_ETF_TICKERS


def fetch_psx_prices_selenium(tickers: list, debug_log: dict) -> dict:
    """
    Scrape current market price for each ticker from dps.psx.com.pk
    using a single headless Chrome session (fast — one browser, all tickers).

    URL patterns:
      Regular stock : https://dps.psx.com.pk/company/{TICKER}
      ETF / Fund    : https://dps.psx.com.pk/etf/{TICKER}

    Price element : first div whose class contains 'price'
                    Text is multi-line; first line = the current price.
                    Format: "Rs. 1,234.56"  or  "1234.56"

    Returns dict { ticker: float_or_None }
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError:
        debug_log["__error__"] = "selenium not installed. Run: pip install selenium"
        return {t: None for t in tickers}

    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1280,800")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )

    prices = {}
    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        debug_log["__error__"] = f"Chrome launch failed: {e}"
        return {t: None for t in tickers}

    try:
        for ticker in tickers:
            # Decide URL: ETF path vs company path
            # Also auto-fallback: if company page returns nothing → try etf path
            paths_to_try = (
                ["etf", "company"] if _is_etf(ticker) else ["company", "etf"]
            )

            found = False
            for path in paths_to_try:
                url = f"https://dps.psx.com.pk/{path}/{ticker}"
                try:
                    driver.get(url)
                    # Wait up to 10 s for any element with 'price' in its class
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//*[contains(@class,'price')]")
                        )
                    )
                    raw_text = element.text.strip()
                    # First non-empty line is the price
                    first_line = next(
                        (ln.strip() for ln in raw_text.splitlines() if ln.strip()), ""
                    )
                    # Strip currency prefix and thousands separators
                    clean = (
                        first_line
                        .replace("Rs.", "").replace("Rs", "")
                        .replace(",", "").strip()
                    )
                    price = float(clean)
                    if price > 0:
                        prices[ticker]    = round(price, 2)
                        debug_log[ticker] = f"✅ [{path}] {price}"
                        found = True
                        break
                    else:
                        debug_log[ticker] = f"⚠️ [{path}] parsed 0 from '{raw_text[:60]}'"
                except Exception as e:
                    debug_log[ticker] = f"❌ [{path}] {type(e).__name__}: {str(e)[:80]}"

            if not found:
                prices[ticker] = None

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    return prices


def get_current_prices(all_tickers: list, holdings_dict: dict, force_refresh=False):
    """
    Main price resolver. Uses session_state cache so Streamlit re-renders
    don't re-launch Chrome on every interaction.

    Cache key: "_cached_prices" and "_cached_debug"
    Force refresh: clears cache and re-scrapes.
    """
    cache_key  = "_cached_prices"
    debug_key  = "_cached_debug"
    tickers_key = "_cached_tickers"

    # Invalidate cache if ticker set changed or force_refresh requested
    cached_tickers = st.session_state.get(tickers_key, [])
    if (force_refresh
            or cache_key not in st.session_state
            or sorted(cached_tickers) != sorted(all_tickers)):
        debug_log = {}
        prices    = fetch_psx_prices_selenium(all_tickers, debug_log)
        st.session_state[cache_key]   = prices
        st.session_state[debug_key]   = debug_log
        st.session_state[tickers_key] = all_tickers

    raw_prices = st.session_state[cache_key]
    debug_log  = st.session_state.get(debug_key, {})

    # Build final prices: scraped > avg_buy fallback
    current_prices = {}
    price_source   = {}
    for t in holdings_dict:
        p = raw_prices.get(t)
        if p and p > 0:
            current_prices[t] = p
            price_source[t]   = "live"
        else:
            current_prices[t] = holdings_dict[t]["avg_price"]
            price_source[t]   = "avg"

    return current_prices, price_source, debug_log




def fmt_pkr(val):
    if abs(val) >= 1_000_000:
        return f"PKR {val/1_000_000:.2f}M"
    elif abs(val) >= 1_000:
        return f"PKR {val/1_000:.1f}K"
    return f"PKR {val:,.0f}"


def kpi_card(label, value, delta=None, accent="#3b82f6"):
    delta_html = ""
    if delta is not None:
        cls = "kpi-pos" if delta >= 0 else "kpi-neg"
        arrow = "▲" if delta >= 0 else "▼"
        delta_html = f'<div class="kpi-delta {cls}">{arrow} {abs(delta):.1f}%</div>'
    return f"""
    <div class="kpi-card" style="--accent:{accent}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {delta_html}
    </div>
    """


CHART_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Sora, sans-serif", color="#94a3b8", size=12),
    margin=dict(l=10, r=10, t=30, b=10),
    legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor="#252836"),
    xaxis=dict(gridcolor="#1e2130", zerolinecolor="#252836"),
    yaxis=dict(gridcolor="#1e2130", zerolinecolor="#252836"),
)


# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ════════════════════════════════════════════════════════════════════════════

def main():

    # ── Suppress use_container_width deprecation: set via config instead ──────
    import warnings
    warnings.filterwarnings("ignore", message=".*use_container_width.*")

    def _dl_button(df: "pd.DataFrame", filename: str, label="⬇ Download as Excel"):
        """Render an inline Excel download button for any DataFrame."""
        import io
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        st.download_button(
            label=label,
            data=buf.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )

    # ── Branding ──────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="display:flex;align-items:center;gap:14px;margin-bottom:6px;">
        <span style="font-size:28px;">📈</span>
        <div>
            <div style="font-size:20px;font-weight:700;color:#f1f5f9;letter-spacing:-0.5px;">
                PSX Portfolio Dashboard
            </div>
            <div style="font-size:12px;color:#475569;letter-spacing:1px;">
                PAKISTAN STOCK EXCHANGE · PERSONAL INVESTMENT TRACKER
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── File Upload ───────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown('<div class="sidebar-label">📂 Upload Your Excel File</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader(
            label="upload",
            type=["xlsx"],
            label_visibility="collapsed",
            help="Must contain sheets: Trades, Dividends, Funds"
        )
        st.markdown("---")
        st.markdown('<div class="sidebar-label">⚙️ Filters</div>', unsafe_allow_html=True)

    if not uploaded:
        st.info("👆 **Upload your Excel file** in the sidebar to get started.  \n"
                "Don't have one? Run `generate_sample_data.py` to create sample data.")
        _show_instructions()
        return

    # ── Load & process ────────────────────────────────────────────────────────
    with st.spinner("Loading your portfolio…"):
        trades, dividends, funds = load_data(uploaded)

    holdings_dict    = compute_holdings(trades)

    all_tickers = sorted(holdings_dict.keys())

    # ── Sidebar filters ───────────────────────────────────────────────────────
    with st.sidebar:
        selected_ticker = st.selectbox(
            "Select a Stock",
            ["All Stocks"] + all_tickers,
            help="Drill into a specific company"
        )

        date_min = trades["Trade Date"].min().date()
        date_max = trades["Trade Date"].max().date()
        date_range = st.date_input(
            "Date Range",
            value=(date_min, date_max),
            min_value=date_min,
            max_value=date_max,
        )

        st.markdown("---")
        st.markdown(
            '<div style="font-size:11px;color:#4b5563;letter-spacing:1px;'
            'text-transform:uppercase;margin-bottom:6px;">💰 Dividend Settings</div>',
            unsafe_allow_html=True
        )
        div_to_broker = st.toggle(
            "Dividends go to broker account",
            value=False,
            help=(
                "ON  → dividends stay in your brokerage account (counted in cash balance)\n"
                "OFF → dividends go to your bank account (shown separately, not in broker cash)"
            ),
        )
        st.markdown("---")
        st.markdown(
            '<div style="font-size:11px;color:#4b5563;letter-spacing:1px;'
            'text-transform:uppercase;margin-bottom:6px;">📡 Live Prices</div>',
            unsafe_allow_html=True
        )
        force_refresh = st.button(
            "🔄 Refresh Prices",
            help="Re-scrape current prices from PSX website",
            use_container_width=True,
        )

    # ── Fetch / cache prices via Selenium scraper ─────────────────────────────
    first_load = "_cached_prices" not in st.session_state
    if first_load or force_refresh:
        with st.spinner("📡 Fetching live prices from PSX website… (30-60 s first load)"):
            current_prices, price_source, debug_log = get_current_prices(
                all_tickers, holdings_dict, force_refresh=True
            )
    else:
        current_prices, price_source, debug_log = get_current_prices(
            all_tickers, holdings_dict, force_refresh=False
        )

    # ── Sidebar: price status panel ───────────────────────────────────────────
    with st.sidebar:
        live_count  = sum(1 for s in price_source.values() if s == "live")
        avg_count   = sum(1 for s in price_source.values() if s == "avg")
        status_col  = "#22c55e" if avg_count == 0 else "#f59e0b"
        st.markdown(
            f'<div style="font-size:11px;color:{status_col};margin-bottom:4px;">'
            f'✅ {live_count} live &nbsp;|&nbsp; '
            f'{"⚠️ " + str(avg_count) + " using avg buy price" if avg_count else ""}</div>',
            unsafe_allow_html=True
        )
        if avg_count:
            st.caption(
                "Tickers using avg price: "
                + ", ".join(t for t, s in price_source.items() if s == "avg")
            )

        with st.expander("🔍 Price fetch log", expanded=False):
            for t, msg in debug_log.items():
                icon = "✅" if "✅" in msg else "⚠️" if "⚠️" in msg else "❌"
                st.markdown(
                    f'<div style="font-size:10px;color:#6b7280;'
                    f'word-break:break-all;margin-bottom:3px;">'
                    f'<b>{t}</b>: {msg[:140]}</div>',
                    unsafe_allow_html=True
                )

    dr_start = pd.Timestamp(date_range[0]) if len(date_range) == 2 else pd.Timestamp(date_min)
    dr_end   = pd.Timestamp(date_range[1]) if len(date_range) == 2 else pd.Timestamp(date_max)

    trades_f = trades[(trades["Trade Date"] >= dr_start) & (trades["Trade Date"] <= dr_end)]
    divs_f   = dividends[(dividends["Payment Date"] >= dr_start) & (dividends["Payment Date"] <= dr_end)]

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 1 — LIVE TICKER BAR
    # ════════════════════════════════════════════════════════════════════════
    ticker_items = []
    for t in all_tickers:
        h   = holdings_dict[t]
        cur = current_prices.get(t, h["avg_price"])
        pct = (cur - h["avg_price"]) / h["avg_price"] * 100 if h["avg_price"] else 0
        cls  = "t-pos" if pct >= 0 else "t-neg"
        sign = "▲" if pct >= 0 else "▼"
        ticker_items.append(
            f'<span class="ticker-item">'
            f'<span class="t-sym">{t}</span>'
            f'<span class="t-price">₨{cur:,.2f}</span>'
            f'<span class="sep">|</span>'
            f'<span class="t-price">avg ₨{h["avg_price"]:,.2f}</span>'
            f'<span class="sep">·</span>'
            f'<span class="{cls}">{sign}{abs(pct):.1f}%</span>'
            f'</span>'
            f'<span class="sep">◆</span>'
        )
    double_items = "".join(ticker_items * 2)   # duplicate for seamless loop
    st.markdown(f"""
    <div class="ticker-wrap">
        <div class="ticker-inner">{double_items}</div>
    </div>
    """, unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 2 — KPI SUMMARY CARDS
    # ════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">📊 Portfolio at a Glance</div>', unsafe_allow_html=True)

    total_invested    = sum(v["total_invested"] for v in holdings_dict.values())
    total_curr_value  = sum(
        holdings_dict[t]["qty"] * current_prices.get(t, holdings_dict[t]["avg_price"])
        for t in holdings_dict
    )
    total_pnl         = total_curr_value - total_invested
    total_pnl_pct     = (total_pnl / total_invested * 100) if total_invested else 0
    total_dividends   = dividends["Net Amount Paid"].sum()
    total_deposited   = funds[funds["Transfer Type"].str.strip().str.lower().str.contains("deposit|credit", na=False)]["Amount Deposit"].sum()

    c1, c2, c3, c4, c5 = st.columns(5)
    cards = [
        (c1, "Money I Invested",       fmt_pkr(total_invested),   None,          "#3b82f6"),
        (c2, "What It's Worth Today",  fmt_pkr(total_curr_value), None,          "#8b5cf6"),
        (c3, "My Profit / Loss",       fmt_pkr(total_pnl),        total_pnl_pct, "#22c55e" if total_pnl >= 0 else "#ef4444"),
        (c4, "Dividends Received",     fmt_pkr(total_dividends),  None,          "#f59e0b"),
        (c5, "Total Cash Added",       fmt_pkr(total_deposited),  None,          "#06b6d4"),
    ]
    for col, label, value, delta, accent in cards:
        with col:
            st.markdown(kpi_card(label, value, delta, accent), unsafe_allow_html=True)

    st.markdown("")

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 3 — STOCK-WISE TABLE
    # ════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">📋 Stock-by-Stock Breakdown</div>', unsafe_allow_html=True)

    rows = []
    for t in all_tickers:
        h   = holdings_dict[t]
        cur = current_prices.get(t, h["avg_price"])
        pnl = (cur - h["avg_price"]) * h["qty"]
        pct = (cur - h["avg_price"]) / h["avg_price"] * 100 if h["avg_price"] else 0
        divs = dividends[dividends["Ticker"] == t]["Net Amount Paid"].sum()
        rows.append({
            "Ticker":            t,
            "Qty Held":          int(h["qty"]),
            "Avg Buy Price ₨":   h["avg_price"],
            "Current Price ₨":   cur,
            "Money In ₨":        round(h["total_invested"], 0),
            "Worth Now ₨":       round(h["qty"] * cur, 0),
            "Profit/Loss ₨":     round(pnl, 0),
            "Gain %":            round(pct, 2),
            "Dividends ₨":       round(divs, 0),
        })

    df_table = pd.DataFrame(rows)

    # Best / Worst badges — single flex row, never wraps
    if not df_table.empty:
        best  = df_table.loc[df_table["Gain %"].idxmax(), "Ticker"]
        worst = df_table.loc[df_table["Gain %"].idxmin(), "Ticker"]
        best_pct  = df_table.loc[df_table["Gain %"].idxmax(), "Gain %"]
        worst_pct = df_table.loc[df_table["Gain %"].idxmin(), "Gain %"]
        st.markdown(
            f'<div style="display:flex;gap:12px;align-items:center;margin-bottom:10px;">' +
            f'<span style="white-space:nowrap;">🏆 <span class="badge badge-green">BEST: {best} (+{best_pct:.1f}%)</span></span>' +
            f'<span style="white-space:nowrap;">💸 <span class="badge badge-red">WORST: {worst} ({worst_pct:.1f}%)</span></span>' +
            f'</div>',
            unsafe_allow_html=True
        )
        b2_dummy = None
        if False:
            with b2_dummy:
                pass   # badge rendered above
        st.markdown("")

    # Render as a custom HTML table so we can color profit/loss cells red or green
    def _build_stock_table(df, ps):
        """Build an HTML table with color-coded P&L columns."""
        src_icon = {"live": "📡", "manual": "✏️", "avg": "⚠️"}
        headers = [
            "Ticker", "Qty Held", "Avg Buy ₨", "Price Today ₨",
            "Money In ₨", "Worth Now ₨", "Profit / Loss ₨", "Gain %", "Dividends ₨"
        ]
        hdr_html = "".join(f"<th>{h}</th>" for h in headers)
        rows_html = ""
        for _, row in df.iterrows():
            t   = row["Ticker"]
            pnl = row["Profit/Loss ₨"]
            pct = row["Gain %"]
            is_pos = pnl >= 0
            pnl_color = "#22c55e" if is_pos else "#ef4444"
            pct_str = f"+{pct:.2f}%" if is_pos else f"{pct:.2f}%"
            pnl_str = f"+{pnl:,.0f}" if is_pos else f"{pnl:,.0f}"
            src_lbl = src_icon.get(ps.get(t, "avg"), "")
            rows_html += f"""<tr>
                <td style="font-weight:600;color:#f1f5f9;">{t} <span style="font-size:10px">{src_lbl}</span></td>
                <td>{int(row["Qty Held"]):,}</td>
                <td>{row["Avg Buy Price ₨"]:,.2f}</td>
                <td>{row["Current Price ₨"]:,.2f}</td>
                <td>{row["Money In ₨"]:,.0f}</td>
                <td>{row["Worth Now ₨"]:,.0f}</td>
                <td style="color:{pnl_color};font-weight:600;">{pnl_str}</td>
                <td style="color:{pnl_color};font-weight:600;">{pct_str}</td>
                <td style="color:#f59e0b;">{row["Dividends ₨"]:,.0f}</td>
            </tr>"""
        return f"""
        <div style="overflow-x:auto;max-height:360px;overflow-y:auto;">
        <table style="width:100%;border-collapse:collapse;
                      font-family:'IBM Plex Mono',monospace;font-size:12px;color:#94a3b8;">
          <thead style="position:sticky;top:0;background:#111318;z-index:1;">
            <tr style="border-bottom:1px solid #252836;">{hdr_html}</tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
        </div>
        <style>
          table tr:hover td {{ background:rgba(255,255,255,0.03); }}
          table th, table td {{ padding:9px 12px; text-align:right; border-bottom:1px solid #1e2130; }}
          table th {{ color:#6b7280; font-size:10px; letter-spacing:1px;
                      text-transform:uppercase; font-weight:500; text-align:right; }}
          table td:first-child, table th:first-child {{ text-align:left; }}
        </style>"""

    st.markdown(_build_stock_table(df_table, price_source), unsafe_allow_html=True)
    _dl_col, _legend_col = st.columns([1, 4])
    with _dl_col:
        _dl_button(df_table, "stock_breakdown.xlsx", "⬇ Download Table")
    with _legend_col:
        st.markdown(
            '<div style="font-size:10px;color:#374151;margin-top:8px;">' +
            '📡 Live price &nbsp; ✏️ Manual &nbsp; ⚠️ Using avg buy price (check PSX website)</div>',
            unsafe_allow_html=True
        )

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 4 — CHARTS ROW (Portfolio Growth + Allocation)
    # ════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">📈 Portfolio Growth Over Time</div>', unsafe_allow_html=True)

    col_left, col_right = st.columns([3, 2])

    with col_left:
        # Build running portfolio value timeline
        buys = trades_f[trades_f["Transaction Type"].str.lower() == "buy"].copy()
        buys["Cum Invested"] = (buys["Price"] * buys["Quantity"]).cumsum()
        buys = buys.rename(columns={"Trade Date": "Date"})

        if not buys.empty:
            fig_growth = go.Figure()
            fig_growth.add_trace(go.Scatter(
                x=buys["Date"],
                y=buys["Cum Invested"],
                mode="lines",
                name="Total Invested",
                line=dict(color="#3b82f6", width=2),
                fill="tozeroy",
                fillcolor="rgba(59,130,246,0.08)",
                hovertemplate="<b>%{x|%b %d, %Y}</b><br>Invested: ₨%{y:,.0f}<extra></extra>",
            ))
            fig_growth.update_layout(
                **CHART_LAYOUT,
                height=280,
                title=dict(text="Cumulative Money Invested", font=dict(size=13, color="#6b7280")),
                showlegend=False,
            )
            st.plotly_chart(fig_growth, use_container_width=True)
        else:
            st.info("No buy trades in selected period.")

    with col_right:
        # Portfolio allocation pie
        if rows:
            pie_df = pd.DataFrame(rows)[["Ticker", "Worth Now ₨"]]
            pie_df = pie_df[pie_df["Worth Now ₨"] > 0]
            fig_pie = px.pie(
                pie_df,
                names="Ticker",
                values="Worth Now ₨",
                color_discrete_sequence=px.colors.qualitative.Vivid,
                hole=0.55,
            )
            fig_pie.update_traces(
                textinfo="label+percent",
                hovertemplate="<b>%{label}</b><br>₨%{value:,.0f}<extra></extra>",
            )
            fig_pie.update_layout(
                **CHART_LAYOUT,
                height=280,
                title=dict(text="How My Money Is Spread", font=dict(size=13, color="#6b7280")),
                showlegend=False,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 5 — PER STOCK DRILL-DOWN
    # ════════════════════════════════════════════════════════════════════════
    if selected_ticker != "All Stocks" and selected_ticker in holdings_dict:
        st.markdown(f'<div class="sec-title">🔍 Deep Dive: {selected_ticker}</div>', unsafe_allow_html=True)

        h   = holdings_dict[selected_ticker]
        cur = current_prices.get(selected_ticker, h["avg_price"])
        pnl = (cur - h["avg_price"]) * h["qty"]
        pct = (cur - h["avg_price"]) / h["avg_price"] * 100 if h["avg_price"] else 0
        t_divs = dividends[dividends["Ticker"] == selected_ticker]["Net Amount Paid"].sum()

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(kpi_card("Shares I Own", f"{int(h['qty']):,}", accent="#3b82f6"), unsafe_allow_html=True)
        with m2:
            st.markdown(kpi_card("Avg I Paid", f"₨{h['avg_price']:,.2f}", accent="#8b5cf6"), unsafe_allow_html=True)
        with m3:
            st.markdown(kpi_card("Price Today", f"₨{cur:,.2f}", delta=pct, accent="#22c55e" if pct >= 0 else "#ef4444"), unsafe_allow_html=True)
        with m4:
            st.markdown(kpi_card("Dividends Earned", fmt_pkr(t_divs), accent="#f59e0b"), unsafe_allow_html=True)

        st.markdown("")

        stock_trades = trades_f[trades_f["Ticker"] == selected_ticker].sort_values("Trade Date")

        # ── Buy/Sell/IPO chart ──
        # Classify every transaction type into one of three buckets
        _BUY_TYPES = {"buy"}
        _IPO_TYPES = {"ipo", "allotment", "ipo allotment", "ipo_allotment", "initial public offering"}
        _SEL_TYPES = {"sell", "sale"}

        def _txn_class(t):
            t = str(t).strip().lower()
            if t in _IPO_TYPES:   return "ipo"
            if t in _BUY_TYPES:   return "buy"
            if t in _SEL_TYPES:   return "sell"
            # Fuzzy catch-all: anything containing "ipo" → ipo bucket
            if "ipo" in t or "allot" in t:  return "ipo"
            if "sell" in t or "sale" in t:  return "sell"
            return "buy"   # default: treat unknown acquiring types as buys

        stock_trades = stock_trades.copy()
        stock_trades["_cls"] = stock_trades["Transaction Type"].apply(_txn_class)

        buys_s  = stock_trades[stock_trades["_cls"] == "buy"]
        ipos_s  = stock_trades[stock_trades["_cls"] == "ipo"]
        sells_s = stock_trades[stock_trades["_cls"] == "sell"]

        fig_stock = go.Figure()

        # Avg buy line
        if h["avg_price"]:
            fig_stock.add_hline(
                y=h["avg_price"],
                line_dash="dot",
                line_color="#f59e0b",
                annotation_text=f"Avg Buy ₨{h['avg_price']:,.2f}",
                annotation_font_color="#f59e0b",
            )

        # Current price line
        fig_stock.add_hline(
            y=cur,
            line_color="#22c55e" if pct >= 0 else "#ef4444",
            line_width=2,
            annotation_text=f"Today ₨{cur:,.2f}",
            annotation_font_color="#22c55e" if pct >= 0 else "#ef4444",
        )

        # Regular buy markers (green triangle-up)
        if not buys_s.empty:
            fig_stock.add_trace(go.Scatter(
                x=buys_s["Trade Date"], y=buys_s["Price"],
                mode="markers",
                name="I Bought",
                marker=dict(color="#22c55e", size=12, symbol="triangle-up",
                            line=dict(color="#fff", width=1)),
                hovertemplate="<b>BOUGHT</b><br>%{x|%b %d, %Y}<br>Price: ₨%{y:,.2f}<br>Qty: %{customdata:,}<extra></extra>",
                customdata=buys_s["Quantity"],
            ))

        # IPO allotment markers (purple diamond — distinct from regular buys)
        if not ipos_s.empty:
            fig_stock.add_trace(go.Scatter(
                x=ipos_s["Trade Date"], y=ipos_s["Price"],
                mode="markers",
                name="IPO Allotment",
                marker=dict(color="#a855f7", size=14, symbol="diamond",
                            line=dict(color="#fff", width=1.5)),
                hovertemplate="<b>IPO ALLOTMENT</b><br>%{x|%b %d, %Y}<br>Price: ₨%{y:,.2f}<br>Qty: %{customdata:,}<extra></extra>",
                customdata=ipos_s["Quantity"],
            ))

        # Sell markers (red triangle-down)
        if not sells_s.empty:
            fig_stock.add_trace(go.Scatter(
                x=sells_s["Trade Date"], y=sells_s["Price"],
                mode="markers",
                name="I Sold",
                marker=dict(color="#ef4444", size=12, symbol="triangle-down",
                            line=dict(color="#fff", width=1)),
                hovertemplate="<b>SOLD</b><br>%{x|%b %d, %Y}<br>Price: ₨%{y:,.2f}<br>Qty: %{customdata:,}<extra></extra>",
                customdata=sells_s["Quantity"],
            ))

        # Dividend markers for this stock on the chart
        stock_divs = dividends[dividends["Ticker"] == selected_ticker].copy()
        stock_divs["Payment Date"] = pd.to_datetime(stock_divs["Payment Date"], errors="coerce")
        stock_divs = stock_divs.dropna(subset=["Payment Date"])

        if not stock_divs.empty:
            # Plot dividends at current-price y level with gold star markers
            fig_stock.add_trace(go.Scatter(
                x=stock_divs["Payment Date"],
                y=[cur] * len(stock_divs),
                mode="markers",
                name="Dividend Paid",
                marker=dict(color="#f59e0b", size=14, symbol="star",
                            line=dict(color="#fff", width=1)),
                hovertemplate=(
                    "<b>💰 DIVIDEND</b><br>"
                    "%{x|%d %b %Y}<br>"
                    "Securities: %{customdata[0]:,}<br>"
                    "Rate/Share: ₨%{customdata[1]:.2f}<br>"
                    "Net Received: ₨%{customdata[2]:,.0f}<extra></extra>"
                ),
                customdata=stock_divs[[
                    "No. of Securities", "Rate Per Security", "Net Amount Paid"
                ]].values,
            ))

        fig_stock.update_layout(
            **CHART_LAYOUT,
            height=360,
            title=dict(
                text=f"{selected_ticker} — Buy points, Dividends & Current Price",
                font=dict(size=13, color="#6b7280")
            ),
        )
        st.plotly_chart(fig_stock, use_container_width=True)
        if not stock_divs.empty:
            st.markdown(
                '<div style="font-size:11px;color:#6b7280;margin-top:-10px;margin-bottom:8px;">' +
                '⭐ Gold stars = dividend payment dates. Hover for amount & securities.</div>',
                unsafe_allow_html=True
            )

        # ── Trade history table ──
        if not stock_trades.empty:
            st.markdown(f"**All {selected_ticker} Trades**")
            disp_cols = ["Trade Date","Transaction Type","Quantity","Price","Commission","Taxes and Fees","Net Total Value"]
            disp = stock_trades[disp_cols].copy()
            disp["Trade Date"] = disp["Trade Date"].dt.strftime("%Y-%m-%d")
            st.dataframe(disp, use_container_width=True, hide_index=True)
            _dl_button(stock_trades[disp_cols], f"{selected_ticker}_trades.xlsx",
                       f"⬇ Download {selected_ticker} Trades")

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 6 — DIVIDEND INSIGHTS
    # ════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">💰 Dividend Income</div>', unsafe_allow_html=True)

    # Use ALL dividends (not date-filtered) for prediction accuracy
    all_divs = dividends.copy()

    if not all_divs.empty:

        # ── 6a: Detailed dividend history table ───────────────────────────────
        st.markdown(
            '<div style="font-size:12px;font-weight:600;color:#94a3b8;margin-bottom:8px;">' +
            '📅 Complete Dividend History</div>', unsafe_allow_html=True
        )
        div_detail = all_divs[[
            "Payment Date", "Ticker", "No. of Securities",
            "Rate Per Security", "Gross Dividend",
            "Zakat Deducted", "Tax Deducted", "Net Amount Paid"
        ]].copy().sort_values("Payment Date", ascending=False)
        div_detail["Payment Date"] = pd.to_datetime(
            div_detail["Payment Date"]).dt.strftime("%d %b %Y")

        # Build HTML table
        def _div_history_table(df):
            hdr = ["Date", "Ticker", "Securities", "Rate/Share ₨",
                   "Gross ₨", "Zakat ₨", "Tax ₨", "Net Received ₨"]
            hdr_html = "".join(f"<th>{h}</th>" for h in hdr)
            rows_html = ""
            for _, r in df.iterrows():
                rows_html += f"""<tr>
                    <td style="color:#94a3b8;">{r["Payment Date"]}</td>
                    <td style="font-weight:600;color:#f1f5f9;">{r["Ticker"]}</td>
                    <td>{int(r["No. of Securities"]):,}</td>
                    <td>{float(r["Rate Per Security"]):.2f}</td>
                    <td>{float(r["Gross Dividend"]):,.0f}</td>
                    <td style="color:#ef4444;">{float(r["Zakat Deducted"]):,.0f}</td>
                    <td style="color:#ef4444;">{float(r["Tax Deducted"]):,.0f}</td>
                    <td style="color:#22c55e;font-weight:600;">{float(r["Net Amount Paid"]):,.0f}</td>
                </tr>"""
            return f"""
            <div style="overflow-x:auto;max-height:280px;overflow-y:auto;margin-bottom:16px;">
            <table style="width:100%;border-collapse:collapse;
                          font-family:'IBM Plex Mono',monospace;font-size:12px;color:#94a3b8;">
              <thead style="position:sticky;top:0;background:#111318;z-index:1;">
                <tr style="border-bottom:1px solid #252836;">{hdr_html}</tr>
              </thead>
              <tbody>{rows_html}</tbody>
            </table></div>
            <style>
              .div-tbl tr:hover td {{ background:rgba(255,255,255,0.03); }}
              .div-tbl th, .div-tbl td {{ padding:8px 12px;text-align:right;
                border-bottom:1px solid #1e2130; }}
              .div-tbl th {{ color:#6b7280;font-size:10px;letter-spacing:1px;
                text-transform:uppercase;font-weight:500; }}
              .div-tbl td:first-child,.div-tbl th:first-child {{ text-align:left; }}
            </style>"""
        st.markdown(_div_history_table(div_detail), unsafe_allow_html=True)
        _dl_button(div_detail, "dividend_history.xlsx", "⬇ Download Dividend History")

        # ── 6b: Next dividend prediction ──────────────────────────────────────
        st.markdown(
            '<div style="font-size:12px;font-weight:600;color:#94a3b8;margin-bottom:8px;">' +
            '🔮 Predicted Next Dividend</div>', unsafe_allow_html=True
        )
        st.markdown(
            '<div style="font-size:11px;color:#4b5563;margin-bottom:10px;">'
            'Based on each stock\'s historical dividend pattern. '
            'Prediction = last payment date + average gap between payments.</div>',
            unsafe_allow_html=True
        )

        from datetime import date as dt_date
        today = pd.Timestamp.today().normalize()

        pred_rows = []
        for ticker, grp in all_divs.groupby("Ticker"):
            dates = pd.to_datetime(grp["Payment Date"]).sort_values().reset_index(drop=True)
            if len(dates) < 1:
                continue
            last_date    = dates.iloc[-1]
            avg_net      = grp["Net Amount Paid"].mean()
            avg_rate     = grp["Rate Per Security"].mean()
            avg_secs     = grp["No. of Securities"].mean()
            total_net    = grp["Net Amount Paid"].sum()
            payments     = len(dates)

            # Predict gap: use average interval if 2+ payments, else assume 180 days
            if len(dates) >= 2:
                gaps     = [(dates.iloc[i+1] - dates.iloc[i]).days for i in range(len(dates)-1)]
                avg_gap  = int(np.mean(gaps))
            else:
                avg_gap  = 180   # assume semi-annual

            predicted_date = last_date + pd.Timedelta(days=avg_gap)
            days_away      = (predicted_date - today).days

            if days_away < 0:
                status = "⚠️ Overdue"
                status_color = "#f59e0b"
            elif days_away <= 30:
                status = "🔔 Soon"
                status_color = "#22c55e"
            else:
                status = f"in {days_away} days"
                status_color = "#94a3b8"

            pred_rows.append({
                "Ticker":           ticker,
                "Last Paid":        last_date.strftime("%d %b %Y"),
                "# Payments":       payments,
                "Avg Gap (days)":   avg_gap,
                "Predicted Next":   predicted_date.strftime("%d %b %Y"),
                "Days Away":        days_away,
                "Status":           status,
                "StatusColor":      status_color,
                "Est. Amount ₨":    round(avg_net, 0),
                "Total Earned ₨":   round(total_net, 0),
            })

        if pred_rows:
            pred_df = pd.DataFrame(pred_rows).sort_values("Days Away")

            def _pred_table(df):
                hdr = ["Ticker", "Last Paid", "Payments", "Avg Gap",
                       "Predicted Next", "Status", "Est. Amount ₨", "Total Earned ₨"]
                hdr_html = "".join(f"<th>{h}</th>" for h in hdr)
                rows_html = ""
                for _, r in df.iterrows():
                    rows_html += f"""<tr>
                        <td style="font-weight:600;color:#f1f5f9;">{r["Ticker"]}</td>
                        <td style="color:#94a3b8;">{r["Last Paid"]}</td>
                        <td style="text-align:center;">{r["# Payments"]}</td>
                        <td style="text-align:center;">{r["Avg Gap (days)"]}d</td>
                        <td style="color:#f59e0b;font-weight:600;">{r["Predicted Next"]}</td>
                        <td style="color:{r["StatusColor"]};font-weight:600;">{r["Status"]}</td>
                        <td style="color:#22c55e;">{r["Est. Amount ₨"]:,.0f}</td>
                        <td style="color:#22c55e;">{r["Total Earned ₨"]:,.0f}</td>
                    </tr>"""
                return f"""
                <div style="overflow-x:auto;margin-bottom:16px;">
                <table style="width:100%;border-collapse:collapse;
                              font-family:'IBM Plex Mono',monospace;font-size:12px;color:#94a3b8;">
                  <thead style="background:#111318;">
                    <tr style="border-bottom:1px solid #252836;">{hdr_html}</tr>
                  </thead>
                  <tbody>{rows_html}</tbody>
                </table></div>
                <style>
                  table tr:hover td {{ background:rgba(255,255,255,0.02); }}
                  table th,table td {{ padding:8px 12px;border-bottom:1px solid #1e2130; }}
                  table th {{ color:#6b7280;font-size:10px;letter-spacing:1px;
                    text-transform:uppercase;font-weight:500;text-align:right; }}
                  table td {{ text-align:right; }}
                  table td:first-child,table th:first-child {{ text-align:left; }}
                </style>"""
            st.markdown(_pred_table(pred_df), unsafe_allow_html=True)
            _dl_button(pred_df.drop(columns=["StatusColor"], errors="ignore"),
                       "dividend_predictions.xlsx", "⬇ Download Predictions")
            st.markdown(
                '<div style="font-size:10px;color:#374151;margin-bottom:16px;">' +
                '⚠️ Prediction is an estimate only. Actual dividend dates are announced by ' +
                'company boards and depend on profits. Always check PSX announcements.</div>',
                unsafe_allow_html=True
            )

        # ── 6c: Charts — by ticker + monthly trend ────────────────────────────
        d1, d2 = st.columns(2)
        with d1:
            divs_by_ticker = (
                all_divs.groupby("Ticker")["Net Amount Paid"]
                .sum().reset_index().sort_values("Net Amount Paid", ascending=True)
            )
            fig_divbar = px.bar(
                divs_by_ticker, x="Net Amount Paid", y="Ticker",
                orientation="h",
                color="Net Amount Paid",
                color_continuous_scale=["#1e3a5f", "#22c55e"],
                labels={"Net Amount Paid": "Dividends Received (₨)"},
            )
            fig_divbar.update_layout(
                **CHART_LAYOUT, height=300,
                coloraxis_showscale=False,
                title=dict(text="Which Stock Pays Most Dividends",
                           font=dict(size=13, color="#6b7280")),
            )
            fig_divbar.update_traces(hovertemplate="<b>%{y}</b><br>₨%{x:,.0f}<extra></extra>")
            st.plotly_chart(fig_divbar, use_container_width=True)

        with d2:
            divs_timeline = all_divs.copy()
            divs_timeline["Month"] = pd.to_datetime(
                divs_timeline["Payment Date"]).dt.to_period("M").astype(str)
            divs_monthly = divs_timeline.groupby("Month")["Net Amount Paid"].sum().reset_index()
            fig_divtime = px.bar(
                divs_monthly, x="Month", y="Net Amount Paid",
                color_discrete_sequence=["#f59e0b"],
                labels={"Net Amount Paid": "Monthly Dividends (₨)"},
            )
            fig_divtime.update_layout(
                **CHART_LAYOUT, height=300,
                title=dict(text="Dividends Over Time (Monthly)",
                           font=dict(size=13, color="#6b7280")),
            )
            fig_divtime.update_traces(
                hovertemplate="<b>%{x}</b><br>₨%{y:,.0f}<extra></extra>"
            )
            st.plotly_chart(fig_divtime, use_container_width=True)

    else:
        st.info("No dividend records found.")

    # ════════════════════════════════════════════════════════════════════════
    #  SECTION 7 — CASH FLOW
    # ════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-title">💸 Cash Flow — Money In & Out</div>', unsafe_allow_html=True)

    # ── A: FUNDS SHEET — bank-level cash movements ───────────────────────────
    funds["_txn"] = funds["Transfer Type"].str.strip().str.lower()
    dep_rows = funds[funds["_txn"].str.contains("deposit|credit|ipo", na=False)]
    wdl_rows = funds[funds["_txn"].str.contains("withdraw|debit|transfer out", na=False)]
    # Use "Amount Transferred To Exposure" — this is the money actually usable
    # for buying stocks (after tax/zakat/charges have been held back by broker).
    # "Amount Deposit" includes held amounts that can't be used for trading.
    total_dep     = dep_rows["Amount Deposit"].sum()           # total cash added (for display)
    total_exposure = dep_rows["Amount Transferred To Exposure"].sum()  # tradeable exposure
    total_wdl     = wdl_rows["Amount Deposit"].abs().sum()

    # ── B: TRADES SHEET — use Net Total Value (already includes commissions/fees)
    #    BUY  net value = what you actually paid out (positive in your sheet)
    #    SELL net value = what you actually received back (positive in your sheet)
    _BUY_MASK  = trades["Transaction Type"].str.strip().str.lower().isin(
        ["buy", "ipo", "allotment", "ipo allotment", "ipo_allotment"]
    )
    _SELL_MASK = trades["Transaction Type"].str.strip().str.lower().str.contains(
        "sell|sale", na=False
    )
    # Net Total Value is always stored as a positive number in the sheet
    total_buy_spend  = trades[_BUY_MASK]["Net Total Value"].abs().sum()
    total_sell_recvd = trades[_SELL_MASK]["Net Total Value"].abs().sum()
    total_dividends_rcvd = dividends["Net Amount Paid"].sum()

    # ── C: Estimated cash in broker account ──────────────────────────────────
    div_in_broker = total_dividends_rcvd if div_to_broker else 0
    estimated_cash = total_exposure + total_sell_recvd + div_in_broker - total_buy_spend - total_wdl

    # ── KPI Cards ─────────────────────────────────────────────────────────────
    cf1, cf2, cf3, cf4, cf5 = st.columns(5)
    with cf1:
        st.markdown(kpi_card(
            "Transferred to Exposure", fmt_pkr(total_exposure), accent="#22c55e"
        ), unsafe_allow_html=True)
    with cf2:
        st.markdown(kpi_card(
            "Spent Buying Stocks", fmt_pkr(total_buy_spend), accent="#3b82f6"
        ), unsafe_allow_html=True)
    with cf3:
        st.markdown(kpi_card(
            "Received from Sales", fmt_pkr(total_sell_recvd), accent="#a855f7"
        ), unsafe_allow_html=True)
    with cf4:
        st.markdown(kpi_card(
            "Withdrawn to Bank", fmt_pkr(total_wdl), accent="#ef4444"
        ), unsafe_allow_html=True)
    with cf5:
        cash_label  = "Est. Cash in Broker" if estimated_cash >= 0 else "Est. Cash Deficit"
        cash_accent = "#f59e0b" if estimated_cash >= 0 else "#ef4444"
        st.markdown(kpi_card(cash_label, fmt_pkr(abs(estimated_cash)), accent=cash_accent),
                    unsafe_allow_html=True)

    # Dividend card shown separately when dividends go to bank
    if not div_to_broker:
        st.markdown(kpi_card(
            "Dividends → Your Bank",
            fmt_pkr(total_dividends_rcvd),
            accent="#f59e0b"
        ), unsafe_allow_html=True)

    # ── Compact calculation tooltip using expander ───────────────────────────
    sign = lambda v: f"+{fmt_pkr(v)}" if v >= 0 else f"−{fmt_pkr(abs(v))}"
    div_row_text = (
        f"  + {fmt_pkr(total_dividends_rcvd)}   Dividends (broker account)"
        if div_to_broker else
        f"  ⚠️ Dividends {fmt_pkr(total_dividends_rcvd)} → your bank (excluded)"
    )
    held_amt = total_dep - total_exposure
    with st.expander("ℹ️ How is Est. Cash calculated? Click to see the breakdown"):
        st.markdown(
            f'''<div style="font-family:'IBM Plex Mono',monospace;font-size:12px;
                            line-height:2;color:#94a3b8;">
<span style="color:#22c55e;">+ {fmt_pkr(total_exposure)}</span>  Transferred to Exposure (usable trading cash)<br>
<span style="color:#4b5563;font-size:11px;">&nbsp;&nbsp;&nbsp;(Total deposited: {fmt_pkr(total_dep)},
 held for charges: {fmt_pkr(held_amt)})</span><br>
<span style="color:#a855f7;">+ {fmt_pkr(total_sell_recvd)}</span>  Received from stock sales<br>
{div_row_text}<br>
<span style="color:#ef4444;">− {fmt_pkr(total_buy_spend)}</span>  Spent buying stocks (Net Total Value)<br>
<span style="color:#ef4444;">− {fmt_pkr(total_wdl)}</span>  Withdrawn to bank<br>
<hr style="border-color:#252836;margin:4px 0;">
<b style="color:{"#22c55e" if estimated_cash >= 0 else "#ef4444"};">
= {sign(estimated_cash)}  Est. Cash in Broker</b>
</div>''',
            unsafe_allow_html=True
        )

    # ── Running Cash Balance Chart ─────────────────────────────────────────────
    # Two separate running totals on one chart:
    #   "Broker Cash" line  = Deposits + Sales (+ Dividends if toggle ON) − Buys − Withdrawals
    #   "Dividend Cash" line = cumulative dividends (only shown when toggle OFF, i.e. go to bank)
    # Events are tagged and sorted so same-day deposits always precede buys.

    events = []
    for _, r in dep_rows.iterrows():
        # Use Exposure amount (tradeable cash), not full deposit (which includes held charges)
        exp_amt = r["Amount Transferred To Exposure"] if r["Amount Transferred To Exposure"] > 0 else r["Amount Deposit"]
        events.append({"Date": r["Date"], "Amount": exp_amt,
                        "Type": "Deposit", "Track": "broker"})
    for _, r in wdl_rows.iterrows():
        events.append({"Date": r["Date"], "Amount": -abs(r["Amount Deposit"]),
                        "Type": "Withdrawal", "Track": "broker"})
    for _, r in trades[_BUY_MASK].iterrows():
        events.append({"Date": r["Trade Date"], "Amount": -abs(r["Net Total Value"]),
                        "Type": "Stock Buy", "Track": "broker"})
    for _, r in trades[_SELL_MASK].iterrows():
        events.append({"Date": r["Trade Date"], "Amount": abs(r["Net Total Value"]),
                        "Type": "Stock Sale", "Track": "broker"})
    for _, r in dividends.iterrows():
        track = "broker" if div_to_broker else "dividend"
        events.append({"Date": r["Payment Date"], "Amount": r["Net Amount Paid"],
                        "Type": "Dividend", "Track": track})

    _SORT_ORDER = {
        "Deposit":    1,
        "Stock Sale": 2,
        "Dividend":   3,   # within broker track, dividends count as cash in
        "Stock Buy":  4,
        "Withdrawal": 5,
    }

    ev_df = pd.DataFrame(events).dropna(subset=["Date"])
    ev_df["Date"]       = pd.to_datetime(ev_df["Date"])
    ev_df["_day_order"] = ev_df["Type"].map(_SORT_ORDER).fillna(4)
    ev_df = (
        ev_df
        .sort_values(["Date", "_day_order"])
        .drop(columns=["_day_order"])
        .reset_index(drop=True)
    )

    # Compute running totals per track separately
    broker_ev  = ev_df[ev_df["Track"] == "broker"].copy()
    div_ev     = ev_df[ev_df["Track"] == "dividend"].copy()

    broker_ev["Running Cash"] = broker_ev["Amount"].cumsum()
    if not div_ev.empty:
        div_ev["Running Cash"] = div_ev["Amount"].cumsum()

    COLOR_MAP = {
        "Deposit":    "#22c55e",
        "Withdrawal": "#ef4444",
        "Stock Buy":  "#3b82f6",
        "Stock Sale": "#a855f7",
        "Dividend":   "#f59e0b",
    }

    fig_funds = go.Figure()

    # ── Broker cash line ──
    if not broker_ev.empty:
        fig_funds.add_trace(go.Scatter(
            x=broker_ev["Date"], y=broker_ev["Running Cash"],
            mode="lines",
            name="Broker Cash Balance",
            line=dict(color="#06b6d4", width=2.5),
            fill="tozeroy",
            fillcolor="rgba(6,182,212,0.06)",
            hoverinfo="skip",
            showlegend=True,
        ))
        # Coloured event dots on broker line
        for evt_type, color in COLOR_MAP.items():
            if evt_type == "Dividend" and not div_to_broker:
                continue   # dividends not on broker line
            sub = broker_ev[broker_ev["Type"] == evt_type]
            if sub.empty:
                continue
            fig_funds.add_trace(go.Scatter(
                x=sub["Date"], y=sub["Running Cash"],
                mode="markers", name=evt_type,
                marker=dict(color=color, size=8, line=dict(color="#0d0f14", width=1)),
                hovertemplate=(
                    f"<b>{evt_type}</b><br>%{{x|%b %d, %Y}}<br>"
                    f"Amount: ₨%{{customdata:,.0f}}<br>"
                    f"Broker balance after: ₨%{{y:,.0f}}<extra></extra>"
                ),
                customdata=sub["Amount"].abs(),
            ))

    # ── Dividend cash line (only when dividends go to bank) ──
    if not div_to_broker and not div_ev.empty:
        fig_funds.add_trace(go.Scatter(
            x=div_ev["Date"], y=div_ev["Running Cash"],
            mode="lines+markers",
            name="Cumulative Dividends (Bank)",
            line=dict(color="#f59e0b", width=2, dash="dot"),
            marker=dict(color="#f59e0b", size=8, symbol="star",
                        line=dict(color="#0d0f14", width=1)),
            hovertemplate=(
                "<b>Dividend → Bank</b><br>%{x|%b %d, %Y}<br>"
                "This payment: ₨%{customdata:,.0f}<br>"
                "Total dividends to date: ₨%{y:,.0f}<extra></extra>"
            ),
            customdata=div_ev["Amount"].abs(),
        ))

    fig_funds.add_hline(y=0, line_dash="dot", line_color="#374151", line_width=1)
    fig_funds.update_layout(
        **CHART_LAYOUT,
        height=340,
        title=dict(
            text="Cash Balance — Broker account" + (" + Dividends to bank" if not div_to_broker else ""),
            font=dict(size=13, color="#6b7280")
        ),
    )
    fig_funds.update_layout(legend=dict(
        orientation="h", y=-0.28, x=0,
        bgcolor="rgba(0,0,0,0)", font=dict(size=11),
    ))
    st.plotly_chart(fig_funds, use_container_width=True)

    # ── Footer ──
    st.markdown("""
    <div style="text-align:center;color:#374151;font-size:11px;margin-top:40px;padding-top:20px;border-top:1px solid #1e2130;">
    PSX Portfolio Dashboard · Built with Streamlit + Plotly · Prices via PSX Data Portal EOD API · IPO allotments included in holdings · Refreshed every 5 min
    </div>
    """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  INSTRUCTIONS (shown before file upload)
# ════════════════════════════════════════════════════════════════════════════
def _show_instructions():
    st.markdown("---")
    with st.expander("📖 How to use this dashboard", expanded=True):
        st.markdown("""
**Step 1 — Prepare your Excel file** with these 3 sheets:

| Sheet | Key Columns |
|---|---|
| **Trades** | Trade Date, Transaction Type, Ticker, Quantity, Price, Commission, Taxes and Fees, Net Total Value |
| **Dividends** | Payment Date, Ticker, Net Amount Paid, Gross Dividend |
| **Funds** | Date, Transfer Type, Amount Deposit, Amount Transferred To Exposure |

**Step 2 — Upload** using the sidebar button.

**Step 3 — Explore!** Use the sidebar to:
- Filter by a specific stock
- Change the date range

> 💡 **No data?** Run `python generate_sample_data.py` to create a sample Excel file.
        """)


if __name__ == "__main__":
    main()
