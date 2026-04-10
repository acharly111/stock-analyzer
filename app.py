import streamlit as st
import pandas as pd
import html
import yfinance as yf
import tempfile
import os
from engine import build_output, save_outputs

st.set_page_config(page_title="Stock Analyzer", layout="centered")
st.title("Stock Analyzer")


# ---------- Helpers ----------
def to_scalar(x):
    if isinstance(x, pd.Series):
        return x.iloc[0]
    return x


@st.cache_data(ttl=900)
def load_master_df():
    master_df = pd.read_csv("stocks_master.csv", encoding="latin1", engine="python")
    master_df.columns = master_df.columns.str.strip()

    if "Industruy" in master_df.columns:
        master_df = master_df.rename(columns={"Industruy": "Industry"})

    return master_df


def clean_cell(x):
    if pd.isna(x):
        return ""
    return html.escape(str(x)).replace("\n", "<br>")


def format_catalyst(text):
    if pd.isna(text):
        return ""
    text = str(text)
    parts = text.split("\n", 1)

    first_line = html.escape(parts[0])
    rest = html.escape(parts[1]) if len(parts) > 1 else ""

    if rest:
        return f"<b>{first_line}</b><br>{rest.replace(chr(10), '<br>')}"
    return f"<b>{first_line}</b>"


def calculate_rsi_and_ma(data):
    data = data.copy()
    data["SMA_200"] = data["Close"].rolling(200).mean()

    delta = data["Close"].diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)

    avg_gain = gain.ewm(alpha=1 / 14, min_periods=14).mean()
    avg_loss = loss.ewm(alpha=1 / 14, min_periods=14).mean()

    rs = avg_gain / avg_loss
    data["RSI_14"] = 100 - (100 / (1 + rs))
    data["RSI_MA_14"] = data["RSI_14"].rolling(14).mean()
    return data


@st.cache_data(ttl=300)
def get_market_snapshot():
    spy = yf.download(
        "SPY",
        period="2y",
        interval="1d",
        progress=False,
        auto_adjust=False,
        group_by="column"
    )

    vix = yf.download(
        "^VIX",
        period="6mo",
        interval="1d",
        progress=False,
        auto_adjust=False,
        group_by="column"
    )

    if spy.empty:
        raise ValueError("SPY data could not be loaded from Yahoo Finance.")
    if vix.empty:
        raise ValueError("VIX data could not be loaded from Yahoo Finance.")

    if isinstance(spy.columns, pd.MultiIndex):
        spy.columns = spy.columns.get_level_values(0)
    if isinstance(vix.columns, pd.MultiIndex):
        vix.columns = vix.columns.get_level_values(0)

    spy = calculate_rsi_and_ma(spy)

    spy_close_series = spy["Close"].dropna()
    spy_sma_200_series = spy["SMA_200"].dropna()
    spy_rsi_series = spy["RSI_14"].dropna()
    spy_rsi_ma_series = spy["RSI_MA_14"].dropna()
    vix_close_series = vix["Close"].dropna()

    spy_close = float(to_scalar(spy_close_series.iloc[-1])) if not spy_close_series.empty else None
    spy_sma_200 = float(to_scalar(spy_sma_200_series.iloc[-1])) if not spy_sma_200_series.empty else None
    spy_rsi = float(to_scalar(spy_rsi_series.iloc[-1])) if not spy_rsi_series.empty else None
    spy_rsi_ma = float(to_scalar(spy_rsi_ma_series.iloc[-1])) if not spy_rsi_ma_series.empty else None
    vix_close = float(to_scalar(vix_close_series.iloc[-1])) if not vix_close_series.empty else None

    spy_dist_200 = None
    if spy_close is not None and spy_sma_200 not in [None, 0]:
        spy_dist_200 = ((spy_close / spy_sma_200) - 1) * 100

    spy_rsi_dist = None
    if spy_rsi is not None and spy_rsi_ma not in [None, 0]:
        spy_rsi_dist = ((spy_rsi / spy_rsi_ma) - 1) * 100

    return {
        "SPY_price": round(spy_close, 2) if spy_close is not None else None,
        "SPY_dist_200": round(spy_dist_200, 2) if spy_dist_200 is not None else None,
        "SPY_RSI": round(spy_rsi, 2) if spy_rsi is not None else None,
        "SPY_RSI_dist_MA14": round(spy_rsi_dist, 2) if spy_rsi_dist is not None else None,
        "VIX": round(vix_close, 2) if vix_close is not None else None,
    }


def market_status_spy_dist(value, overbought=10.0, oversold=-10.0):
    if value is None:
        return "N/A", ""
    if value > overbought:
        return f"{value:.2f}% overbought", "#f8d7da"
    if value < oversold:
        return f"{value:.2f}% oversold", "#d4edda"
    return f"{value:.2f}% neutral", "#fff3cd"


def market_status_rsi(value, overbought=69.0, oversold=35.0):
    if value is None:
        return "N/A", ""
    if value > overbought:
        return f"{value:.2f} overbought", "#f8d7da"
    if value < oversold:
        return f"{value:.2f} oversold", "#d4edda"
    return f"{value:.2f} neutral", "#fff3cd"


def market_status_rsi_dist(value, overbought=20.0, oversold=-20.0):
    if value is None:
        return "N/A", ""
    if value > overbought:
        return f"{value:.2f}% overbought", "#f8d7da"
    if value < oversold:
        return f"{value:.2f}% oversold", "#d4edda"
    return f"{value:.2f}% neutral", "#fff3cd"


def market_status_vix(value, fear=30.0, no_fear=20.0):
    if value is None:
        return "N/A", ""
    if value > fear:
        return f"{value:.2f} Fear", ""
    if value < no_fear:
        return f"{value:.2f} No Fear", ""
    return f"{value:.2f} Mid Fear", ""


def build_excel_bytes(df):
    excel_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    excel_tmp.close()

    txt_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
    txt_tmp.close()

    try:
        save_outputs(
            df,
            excel_file=excel_tmp.name,
            txt_file=txt_tmp.name,
            selected_sma=st.session_state["selected_sma"],
            buy_rsi_dist_max=st.session_state["buy_rsi_dist_max"],
            buy_dist_selected_sma_max=st.session_state["buy_dist_selected_sma_max"],
            sell_rsi_min=st.session_state["sell_rsi_min"],
            sell_dist_selected_sma_min=st.session_state["sell_dist_selected_sma_min"]
        )

        with open(excel_tmp.name, "rb") as f:
            return f.read()

    finally:
        try:
            os.remove(excel_tmp.name)
        except Exception:
            pass
        try:
            os.remove(txt_tmp.name)
        except Exception:
            pass


# ---------- Session defaults ----------
defaults = {
    "selected_sma": 200,
    "buy_rsi_dist_max": -10.0,
    "buy_dist_selected_sma_max": 10.0,
    "sell_rsi_min": 70.0,
    "sell_dist_selected_sma_min": 40.0,
    "spy_dist_overbought": 10.0,
    "spy_dist_oversold": -10.0,
    "rsi_overbought": 69.0,
    "rsi_oversold": 35.0,
    "rsi_dist_overbought": 20.0,
    "rsi_dist_oversold": -20.0,
    "vix_fear": 30.0,
    "vix_no_fear": 20.0,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

if "analysis_df" not in st.session_state:
    st.session_state["analysis_df"] = None


# ---------- Tabs ----------
data_tab, params_tab, stocks_tab = st.tabs(["Market & Analysis", "Parameters", "Stock List"])


with data_tab:
    if st.button("Refresh market + stocks data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.subheader("General Market Condition")

    try:
        market_data = get_market_snapshot()
    except Exception as e:
        market_data = {
            "SPY_dist_200": None,
            "SPY_RSI": None,
            "SPY_RSI_dist_MA14": None,
            "VIX": None,
        }
        st.error(f"Market data error: {e}")

    spy_dist_text, spy_dist_color = market_status_spy_dist(
        market_data["SPY_dist_200"],
        overbought=st.session_state["spy_dist_overbought"],
        oversold=st.session_state["spy_dist_oversold"]
    )
    rsi_text, rsi_color = market_status_rsi(
        market_data["SPY_RSI"],
        overbought=st.session_state["rsi_overbought"],
        oversold=st.session_state["rsi_oversold"]
    )
    rsi_dist_text, rsi_dist_color = market_status_rsi_dist(
        market_data["SPY_RSI_dist_MA14"],
        overbought=st.session_state["rsi_dist_overbought"],
        oversold=st.session_state["rsi_dist_oversold"]
    )
    vix_text, _ = market_status_vix(
        market_data["VIX"],
        fear=st.session_state["vix_fear"],
        no_fear=st.session_state["vix_no_fear"]
    )

    market_rows = [
        ("SPY vs 200SMA", spy_dist_text, spy_dist_color),
        ("RSI", rsi_text, rsi_color),
        ("RSI vs 14MA", rsi_dist_text, rsi_dist_color),
        ("VIX", vix_text, ""),
    ]

    market_rows_html = ""
    for metric, value, color in market_rows:
        style = f'background-color:{color};' if color else ""
        market_rows_html += f"""
        <tr>
            <td>{html.escape(metric)}</td>
            <td style="{style}">{html.escape(value)}</td>
        </tr>
        """

    st.markdown(
        f"""
        <style>
        .market-table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            font-size: 17px;
            margin-bottom: 16px;
        }}
        .market-table th, .market-table td {{
            border: 1px solid #ddd;
            padding: 12px;
            text-align: center;
            vertical-align: middle;
            word-wrap: break-word;
        }}
        .market-table th {{
            background-color: #f5f5f5;
            font-weight: 600;
        }}
        .market-table th:nth-child(1), .market-table td:nth-child(1) {{
            width: 45%;
        }}
        .market-table th:nth-child(2), .market-table td:nth-child(2) {{
            width: 55%;
        }}
        </style>

        <table class="market-table">
            <thead>
                <tr>
                    <th>Metric</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                {market_rows_html}
            </tbody>
        </table>
        """,
        unsafe_allow_html=True
    )

    try:
        master_df = load_master_df()
    except Exception as e:
        st.error(f"Error loading CSV: {e}")
        master_df = pd.DataFrame(columns=["Ticker", "Company", "Industry", "Catalyst"])

    st.divider()

    default_tickers = ",".join(master_df["Ticker"].dropna().astype(str).tolist()) if not master_df.empty else ""

    st.subheader("Run Analysis")

    tickers_text = st.text_area(
        "Tickers (comma separated)",
        value=default_tickers,
        height=140
    )

    st.markdown(
        """
        <div style="display:flex; justify-content:space-between; font-weight:700; margin-bottom:6px;">
            <span>Bold</span>
            <span>Cautious</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    selected_sma = st.select_slider(
        f"SMA for BUY/SELL criteria: SMA {st.session_state['selected_sma']}",
        options=[20, 50, 100, 150, 200],
        value=st.session_state["selected_sma"]
    )
    st.session_state["selected_sma"] = selected_sma

    buy_dist_slider_value = st.select_slider(
        "BUY: Price dist [%] from SMA",
        options=[float(x) for x in range(20, -1, -1)],
        value=float(st.session_state["buy_dist_selected_sma_max"])
    )
    st.session_state["buy_dist_selected_sma_max"] = float(buy_dist_slider_value)

    buy_rsi_slider_value = st.select_slider(
        "BUY: RSI distance [%] from MA",
        options=[float(x) for x in range(0, -21, -1)],
        value=float(st.session_state["buy_rsi_dist_max"])
    )
    st.session_state["buy_rsi_dist_max"] = float(buy_rsi_slider_value)

    run_button = st.button("Run analysis", use_container_width=True)

    if run_button:
        tickers = [t.strip().upper() for t in tickers_text.replace("\n", "").split(",") if t.strip()]

        with st.spinner("Running analysis..."):
            df = build_output(
                tickers,
                selected_sma=st.session_state["selected_sma"],
                buy_rsi_dist_max=st.session_state["buy_rsi_dist_max"],
                buy_dist_selected_sma_max=st.session_state["buy_dist_selected_sma_max"],
                sell_rsi_min=st.session_state["sell_rsi_min"],
                sell_dist_selected_sma_min=st.session_state["sell_dist_selected_sma_min"]
            )

        st.session_state["analysis_df"] = df
        st.success("Analysis complete")

    if st.session_state["analysis_df"] is not None:
        df = st.session_state["analysis_df"]

        summary_df = df[[
            "symbol",
            "risk_level",
            "Strong_vs_SPY",
            "next_earnings_date",
            "BUY/SELL signal"
        ]].copy()
        summary_df.columns = ["Name", "Risk level", "Stock vs SPY", "Next earning", "BUY SELL"]

        today = pd.Timestamp.today().normalize()

        summary_rows_html = ""
        for _, row in summary_df.iterrows():
            earnings_text = ""
            earnings_style = ""

            if pd.notna(row["Next earning"]) and str(row["Next earning"]).strip():
                earnings_text = str(row["Next earning"]).strip()
                try:
                    earnings_dt = pd.Timestamp(earnings_text).normalize()
                    days_to_earnings = (earnings_dt - today).days
                    if 0 <= days_to_earnings < 7:
                        earnings_style = 'background-color:#f8d7da;'
                except Exception:
                    pass

            summary_rows_html += f"""
            <tr>
                <td>{html.escape(str(row['Name']) if pd.notna(row['Name']) else '')}</td>
                <td>{html.escape(str(row['Risk level']) if pd.notna(row['Risk level']) else '')}</td>
                <td>{html.escape(str(row['Stock vs SPY']) if pd.notna(row['Stock vs SPY']) else '')}</td>
                <td style="{earnings_style}">{html.escape(earnings_text)}</td>
                <td>{html.escape(str(row['BUY SELL']) if pd.notna(row['BUY SELL']) else '')}</td>
            </tr>
            """

        st.subheader("Summary")
        st.markdown(
            f"""
            <style>
            .summary-table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
                font-size: 14px;
                margin-bottom: 16px;
            }}
            .summary-table th, .summary-table td {{
                border: 1px solid #ddd;
                padding: 10px;
                text-align: center;
                vertical-align: middle;
                word-wrap: break-word;
            }}
            .summary-table th {{
                background-color: #dbeafe;
                color: #1e3a8a;
                font-weight: 700;
            }}
            </style>

            <table class="summary-table">
                <thead>
                    <tr>
                        <th>Name</th>
                        <th>Risk level</th>
                        <th>Stock vs SPY</th>
                        <th>Next earning</th>
                        <th>BUY SELL</th>
                    </tr>
                </thead>
                <tbody>
                    {summary_rows_html}
                </tbody>
            </table>
            """,
            unsafe_allow_html=True
        )

        full_df = df.drop(columns=["risk_level", "Strong_vs_SPY"], errors="ignore")

        with st.expander("Show full analysis table"):
            st.dataframe(full_df, use_container_width=True, hide_index=True)

        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download CSV",
            data=csv_data,
            file_name="risk_analysis.csv",
            mime="text/csv",
            use_container_width=True
        )

        if st.button("Prepare Excel download", use_container_width=True):
            try:
                excel_bytes = build_excel_bytes(df)
                st.download_button(
                    "Download Excel",
                    data=excel_bytes,
                    file_name="risk_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Excel file could not be created: {e}")


with params_tab:
    st.subheader("Market Parameters")

    st.session_state["spy_dist_overbought"] = st.number_input(
        "SPY dist from 200SMA: overbought above (%)",
        value=float(st.session_state["spy_dist_overbought"])
    )
    st.session_state["spy_dist_oversold"] = st.number_input(
        "SPY dist from 200SMA: oversold below (%)",
        value=float(st.session_state["spy_dist_oversold"])
    )
    st.session_state["rsi_overbought"] = st.number_input(
        "RSI: overbought above",
        value=float(st.session_state["rsi_overbought"])
    )
    st.session_state["rsi_oversold"] = st.number_input(
        "RSI: oversold below",
        value=float(st.session_state["rsi_oversold"])
    )
    st.session_state["rsi_dist_overbought"] = st.number_input(
        "RSI dist from 14MA: overbought above (%)",
        value=float(st.session_state["rsi_dist_overbought"])
    )
    st.session_state["rsi_dist_oversold"] = st.number_input(
        "RSI dist from 14MA: oversold below (%)",
        value=float(st.session_state["rsi_dist_oversold"])
    )
    st.session_state["vix_fear"] = st.number_input(
        "VIX: Fear above",
        value=float(st.session_state["vix_fear"])
    )
    st.session_state["vix_no_fear"] = st.number_input(
        "VIX: No Fear below",
        value=float(st.session_state["vix_no_fear"])
    )

    st.divider()

    st.subheader("Stock Parameters")

    st.markdown(f"Selected SMA for stock criteria: **SMA {st.session_state['selected_sma']}**")

    st.session_state["buy_rsi_dist_max"] = st.number_input(
        "BUY: RSI distance [%] from MA",
        value=float(st.session_state["buy_rsi_dist_max"])
    )
    st.session_state["buy_dist_selected_sma_max"] = st.number_input(
        f"BUY: Price dist [%] from SMA {st.session_state['selected_sma']}",
        value=float(st.session_state["buy_dist_selected_sma_max"])
    )
    st.session_state["sell_rsi_min"] = st.number_input(
        "SELL: RSI min",
        value=float(st.session_state["sell_rsi_min"])
    )
    st.session_state["sell_dist_selected_sma_min"] = st.number_input(
        f"SELL: Price dist from SMA {st.session_state['selected_sma']} min (%)",
        value=float(st.session_state["sell_dist_selected_sma_min"])
    )

    st.info("The two BUY sliders on the first tab update these stock parameters automatically.")


with stocks_tab:
    try:
        master_df = load_master_df()
    except Exception as e:
        st.error(f"Error loading CSV: {e}")
        master_df = pd.DataFrame(columns=["Ticker", "Company", "Industry", "Catalyst"])

    st.subheader("Stock List")

    if not master_df.empty:
        stock_list_view = master_df[["Ticker", "Company", "Industry", "Catalyst"]].copy()

        rows_html = ""
        for _, row in stock_list_view.iterrows():
            rows_html += f"""
            <tr>
                <td>{clean_cell(row['Ticker'])}</td>
                <td>{clean_cell(row['Company'])}</td>
                <td>{clean_cell(row['Industry'])}</td>
                <td>{format_catalyst(row['Catalyst'])}</td>
            </tr>
            """

        st.markdown(
            f"""
            <style>
            .stocks-table {{
                width: 100%;
                border-collapse: collapse;
                table-layout: fixed;
                font-size: 12px;
            }}
            .stocks-table th, .stocks-table td {{
                border: 1px solid #ddd;
                padding: 6px 8px;
                vertical-align: top;
                text-align: left;
                word-wrap: break-word;
                white-space: normal;
                line-height: 1.15;
            }}
            .stocks-table th {{
                background-color: #f5f5f5;
                font-weight: 600;
            }}
            .stocks-table th:nth-child(1), .stocks-table td:nth-child(1) {{
                width: 16%;
            }}
            .stocks-table th:nth-child(2), .stocks-table td:nth-child(2) {{
                width: 16%;
            }}
            .stocks-table th:nth-child(3), .stocks-table td:nth-child(3) {{
                width: 16%;
            }}
            .stocks-table th:nth-child(4), .stocks-table td:nth-child(4) {{
                width: 52%;
            }}
            .stocks-table tbody tr {{
                height: 28px;
            }}
            </style>

            <table class="stocks-table">
                <thead>
                    <tr>
                        <th>Ticker</th>
                        <th>Company</th>
                        <th>Industry</th>
                        <th>Catalyst</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
            """,
            unsafe_allow_html=True
        )
    else:
        st.warning("stocks_master.csv was not found or could not be read.")