import streamlit as st
import pandas as pd
import html
import yfinance as yf
from engine import build_output, save_outputs

st.set_page_config(page_title="Stock Analyzer", layout="centered")

st.title("Stock Analyzer")


def to_scalar(x):
    if isinstance(x, pd.Series):
        return x.iloc[0]
    return x


def calculate_rsi_and_ma(data):
    delta = data["Close"].diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)

    avg_gain = gain.ewm(alpha=1/14, min_periods=14).mean()
    avg_loss = loss.ewm(alpha=1/14, min_periods=14).mean()

    rs = avg_gain / avg_loss
    data["RSI_14"] = 100 - (100 / (1 + rs))
    data["RSI_MA_14"] = data["RSI_14"].rolling(14).mean()
    data["SMA_200"] = data["Close"].rolling(200).mean()
    return data


def get_market_snapshot():
    try:
        spy = yf.download(
            "SPY",
            period="1y",
            interval="1d",
            progress=False,
            auto_adjust=False,
            group_by="column"
        )

        vix = yf.download(
            "^VIX",
            period="3mo",
            interval="1d",
            progress=False,
            auto_adjust=False,
            group_by="column"
        )

        if isinstance(spy.columns, pd.MultiIndex):
            spy.columns = spy.columns.get_level_values(0)
        if isinstance(vix.columns, pd.MultiIndex):
            vix.columns = vix.columns.get_level_values(0)

        spy = calculate_rsi_and_ma(spy)

        spy_price = float(to_scalar(spy["Close"].iloc[-1]))
        spy_sma_200 = float(to_scalar(spy["SMA_200"].iloc[-1])) if pd.notna(to_scalar(spy["SMA_200"].iloc[-1])) else None
        spy_rsi = float(to_scalar(spy["RSI_14"].iloc[-1])) if pd.notna(to_scalar(spy["RSI_14"].iloc[-1])) else None
        spy_rsi_ma = float(to_scalar(spy["RSI_MA_14"].iloc[-1])) if pd.notna(to_scalar(spy["RSI_MA_14"].iloc[-1])) else None

        spy_dist_200 = None
        if spy_sma_200 not in [None, 0]:
            spy_dist_200 = ((spy_price / spy_sma_200) - 1) * 100

        spy_rsi_dist = None
        if spy_rsi is not None and spy_rsi_ma not in [None, 0]:
            spy_rsi_dist = ((spy_rsi / spy_rsi_ma) - 1) * 100

        vix_value = None
        if not vix.empty and pd.notna(to_scalar(vix["Close"].iloc[-1])):
            vix_value = float(to_scalar(vix["Close"].iloc[-1]))

        return {
            "SPY_price": round(spy_price, 2),
            "SPY_dist_200": round(spy_dist_200, 2) if spy_dist_200 is not None else None,
            "SPY_RSI": round(spy_rsi, 2) if spy_rsi is not None else None,
            "SPY_RSI_dist_MA14": round(spy_rsi_dist, 2) if spy_rsi_dist is not None else None,
            "VIX": round(vix_value, 2) if vix_value is not None else None,
        }
    except Exception as e:
        return {
            "error": str(e),
            "SPY_price": None,
            "SPY_dist_200": None,
            "SPY_RSI": None,
            "SPY_RSI_dist_MA14": None,
            "VIX": None,
        }


def market_status_spy_dist(value, overbought=10.0, oversold=-10.0):
    if value is None:
        return "", ""
    if value > overbought:
        return f"{value:.2f}% + overbought", "#f8d7da"
    if value < oversold:
        return f"{value:.2f}% + oversold", "#d4edda"
    return f"{value:.2f}% + neutral", "#fff3cd"


def market_status_rsi(value, overbought=69.0, oversold=35.0):
    if value is None:
        return "", ""
    if value > overbought:
        return f"{value:.2f} + overbought", "#f8d7da"
    if value < oversold:
        return f"{value:.2f} + oversold", "#d4edda"
    return f"{value:.2f} + neutral", "#fff3cd"


def market_status_rsi_dist(value, overbought=20.0, oversold=-20.0):
    if value is None:
        return "", ""
    if value > overbought:
        return f"{value:.2f}% + overbought", "#f8d7da"
    if value < oversold:
        return f"{value:.2f}% + oversold", "#d4edda"
    return f"{value:.2f}% + neutral", "#fff3cd"


def market_status_vix(value, fear=30.0, no_fear=20.0):
    if value is None:
        return "", ""
    if value > fear:
        return f"{value:.2f} + Fear", ""
    if value < no_fear:
        return f"{value:.2f} + No Fear", ""
    return f"{value:.2f} + Mid Fear", ""


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


market_tab, stocks_tab = st.tabs(["Market Condition", "Stocks & Analysis"])

with market_tab:
    st.subheader("General Market Condition")

    col1, col2 = st.columns(2)

    with col1:
        spy_dist_overbought = st.number_input("SPY dist from 200SMA: overbought above (%)", value=10.0)
        spy_dist_oversold = st.number_input("SPY dist from 200SMA: oversold below (%)", value=-10.0)
        rsi_overbought = st.number_input("RSI: overbought above", value=69.0)
        rsi_oversold = st.number_input("RSI: oversold below", value=35.0)

    with col2:
        rsi_dist_overbought = st.number_input("RSI dist from 14MA: overbought above (%)", value=20.0)
        rsi_dist_oversold = st.number_input("RSI dist from 14MA: oversold below (%)", value=-20.0)
        vix_fear = st.number_input("VIX: Fear above", value=30.0)
        vix_no_fear = st.number_input("VIX: No Fear below", value=20.0)

    market_data = get_market_snapshot()

    if "error" in market_data and market_data["error"]:
        st.error(f"Market data error: {market_data['error']}")

    spy_dist_text, spy_dist_color = market_status_spy_dist(
        market_data["SPY_dist_200"],
        overbought=spy_dist_overbought,
        oversold=spy_dist_oversold
    )

    rsi_text, rsi_color = market_status_rsi(
        market_data["SPY_RSI"],
        overbought=rsi_overbought,
        oversold=rsi_oversold
    )

    rsi_dist_text, rsi_dist_color = market_status_rsi_dist(
        market_data["SPY_RSI_dist_MA14"],
        overbought=rsi_dist_overbought,
        oversold=rsi_dist_oversold
    )

    vix_text, _ = market_status_vix(
        market_data["VIX"],
        fear=vix_fear,
        no_fear=vix_no_fear
    )

    market_rows = [
        ("SPY distance from 200SMA", spy_dist_text, spy_dist_color),
        ("RSI", rsi_text, rsi_color),
        ("RSI distance from 14MA", rsi_dist_text, rsi_dist_color),
        ("VIX", vix_text, ""),
    ]

    rows_html = ""
    for metric, value, color in market_rows:
        style = f'background-color:{color};' if color else ""
        rows_html += f"""
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
            font-size: 14px;
        }}
        .market-table th, .market-table td {{
            border: 1px solid #ddd;
            padding: 10px;
            text-align: left;
            vertical-align: top;
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
                {rows_html}
            </tbody>
        </table>
        """,
        unsafe_allow_html=True
    )

with stocks_tab:
    # ---------- First page: stock master table ----------
    try:
        master_df = pd.read_csv("stocks_master.csv", encoding="latin1", engine="python")
        master_df.columns = master_df.columns.str.strip()

        if "Industruy" in master_df.columns:
            master_df = master_df.rename(columns={"Industruy": "Industry"})

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

    st.divider()

    default_tickers = ",".join(master_df["Ticker"].dropna().astype(str).tolist()) if not master_df.empty else ""

    st.subheader("Run Analysis")

    tickers_text = st.text_area(
        "Tickers (comma separated)",
        value=default_tickers,
        height=140
    )

    st.subheader("Parameters")

    buy_rsi_dist_max = st.number_input("BUY: RSI distance max (%)", value=-10.0, key="buy_rsi_dist_max")
    buy_dist_200_max = st.number_input("BUY: Price dist from 200SMA max (%)", value=10.0, key="buy_dist_200_max")
    buy_dist_150_max = st.number_input("BUY: Price dist from 150SMA max (%)", value=5.0, key="buy_dist_150_max")
    sell_rsi_min = st.number_input("SELL: RSI min", value=70.0, key="sell_rsi_min")
    sell_dist_150_min = st.number_input("SELL: Price dist from 150SMA min (%)", value=50.0, key="sell_dist_150_min")

    run_button = st.button("Run analysis", use_container_width=True)

    if run_button:
        tickers = [t.strip().upper() for t in tickers_text.replace("\n", "").split(",") if t.strip()]

        with st.spinner("Running analysis..."):
            df = build_output(
                tickers,
                buy_rsi_dist_max=buy_rsi_dist_max,
                buy_dist_200_max=buy_dist_200_max,
                buy_dist_150_max=buy_dist_150_max,
                sell_rsi_min=sell_rsi_min,
                sell_dist_150_min=sell_dist_150_min
            )

        st.success("Analysis complete")

        summary_df = df[[
            "symbol",
            "risk_level",
            "Strong_vs_SPY",
            "BUY/SELL signal"
        ]].copy()
        summary_df.columns = ["Name", "Risk_level", "strong_vs_SPY", "BUY/SELL"]

        st.subheader("Summary")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

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

        temp_excel_name = "risk_analysis_streamlit.xlsx"
        temp_txt_name = "risk_analysis_streamlit.txt"

        save_outputs(
            df,
            excel_file=temp_excel_name,
            txt_file=temp_txt_name,
            buy_rsi_dist_max=buy_rsi_dist_max,
            buy_dist_200_max=buy_dist_200_max,
            buy_dist_150_max=buy_dist_150_max,
            sell_rsi_min=sell_rsi_min,
            sell_dist_150_min=sell_dist_150_min
        )

        with open(temp_excel_name, "rb") as f:
            excel_bytes = f.read()

        st.download_button(
            "Download Excel",
            data=excel_bytes,
            file_name="risk_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )