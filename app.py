import streamlit as st
import pandas as pd
import html
from engine import build_output, save_outputs

st.set_page_config(page_title="Stock Analyzer", layout="centered")

st.title("Stock Analyzer")

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
    else:
        return f"<b>{first_line}</b>"

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

# ---------- Analysis section ----------
default_tickers = ",".join(master_df["Ticker"].dropna().astype(str).tolist()) if not master_df.empty else ""

st.subheader("Run Analysis")

tickers_text = st.text_area(
    "Tickers (comma separated)",
    value=default_tickers,
    height=140
)

st.subheader("Parameters")

buy_rsi_dist_max = st.number_input("BUY: RSI distance max (%)", value=-10.0)
buy_dist_200_max = st.number_input("BUY: Price dist from 200SMA max (%)", value=10.0)
buy_dist_150_max = st.number_input("BUY: Price dist from 150SMA max (%)", value=5.0)
sell_rsi_min = st.number_input("SELL: RSI min", value=70.0)
sell_dist_150_min = st.number_input("SELL: Price dist from 150SMA min (%)", value=50.0)

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

    # Summary table: Name, Risk_level, strong_vs_SPY, BUY/SELL
    summary_df = df[[
        "symbol",
        "risk_level",
        "Strong_vs_SPY",
        "BUY/SELL signal"
    ]].copy()
    summary_df.columns = ["Name", "Risk_level", "strong_vs_SPY", "BUY/SELL"]

    st.subheader("Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # Full table without Risk_level and Strong_vs_SPY
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