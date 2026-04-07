import streamlit as st
import pandas as pd
from engine import build_output, save_outputs

st.set_page_config(page_title="Stock Analyzer", layout="centered")

st.title("Stock Analyzer")

# ---------- First page: stock master table ----------
try:
    master_df = pd.read_csv("stocks_master.csv")
except Exception:
    master_df = pd.DataFrame(columns=["Ticker", "Company", "Industry", "Catalyst"])

st.subheader("Stock List")

if not master_df.empty:
    stock_list_view = master_df[["Ticker", "Company", "Industry", "Catalyst"]].copy()
    st.dataframe(stock_list_view, use_container_width=True, hide_index=True)
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

    summary_df = df[[
        "symbol",
        "RSI_dist_%_from_RSI_MA_14",
        "Price_dist_%_from_SMA_200",
        "BUY/SELL signal"
    ]].copy()
    summary_df.columns = ["Name", "dist RSI", "dist 200SMA", "BUY/SELL"]

    st.subheader("Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    with st.expander("Show full analysis table"):
        st.dataframe(df, use_container_width=True, hide_index=True)

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