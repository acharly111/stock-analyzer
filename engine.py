import yfinance as yf
import pandas as pd
from openpyxl import load_workbook


def calculate_atr(data, period=14):
    high = data["High"]
    low = data["Low"]
    close = data["Close"]

    tr1 = high - low
    tr2 = (high - close.shift()).abs()
    tr3 = (low - close.shift()).abs()

    true_range = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    atr = true_range.ewm(alpha=1 / period, min_periods=period).mean()
    return atr


def classify_risk(atr_pct):
    if atr_pct >= 0.06:
        return "Very High"
    elif atr_pct >= 0.045:
        return "High"
    elif atr_pct >= 0.03:
        return "Mid"
    elif atr_pct >= 0.015:
        return "Low"
    else:
        return "Very Low"


def to_scalar(x):
    if isinstance(x, pd.Series):
        return x.iloc[0]
    return x


def get_close_on_date(data, target_date):
    target_date = pd.Timestamp(target_date)

    if data.empty:
        return None

    matches = data.loc[data.index.normalize() == target_date]
    if matches.empty:
        return None

    value = matches["Close"].iloc[-1]
    value = to_scalar(value)

    if pd.isna(value):
        return None

    return round(float(value), 2)


def calc_pct_change(start_value, end_value):
    if start_value is None or end_value is None:
        return None
    if start_value == 0:
        return None
    return round(((end_value / start_value) - 1) * 100, 2)


def analyze_stocks(tickers):
    results = []

    for ticker in tickers:
        try:
            data = yf.download(
                ticker,
                period="1y",
                interval="1d",
                progress=False,
                auto_adjust=False,
                group_by="column"
            )

            if data.empty:
                results.append({
                    "symbol": ticker,
                    "price": None,
                    "ATR_14": None,
                    "ATR_pct": None,
                    "risk_level": "No Data",
                    "Close_2026_03_09": None,
                    "Close_2026_03_30": None,
                    "Close_2026_04_02": None,
                    "Change_%_03_09_to_04_02": None,
                    "Change_%_03_30_to_04_02": None,
                    "Strong_vs_SPY": None,
                    "SMA_200": None,
                    "SMA_150": None,
                    "RSI_14": None,
                    "RSI_MA_14": None,
                    "RSI_dist_%_from_RSI_MA_14": None,
                    "Price_dist_%_from_SMA_150": None,
                    "Price_dist_%_from_SMA_200": None,
                    "BUY/SELL signal": None,
                })
                continue

            if isinstance(data.columns, pd.MultiIndex):
                data.columns = data.columns.get_level_values(0)

            data["ATR_14"] = calculate_atr(data)

            data["SMA_200"] = data["Close"].rolling(200).mean()
            data["SMA_150"] = data["Close"].rolling(150).mean()

            delta = data["Close"].diff()
            gain = delta.clip(lower=0)
            loss = -delta.clip(upper=0)

            avg_gain = gain.ewm(alpha=1/14, min_periods=14).mean()
            avg_loss = loss.ewm(alpha=1/14, min_periods=14).mean()

            rs = avg_gain / avg_loss
            data["RSI_14"] = 100 - (100 / (1 + rs))
            data["RSI_MA_14"] = data["RSI_14"].rolling(14).mean()

            price = to_scalar(data["Close"].iloc[-1])
            atr = to_scalar(data["ATR_14"].iloc[-1])

            sma_200 = to_scalar(data["SMA_200"].iloc[-1])
            sma_150 = to_scalar(data["SMA_150"].iloc[-1])
            rsi_14 = to_scalar(data["RSI_14"].iloc[-1])
            rsi_ma_14 = to_scalar(data["RSI_MA_14"].iloc[-1])

            close_2026_03_09 = get_close_on_date(data, "2026-03-09")
            close_2026_03_30 = get_close_on_date(data, "2026-03-30")
            close_2026_04_02 = get_close_on_date(data, "2026-04-02")

            if pd.isna(price) or pd.isna(atr) or float(price) == 0:
                results.append({
                    "symbol": ticker,
                    "price": None,
                    "ATR_14": None,
                    "ATR_pct": None,
                    "risk_level": "No Data",
                    "Close_2026_03_09": close_2026_03_09,
                    "Close_2026_03_30": close_2026_03_30,
                    "Close_2026_04_02": close_2026_04_02,
                    "Change_%_03_09_to_04_02": None,
                    "Change_%_03_30_to_04_02": None,
                    "Strong_vs_SPY": None,
                    "SMA_200": round(float(sma_200), 2) if pd.notna(sma_200) else None,
                    "SMA_150": round(float(sma_150), 2) if pd.notna(sma_150) else None,
                    "RSI_14": round(float(rsi_14), 2) if pd.notna(rsi_14) else None,
                    "RSI_MA_14": round(float(rsi_ma_14), 2) if pd.notna(rsi_ma_14) else None,
                    "RSI_dist_%_from_RSI_MA_14": None,
                    "Price_dist_%_from_SMA_150": None,
                    "Price_dist_%_from_SMA_200": None,
                    "BUY/SELL signal": None,
                })
                continue

            price = float(price)
            atr = float(atr)
            atr_pct = atr / price
            risk = classify_risk(atr_pct)

            results.append({
                "symbol": ticker,
                "price": round(price, 2),
                "ATR_14": round(atr, 2),
                "ATR_pct": round(atr_pct * 100, 2),
                "risk_level": risk,
                "Close_2026_03_09": close_2026_03_09,
                "Close_2026_03_30": close_2026_03_30,
                "Close_2026_04_02": close_2026_04_02,
                "Change_%_03_09_to_04_02": None,
                "Change_%_03_30_to_04_02": None,
                "Strong_vs_SPY": None,
                "SMA_200": round(float(sma_200), 2) if pd.notna(sma_200) else None,
                "SMA_150": round(float(sma_150), 2) if pd.notna(sma_150) else None,
                "RSI_14": round(float(rsi_14), 2) if pd.notna(rsi_14) else None,
                "RSI_MA_14": round(float(rsi_ma_14), 2) if pd.notna(rsi_ma_14) else None,
                "RSI_dist_%_from_RSI_MA_14": None,
                "Price_dist_%_from_SMA_150": None,
                "Price_dist_%_from_SMA_200": None,
                "BUY/SELL signal": None,
            })

        except Exception:
            results.append({
                "symbol": ticker,
                "price": None,
                "ATR_14": None,
                "ATR_pct": None,
                "risk_level": "Error",
                "Close_2026_03_09": None,
                "Close_2026_03_30": None,
                "Close_2026_04_02": None,
                "Change_%_03_09_to_04_02": None,
                "Change_%_03_30_to_04_02": None,
                "Strong_vs_SPY": None,
                "SMA_200": None,
                "SMA_150": None,
                "RSI_14": None,
                "RSI_MA_14": None,
                "RSI_dist_%_from_RSI_MA_14": None,
                "Price_dist_%_from_SMA_150": None,
                "Price_dist_%_from_SMA_200": None,
                "BUY/SELL signal": None,
            })

    return pd.DataFrame(results)


def get_spy_row():
    try:
        data = yf.download(
            "SPY",
            period="1y",
            interval="1d",
            progress=False,
            auto_adjust=False,
            group_by="column"
        )

        if data.empty:
            return {
                "symbol": "SPY",
                "price": None,
                "ATR_14": None,
                "ATR_pct": None,
                "risk_level": "Benchmark",
                "Close_2026_03_09": None,
                "Close_2026_03_30": None,
                "Close_2026_04_02": None,
                "Change_%_03_09_to_04_02": None,
                "Change_%_03_30_to_04_02": None,
                "Strong_vs_SPY": None,
                "SMA_200": None,
                "SMA_150": None,
                "RSI_14": None,
                "RSI_MA_14": None,
                "RSI_dist_%_from_RSI_MA_14": None,
                "Price_dist_%_from_SMA_150": None,
                "Price_dist_%_from_SMA_200": None,
                "BUY/SELL signal": None,
            }

        if isinstance(data.columns, pd.MultiIndex):
            data.columns = data.columns.get_level_values(0)

        data["SMA_200"] = data["Close"].rolling(200).mean()
        data["SMA_150"] = data["Close"].rolling(150).mean()

        delta = data["Close"].diff()
        gain = delta.clip(lower=0)
        loss = -delta.clip(upper=0)

        avg_gain = gain.ewm(alpha=1/14, min_periods=14).mean()
        avg_loss = loss.ewm(alpha=1/14, min_periods=14).mean()

        rs = avg_gain / avg_loss
        data["RSI_14"] = 100 - (100 / (1 + rs))
        data["RSI_MA_14"] = data["RSI_14"].rolling(14).mean()

        close_2026_03_09 = get_close_on_date(data, "2026-03-09")
        close_2026_03_30 = get_close_on_date(data, "2026-03-30")
        close_2026_04_02 = get_close_on_date(data, "2026-04-02")

        last_price = to_scalar(data["Close"].iloc[-1])
        last_price = round(float(last_price), 2) if not pd.isna(last_price) else None

        sma_200 = to_scalar(data["SMA_200"].iloc[-1])
        sma_150 = to_scalar(data["SMA_150"].iloc[-1])
        rsi_14 = to_scalar(data["RSI_14"].iloc[-1])
        rsi_ma_14 = to_scalar(data["RSI_MA_14"].iloc[-1])

        return {
            "symbol": "SPY",
            "price": last_price,
            "ATR_14": None,
            "ATR_pct": None,
            "risk_level": "Benchmark",
            "Close_2026_03_09": close_2026_03_09,
            "Close_2026_03_30": close_2026_03_30,
            "Close_2026_04_02": close_2026_04_02,
            "Change_%_03_09_to_04_02": None,
            "Change_%_03_30_to_04_02": None,
            "Strong_vs_SPY": None,
            "SMA_200": round(float(sma_200), 2) if pd.notna(sma_200) else None,
            "SMA_150": round(float(sma_150), 2) if pd.notna(sma_150) else None,
            "RSI_14": round(float(rsi_14), 2) if pd.notna(rsi_14) else None,
            "RSI_MA_14": round(float(rsi_ma_14), 2) if pd.notna(rsi_ma_14) else None,
            "RSI_dist_%_from_RSI_MA_14": None,
            "Price_dist_%_from_SMA_150": None,
            "Price_dist_%_from_SMA_200": None,
            "BUY/SELL signal": None,
        }

    except Exception:
        return {
            "symbol": "SPY",
            "price": None,
            "ATR_14": None,
            "ATR_pct": None,
            "risk_level": "Benchmark",
            "Close_2026_03_09": None,
            "Close_2026_03_30": None,
            "Close_2026_04_02": None,
            "Change_%_03_09_to_04_02": None,
            "Change_%_03_30_to_04_02": None,
            "Strong_vs_SPY": None,
            "SMA_200": None,
            "SMA_150": None,
            "RSI_14": None,
            "RSI_MA_14": None,
            "RSI_dist_%_from_RSI_MA_14": None,
            "Price_dist_%_from_SMA_150": None,
            "Price_dist_%_from_SMA_200": None,
            "BUY/SELL signal": None,
        }


def fill_gui_columns(
    df,
    buy_rsi_dist_max=-10,
    buy_dist_200_max=10,
    buy_dist_150_max=5,
    sell_rsi_min=70,
    sell_dist_150_min=50
):
    df = df.copy()

    # Python-calculated versions of the Excel formula columns
    df["Change_%_03_09_to_04_02"] = df.apply(
        lambda row: calc_pct_change(row["Close_2026_03_09"], row["Close_2026_04_02"]),
        axis=1
    )

    df["Change_%_03_30_to_04_02"] = df.apply(
        lambda row: calc_pct_change(row["Close_2026_03_30"], row["Close_2026_04_02"]),
        axis=1
    )

    df["RSI_dist_%_from_RSI_MA_14"] = df.apply(
        lambda row: round(((row["RSI_14"] / row["RSI_MA_14"]) - 1) * 100, 2)
        if pd.notna(row["RSI_14"]) and pd.notna(row["RSI_MA_14"]) and row["RSI_MA_14"] not in [0, None]
        else None,
        axis=1
    )

    df["Price_dist_%_from_SMA_150"] = df.apply(
        lambda row: round(((row["price"] / row["SMA_150"]) - 1) * 100, 2)
        if pd.notna(row["price"]) and pd.notna(row["SMA_150"]) and row["SMA_150"] not in [0, None]
        else None,
        axis=1
    )

    df["Price_dist_%_from_SMA_200"] = df.apply(
        lambda row: round(((row["price"] / row["SMA_200"]) - 1) * 100, 2)
        if pd.notna(row["price"]) and pd.notna(row["SMA_200"]) and row["SMA_200"] not in [0, None]
        else None,
        axis=1
    )

    # Strong_vs_SPY
    spy_rows = df[df["symbol"] == "SPY"]
    if not spy_rows.empty:
        spy_change_1 = spy_rows.iloc[0]["Change_%_03_09_to_04_02"]
        spy_change_2 = spy_rows.iloc[0]["Change_%_03_30_to_04_02"]

        def strong_vs_spy(row):
            if row["symbol"] == "SPY":
                return "Benchmark"
            if pd.isna(row["Change_%_03_09_to_04_02"]) or pd.isna(row["Change_%_03_30_to_04_02"]):
                return ""
            if pd.isna(spy_change_1) or pd.isna(spy_change_2):
                return ""
            if row["Change_%_03_09_to_04_02"] > spy_change_1 and row["Change_%_03_30_to_04_02"] > spy_change_2:
                return "Strong"
            return ""

        df["Strong_vs_SPY"] = df.apply(strong_vs_spy, axis=1)

    # BUY/SELL signal
    def signal(row):
        if row["symbol"] == "SPY":
            return "Benchmark"

        buy_ok = (
            pd.notna(row["RSI_dist_%_from_RSI_MA_14"]) and
            pd.notna(row["price"]) and
            pd.notna(row["SMA_200"]) and
            pd.notna(row["Price_dist_%_from_SMA_200"]) and
            pd.notna(row["Price_dist_%_from_SMA_150"]) and
            row["RSI_dist_%_from_RSI_MA_14"] < buy_rsi_dist_max and
            row["price"] > row["SMA_200"] and
            (
                row["Price_dist_%_from_SMA_200"] < buy_dist_200_max or
                row["Price_dist_%_from_SMA_150"] < buy_dist_150_max
            )
        )

        sell_ok = (
            pd.notna(row["RSI_14"]) and
            pd.notna(row["Price_dist_%_from_SMA_150"]) and
            row["RSI_14"] >= sell_rsi_min and
            row["Price_dist_%_from_SMA_150"] > sell_dist_150_min
        )

        if buy_ok:
            return "BUY"
        if sell_ok:
            return "SELL"
        return ""

    df["BUY/SELL signal"] = df.apply(signal, axis=1)

    return df


def build_output(
    tickers,
    buy_rsi_dist_max=-10,
    buy_dist_200_max=10,
    buy_dist_150_max=5,
    sell_rsi_min=70,
    sell_dist_150_min=50
):
    df = analyze_stocks(tickers)
    spy_row = get_spy_row()
    df = pd.concat([df, pd.DataFrame([spy_row])], ignore_index=True)

    df = fill_gui_columns(
        df,
        buy_rsi_dist_max=buy_rsi_dist_max,
        buy_dist_200_max=buy_dist_200_max,
        buy_dist_150_max=buy_dist_150_max,
        sell_rsi_min=sell_rsi_min,
        sell_dist_150_min=sell_dist_150_min
    )

    return df


def apply_excel_formulas(
    excel_file,
    buy_rsi_dist_max=-10,
    buy_dist_200_max=10,
    buy_dist_150_max=5,
    sell_rsi_min=70,
    sell_dist_150_min=50
):
    wb = load_workbook(excel_file)
    ws = wb.active

    last_row = ws.max_row
    spy_row_num = last_row

    ws["S1"] = "BUY/SELL signal"

    ws["U1"] = "Parameter"
    ws["V1"] = "Value"

    ws["U2"] = "BUY: RSI distance max (%)"
    ws["V2"] = buy_rsi_dist_max

    ws["U3"] = "BUY: Price dist from 200SMA max (%)"
    ws["V3"] = buy_dist_200_max

    ws["U4"] = "BUY: Price dist from 150SMA max (%)"
    ws["V4"] = buy_dist_150_max

    ws["U5"] = "SELL: RSI min"
    ws["V5"] = sell_rsi_min

    ws["U6"] = "SELL: Price dist from 150SMA min (%)"
    ws["V6"] = sell_dist_150_min

    for r in range(2, last_row + 1):
        ws[f"I{r}"] = f'=IF(OR(F{r}="",H{r}="",F{r}=0),"",((H{r}/F{r})-1)*100)'
        ws[f"J{r}"] = f'=IF(OR(G{r}="",H{r}="",G{r}=0),"",((H{r}/G{r})-1)*100)'

        if r == spy_row_num:
            ws[f"K{r}"] = "Benchmark"
        else:
            ws[f"K{r}"] = (
                f'=IF(OR(I{r}="",J{r}="",I{spy_row_num}="",J{spy_row_num}=""),"",'
                f'IF(AND(I{r}>I{spy_row_num},J{r}>J{spy_row_num}),"Strong",""))'
            )

        ws[f"P{r}"] = f'=IF(OR(N{r}="",O{r}="",O{r}=0),"",((N{r}/O{r})-1)*100)'
        ws[f"Q{r}"] = f'=IF(OR(B{r}="",M{r}="",M{r}=0),"",((B{r}/M{r})-1)*100)'
        ws[f"R{r}"] = f'=IF(OR(B{r}="",L{r}="",L{r}=0),"",((B{r}/L{r})-1)*100)'

        if r == spy_row_num:
            ws[f"S{r}"] = "Benchmark"
        else:
            ws[f"S{r}"] = (
                f'=IF(OR(P{r}="",Q{r}="",R{r}="",N{r}="",B{r}="",L{r}="",M{r}=""),"",'
                f'IF(AND(P{r}<$V$2,B{r}>L{r},OR(R{r}<$V$3,Q{r}<$V$4)),"BUY",'
                f'IF(AND(N{r}>=$V$5,Q{r}>$V$6),"SELL","")))'
            )

    for r in range(2, last_row + 1):
        ws[f"I{r}"].number_format = '0.00'
        ws[f"J{r}"].number_format = '0.00'
        ws[f"P{r}"].number_format = '0.00'
        ws[f"Q{r}"].number_format = '0.00'
        ws[f"R{r}"].number_format = '0.00'

    for cell in ["V2", "V3", "V4", "V5", "V6"]:
        ws[cell].number_format = '0.00'

    wb.save(excel_file)


def save_outputs(
    df,
    excel_file="risk_analysis.xlsx",
    txt_file="risk_analysis.txt",
    buy_rsi_dist_max=-10,
    buy_dist_200_max=10,
    buy_dist_150_max=5,
    sell_rsi_min=70,
    sell_dist_150_min=50
):
    df.to_excel(excel_file, index=False)
    apply_excel_formulas(
        excel_file,
        buy_rsi_dist_max=buy_rsi_dist_max,
        buy_dist_200_max=buy_dist_200_max,
        buy_dist_150_max=buy_dist_150_max,
        sell_rsi_min=sell_rsi_min,
        sell_dist_150_min=sell_dist_150_min
    )

    with open(txt_file, "w", encoding="utf-8") as f:
        f.write(df.to_string(index=False))