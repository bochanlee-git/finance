import os
import smtplib
import yfinance as yf
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


TICKERS = ["WELL", "EQIX", "PLD", "TSCO", "MSCI", "AVGO", "ASML", "UNH", "MDT", "NEE", "MCD", "XOM", "WMT", "SYY", "SWK", "SPGI", "PPG", "NUE", "NDSN", "LEG", "KMB", "DOV", "GWW", "GPC", "FNF", "EMR", "CL", "CINF", "BDX", "ABT", "AMT", "HRL", "ABBV", "ADP", "MO", "PEP", "MMM", "JNJ", "DIVO", "O", "JEPQ", "JEPI", "SCHD", "BAC", "TGT", "ITW", "LOW", "PG", "JPM", "KO"]

def calc_rsi(close, period=14):
    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)

    avg_gain = gain.rolling(period, min_periods=period).mean()
    avg_loss = loss.rolling(period, min_periods=period).mean()

    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))


def calc_mfi(high, low, close, volume, period=14):
    typical_price = (high + low + close) / 3
    money_flow = typical_price * volume

    direction = typical_price.diff()

    positive_flow = money_flow.where(direction > 0, 0)
    negative_flow = money_flow.where(direction < 0, 0)

    positive_mf = positive_flow.rolling(period, min_periods=period).sum()
    negative_mf = negative_flow.rolling(period, min_periods=period).sum()

    return 100 - (100 / (1 + positive_mf / negative_mf))


def classify_stock(rsi, mfi):
    if rsi <= 30 and mfi <= 20:
        return "A"
    elif rsi <= 30 and mfi > 20:
        return "B"
    elif rsi >= 70 and mfi >= 80:
        return "C"
    elif rsi >= 70 and mfi < 80:
        return "D"
    else:
        return "E"


def analyze_tickers(tickers):
    results = []

    for ticker in tickers:
        try:
            df = yf.download(
                ticker,
                period="6mo",
                interval="1d",
                progress=False,
                auto_adjust=False
            )

            if df.empty:
                raise ValueError("No data")

            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)

            df["RSI"] = calc_rsi(df["Close"])
            df["MFI"] = calc_mfi(df["High"], df["Low"], df["Close"], df["Volume"])

            latest = df[["RSI", "MFI"]].dropna().iloc[-1]

            rsi = float(latest["RSI"])
            mfi = float(latest["MFI"])
            score = (rsi * 0.7) + (mfi * 0.3)

            results.append({
                "Ticker": ticker,
                "RSI": round(rsi, 2),
                "MFI": round(mfi, 2),
                "Score": round(score, 2),
                "Group": classify_stock(rsi, mfi)
            })

        except Exception as e:
            print(f"{ticker} error: {e}")

            results.append({
                "Ticker": ticker,
                "RSI": None,
                "MFI": None,
                "Score": None,
                "Group": f"Error: {e}"
            })

    result_df = pd.DataFrame(results)
    result_df = result_df.sort_values(by="Score", ascending=True, na_position="last")

    return result_df


def send_email(result_df):
    sender = os.environ["EMAIL_SENDER"]
    password = os.environ["EMAIL_PASSWORD"]
    receiver = os.environ["EMAIL_RECEIVER"]

    subject = "Daily RSI/MFI Stock Report"

    html_table = result_df.to_html(index=False)

    body = f"""
    <h2>Daily RSI/MFI Stock Report</h2>
    <p>Score = RSI × 0.7 + MFI × 0.3</p>
    {html_table}
    """

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.send_message(msg)


if __name__ == "__main__":
    result = analyze_tickers(TICKERS)
    print(result)
    send_email(result)
