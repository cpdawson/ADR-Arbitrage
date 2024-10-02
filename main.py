import yfinance as yf
import pandas as pd

# Define the currency pairs and fetch the exchange rates
currency_pairs = {
    "USD to JPY": "USDJPY=X",
    "USD to EUR": "USDEUR=X",
    "USD to GBP": "USDGBP=X",
    "USD to CNY": "USDCNY=X",
    "USD to CAD": "USDCAD=X",
    "USD to AUD": "USDAUD=X",
    "USD to NZD": "USDNZD=X",
    "USD to CHF": "USDCHF=X",
    "USD to SEK": "USDSEK=X",
    "USD to NOK": "USDNOK=X",
    "USD to SGD": "USDSGD=X",
    "USD to HKD": "USDHKD=X",
    "USD to KRW": "USDKRW=X",
    "USD to INR": "USDINR=X",
    "USD to MXN": "USDMXN=X",
    "USD to BRL": "USDBRL=X",
}

for pair, ticker in currency_pairs.items():
    try:
        data = yf.Ticker(ticker)
        price = data.history(period="1d")['Close'].iloc[-1]
        currency_pairs[pair] = price
    except Exception as e:
        print(f"Error fetching exchange rate for {pair}: {e}")
        currency_pairs[pair] = None

# Load the Excel file
file_path = 'updated_tickers.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Name the first column as "Company Name"
df.columns.values[0] = "Company Name"

# Initialize the stocks dictionary
stocks = {}

# Iterate through the DataFrame to construct the dictionary
for index, row in df.iterrows():
    company_name = row["Company Name"]
    adr_ticker = row["adr"]
    foreign_ticker = row["foreign"]
    ratio = row["ratio"]
    currency = row["currency"]

    # Update the stocks dictionary
    stocks[company_name] = {
        "adr": adr_ticker,
        "foreign": foreign_ticker,
        "ratio": ratio,
        "currency": currency
    }

# Fetch and convert prices
for stock, info in stocks.items():
    adr_ticker = info["adr"]
    foreign_ticker = info["foreign"]
    adr_ratio = info["ratio"]
    currency_pair = info["currency"]
    conversion_rate = currency_pairs.get(currency_pair)

    if conversion_rate is None:
        print(f"Skipping {stock} due to missing conversion rate for {currency_pair}")
        continue

    try:
        # Fetch ADR price
        adr_data = yf.Ticker(adr_ticker)
        adr_price = adr_data.history(period="1d")['Close'].iloc[-1]

        # Fetch foreign market price
        foreign_data = yf.Ticker(foreign_ticker)
        foreign_price = foreign_data.history(period="1d")['Close'].iloc[-1]

        # Convert foreign price to USD and adjust based on ADR ratio
        foreign_price_usd = foreign_price / conversion_rate
        adjusted_adr_price = foreign_price_usd * adr_ratio

        print(f"{stock} ADR ({adr_ticker}) Price: {adr_price}")
        print(f"{stock} Foreign ({foreign_ticker}) Price: {foreign_price}")
        print(f"Conversion Rate ({currency_pair}): {conversion_rate}")
        print(f"Ratio (ADR:ORD): {adr_ratio}")
        print(f"{stock} Adjusted ADR Price based on Foreign Price: {adjusted_adr_price}")
        print("------------------------------------------------")
    except IndexError:
        pass
    except Exception as e:
        pass