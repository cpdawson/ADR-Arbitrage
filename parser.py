import pandas as pd
import yfinance as yf

# Load the Excel file
file_path = 'ADR_data.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Define the stocks dictionary
stocks = {
    "Toyota": {"adr": "TM", "foreign": "7203.T", "ratio": 10, "currency": "USD to JPY"},
}


# Function to get ticker symbol using yfinance
def get_ticker(company_name):
    try:
        ticker = yf.Ticker(company_name)
        return ticker
    except:
        return None


# Extract necessary information and update the stocks dictionary
for index, row in df.iterrows():
    company_name = row["Company Name"]
    adr_ticker = get_ticker(company_name + " ADR")
    foreign_ticker = get_ticker(company_name)  # Searching for the foreign ticker by company name
    ratio = int(row["Ratio (ORD:DR)"].split(":")[0])
    currency = "USD to " + row["Country"]  # Assuming the currency pair based on the country

    if adr_ticker and foreign_ticker:
        adr_ticker = adr_ticker.ticker
        foreign_ticker = foreign_ticker.ticker
    else:
        adr_ticker = row["DR Ticker"]
        foreign_ticker = None

    # Update the stocks dictionary
    stocks[company_name] = {
        "adr": adr_ticker,
        "foreign": foreign_ticker,
        "ratio": ratio,
        "currency": currency
    }

# Save the updated dictionary to a new Excel file for verification
updated_stocks_df = pd.DataFrame.from_dict(stocks, orient='index')
updated_stocks_df.to_excel('updated_stocks.xlsx')

print("Stocks dictionary updated and saved to updated_stocks.xlsx")
