import pandas as pd
import math
from yahooquery import Ticker
from datetime import datetime

# Import stocks list
url = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
data = pd.read_html(url)
spy_companies = data[0].iloc[:, [0, 1, 3, 4, 5, 6, 7, 8]]
#print(spy_companies)

# Loop through tickers and create new dataframe
ticker_str = " ".join(spy_companies['Symbol'])
#print(ticker_str)
stocks_info = Ticker(ticker_str)
stocks_dict = stocks_info.price
#print(stocks_dict)

# Initialize dataframe columns
cols = ['Ticker', 'Name', 'Stock Price', 'Market Capitalization', '# Shares to Buy']
stocks_df = pd.DataFrame(columns=cols)

# Fill in dataframe rows
for stock in spy_companies['Symbol']:
    # Get full stock name
    if stocks_dict[stock].get('longName') == None or stocks_dict[stock].get('longName') == "None":
        stock_name = "N/A"
    else:
        stock_name = stocks_dict[stock]['longName']

    # Get latest stock price
    if stocks_dict[stock].get('regularMarketPrice') == None or stocks_dict[stock].get('regularMarketPrice') == "None":
        stock_price = "N/A"
    else:
        stock_price = stocks_dict[stock]['regularMarketPrice']

    # Get latest market capitalization
    if stocks_dict[stock].get('marketCap') == None or stocks_dict[stock].get('marketCap') == "None":
        stock_price = "N/A"
    else:
        market_cap = stocks_dict[stock]['marketCap']

    # Append row to dataframe
    stocks_df.loc[len(stocks_df)] = [stock, stock_name, stock_price, market_cap, 'N/A']
    #print(stock, stock_name, stock_price, market_cap)

# Sort dataframe by ticker
stocks_df = stocks_df.sort_values('Ticker') 
# print(stocks_df)

# Get users portfolio size
while True:
    try:
        portfolio_size = input('Enter the amount of your portfolio: ')
        portfolio_size = float(portfolio_size)
        break
    except ValueError:
        print("Error: That's not a valid number, try again")
        continue

# Calculate buying power
position_size = portfolio_size / len(stocks_df.index)
for i in range(0, len(stocks_df.index)):
    if stocks_df.loc[i, 'Stock Price'] != "N/A":
        stocks_df.loc[i, '# Shares to Buy'] = math.floor(portfolio_size / stocks_df.loc[i, 'Stock Price'])

print(stocks_df)

# Initialize excel output
completed_time = "_".join(str(datetime.now())[:-7].replace(':', '-').split())
#print(completed_time)
writer = pd.ExcelWriter(
    './equal_weight_index_fund/buying_power_stats/stockBuyPwr_' + completed_time, engine = 'xlsxwriter'
)
stocks_df.to_excel(writer, 'stockBuyPwr_' + completed_time, index = False)

# Initialize excel sheet colors and input formats
background_color = '#243447'
font_color = 'ffffff'

string_format = writer.book.add_format({
    'font_color': font_color, 'bg_color': background_color, 'border': 1
})
dollar_format = writer.book.add_format({
    'num_format': '$0.00','font_color': font_color, 'bg_color': background_color, 'border': 1
})
integer_format = writer.book.add_format({
    'num_format': '0', 'font_color': font_color, 'bg_color': background_color, 'border': 1
})

# Apply input formats
column_formats = {
    'A': ['Ticker', 15, string_format], 'B': ['Name', 65, string_format], 
    'C': ['Stock Price', 15, dollar_format], 'D': ['Market Capitalization', 20, integer_format], 
    'E': ['# Shares to Buy', 15, integer_format]
}

# Add purchasing power indicator
writer.sheets['stockBuyPwr_' + completed_time].set_column('G:G', 15)
writer.sheets['stockBuyPwr_' + completed_time].write('G3', 'Purchasing Power', string_format)
writer.sheets['stockBuyPwr_' + completed_time].write('G4', portfolio_size, dollar_format)

# Apply visual formats to columns
for col in column_formats.keys():
    writer.sheets['stockBuyPwr_' + completed_time].set_column(f'{col}:{col}', column_formats[col][1], column_formats[col][2])
    writer.sheets['stockBuyPwr_' + completed_time].write(f'{col}1', column_formats[col][0], column_formats[col][2])
writer.save()
