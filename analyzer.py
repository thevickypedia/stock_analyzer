import pandas as pd
import numpy as np
from fetcher import nasdaq

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

stocks = nasdaq()
n = 0
analysis = ''

for stock in stocks:
    n = n + 1
    url = f'https://finance.yahoo.com/quote/{stock}/'

    sheet = pd.read_html(url, flavor='bs4')[-1]
    if 'N/A (N/A)' not in list(sheet[1]) and np.nan not in list(sheet[1]):
        market_capital = sheet.iat[0, 1]
        pe_ratio = sheet.iat[2, 1]
        forward_dividend_yield = sheet.iat[5, 1]
        analysis += f'{stock}\nCapital: {market_capital}\nPE Ratio: {pe_ratio}\nYield: {forward_dividend_yield}\n\n'
    else:
        print(f'Unable to get analysis for {stock}')

print(f'Total number of stocks analyzed{n}\n\n{analysis}')
