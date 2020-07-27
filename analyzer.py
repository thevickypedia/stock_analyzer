import pandas as pd
import numpy as np
from fetcher import nasdaq
import logging
import xlsxwriter

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s %(name)s %(levelname)s %(message)s')

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
workbook = xlsxwriter.Workbook('stocks.xlsx')
worksheet = workbook.add_worksheet('Results')

stocks = nasdaq()
n = 0

logging.info('Initializing Analysis on all NASDAQ stocks')

for stock in stocks:
    n = n + 1
    url = f'https://finance.yahoo.com/quote/{stock}/'
    try:
        sheet = pd.read_html(url, flavor='bs4')[-1]
        if 'N/A (N/A)' not in list(sheet[1]) and np.nan not in list(sheet[1]):
            market_capital = sheet.iat[0, 1]
            pe_ratio = sheet.iat[2, 1]
            forward_dividend_yield = sheet.iat[5, 1]
            worksheet.write(0, 0, "Stock Ticker")
            worksheet.write(0, 1, "Capital")
            worksheet.write(0, 2, "PE Ratio")
            worksheet.write(0, 3, "Yield")
            worksheet.write(n, 0, f'{stock}')
            worksheet.write(n, 1, f'{market_capital}')
            worksheet.write(n, 2, f'{pe_ratio}')
            worksheet.write(n, 3, f'{forward_dividend_yield}')
        else:
            logging.critical(f'Received null values on analysis for {stock}')
    except:
        logging.error(f'Unable to analyze {stock}')

workbook.close()
