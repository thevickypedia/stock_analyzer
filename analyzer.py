import pandas as pd
import numpy as np
from fetcher import nasdaq
import logging
import xlsxwriter


class Analyzer:
    def __init__(self):
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s %(name)s %(levelname)s %(message)s')
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        self.stocks = nasdaq()
        self.workbook = xlsxwriter.Workbook('stocks.xlsx')
        self.worksheet = self.workbook.add_worksheet('Results')
        self.worksheet.write(0, 0, "Stock Ticker")
        self.worksheet.write(0, 1, "Capital")
        self.worksheet.write(0, 2, "PE Ratio")
        self.worksheet.write(0, 3, "Yield")

    def write(self):
        n = 0
        logging.info('Initializing Analysis on all NASDAQ stocks')
        for stock in self.stocks:
            url = f'https://finance.yahoo.com/quote/{stock}/'
            try:
                sheet = pd.read_html(url, flavor='bs4')[-1]
                if 'N/A (N/A)' not in list(sheet[1]) and np.nan not in list(sheet[1]):
                    n = n + 1
                    market_capital = sheet.iat[0, 1]
                    pe_ratio = sheet.iat[2, 1]
                    forward_dividend_yield = sheet.iat[5, 1]
                    self.worksheet.write(n, 0, f'{stock}')
                    self.worksheet.write(n, 1, f'{market_capital}')
                    self.worksheet.write(n, 2, f'{pe_ratio}')
                    self.worksheet.write(n, 3, f'{forward_dividend_yield}')
                else:
                    logging.warning(f'Received null values on analysis for {stock}')
            except KeyboardInterrupt:
                logging.error('Terminating session and saving the workbook')
                self.workbook.close()
                exit(0)
            except:
                logging.debug(f'Unable to analyze {stock}')

        self.workbook.close()


if __name__ == '__main__':
    Analyzer().write()
