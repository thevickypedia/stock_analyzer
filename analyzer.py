import pandas as pd
import numpy as np
from lib.helper_functions import nasdaq, logger
import xlsxwriter
import time
from datetime import datetime

start_time = time.time()


class Analyzer:
    def __init__(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        self.stocks = nasdaq()
        filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet('Results')
        self.worksheet.write(0, 0, "Stock Ticker")
        self.worksheet.write(0, 1, "Capital")
        self.worksheet.write(0, 2, "PE Ratio")
        self.worksheet.write(0, 3, "Yield")

    def write(self):
        n = 0
        logger.info('Initializing Analysis on all NASDAQ stocks')
        print('Initializing Analysis on all NASDAQ stocks..')
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
                    logger.warning(f'Received null values for analysis on {stock}')
            except KeyboardInterrupt:
                logger.error('Manual Override: Terminating session and saving the workbook')
                print('Manual Override: Terminating session and saving the workbook')
                self.workbook.close()
                exec_time = self.time_converter(round(time.time() - start_time))
                logger.info(f'Total execution time: {exec_time}')
                logger.info(f'Stocks Analyzed: {n}')
                print(f'Total execution time: {exec_time}')
                print(f'Stocks Analyzed: {n}')
                exit(0)
            except:
                logger.debug(f'Unable to analyze {stock}')

        self.workbook.close()
        return round(time.time() - start_time), n

    def time_converter(self, seconds):
        seconds = seconds % (24 * 3600)
        hour = seconds // 3600
        seconds %= 3600
        minutes = seconds // 60
        seconds %= 60
        if hour:
            return f'{hour} hours {minutes} minutes {seconds} seconds'
        elif minutes:
            return f'{minutes} minutes {seconds} seconds'
        elif seconds:
            return f'{seconds} seconds'


if __name__ == '__main__':
    timed_response, no = Analyzer().write()
    time_taken = Analyzer().time_converter(timed_response)
    logger.info(f'Total execution time: {time_taken}')
    logger.info(f'Stocks Analyzed: {no}')
    print(f'Total execution time: {time_taken}')
    print(f'Stocks Analyzed: {no}')
