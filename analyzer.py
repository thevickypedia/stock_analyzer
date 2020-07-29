import os
# import sys
import time
from datetime import datetime
import requests
from bs4 import BeautifulSoup as bs

import pandas as pd
import xlsxwriter
from tqdm import tqdm

logdir = os.path.isdir('logs')
datadir = os.path.isdir('data')
if not logdir:
    os.mkdir('logs')
if not datadir:
    os.mkdir('data')

from lib.helper_functions import nasdaq, logger

start_time = time.time()
current_year = int(datetime.today().year)


class Analyzer:
    def __init__(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')
        self.workbook = xlsxwriter.Workbook(filename)
        self.worksheet = self.workbook.add_worksheet('Results')
        self.worksheet.write(0, 0, "Stock Ticker")
        self.worksheet.write(0, 1, "Capital")
        self.worksheet.write(0, 2, "PE Ratio")
        self.worksheet.write(0, 3, "Yield")
        self.worksheet.write(0, 4, "Current Price")

    def write(self):
        n = 0
        i = 0
        # total = len(stocks)
        logger.info('Initializing Analysis on all NASDAQ stocks')
        print('Initializing Analysis on all NASDAQ stocks..')
        for stock in tqdm(stocks, desc='Analyzing Stocks', unit='stock', leave=False):
            summary = f'https://finance.yahoo.com/quote/{stock}/'
            i = i + 1
            try:
                r = requests.get(f'https://finance.yahoo.com/quote/{stock}')
                scrapped = bs(r.text, "html.parser")
                raw_data = scrapped.find_all('div', {'class': 'My(6px) Pos(r) smartphone_Mt(6px)'})[0]
                price = float(raw_data.find('span').text)

                summary_result = pd.read_html(summary, flavor='bs4')
                market_capital = summary_result[-1].iat[0, 1]
                pe_ratio = summary_result[-1].iat[2, 1]
                forward_dividend_yield = summary_result[-1].iat[5, 1]

                n = n + 1

                self.worksheet.write(n, 0, f'{stock}')
                self.worksheet.write(n, 1, f'{market_capital}')
                self.worksheet.write(n, 2, f'{pe_ratio}')
                self.worksheet.write(n, 3, f'{forward_dividend_yield}')

                self.worksheet.write(n, 4, f'{price}')

            except KeyboardInterrupt:
                logger.error('Manual Override: Terminating session and saving the workbook')
                print('\nManual Override: Terminating session and saving the workbook')
                self.workbook.close()
                logger.info(f'Stocks Analyzed: {n}')
                logger.info(f'Total Stocks looked up: {i}')
                print(f'Stocks Analyzed: {n}')
                print(f'Total Stocks looked up: {i}')
                null = i - n
                if null:
                    logger.info(f'Number of stocks failed to analyze: {null}')
                    print(f'Number of stocks failed to analyze: {null}')
                exec_time = self.time_converter(round(time.time() - start_time))
                logger.info(f'Total execution time: {exec_time}')
                print(f'Total execution time: {exec_time}')
                exit(0)
            except:
                logger.debug(f'Unable to analyze {stock}')

            # display = (f'\rCurrent status: {i}/{total}\tProgress: [%s%s] %d %%' % (
            #     ('-' * int((i * 100 / total) / 100 * 30 - 1) + '>'),
            #     (' ' * (30 - len('-' * int((i * 100 / total) / 100 * 30 - 1) + '>'))), (float(i) * 100 / total)))
            # sys.stdout.write(display)
            # sys.stdout.flush()

        self.workbook.close()
        return round(time.time() - start_time), n, i

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
    stocks = nasdaq()
    timed_response, analyzed, overall = Analyzer().write()
    time_taken = Analyzer().time_converter(timed_response)
    logger.info(f'Stocks Analyzed: {analyzed}')
    logger.info(f'Total Stocks looked up: {overall}')
    print(f'\nStocks Analyzed: {analyzed}')
    print(f'Total Stocks looked up: {overall}')
    left_overs = overall - analyzed
    if left_overs:
        logger.info(f'Number of stocks failed to analyze: {left_overs}')
        print(f'Number of stocks failed to analyze: {left_overs}')
    logger.info(f'Total execution time: {time_taken}')
    print(f'Total execution time: {time_taken}')
