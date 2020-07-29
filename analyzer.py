import os
# import sys
import time
from datetime import datetime

import pandas as pd
import requests
import xlsxwriter
from bs4 import BeautifulSoup as bs
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
        self.worksheet.write(0, 5, "52 Week High")
        self.worksheet.write(0, 6, "52 Week Low")
        self.worksheet.write(0, 7, "Profit Margin")
        self.worksheet.write(0, 8, "Price Book Ratio")
        self.worksheet.write(0, 9, "Return on Equity")
        self.worksheet.write(0, 10, f"{current_year + 1} Analysis")
        self.worksheet.write(0, 11, f"{current_year} - {current_year + 5} Analysis")
        self.worksheet.write(0, 12, f"{current_year - 5} - {current_year} Analysis")

    def write(self):
        n = 0
        i = 0
        # total = len(stocks)
        logger.info('Initializing Analysis on all NASDAQ stocks')
        print('Initializing Analysis on all NASDAQ stocks..')
        for stock in tqdm(stocks, desc='Analyzing Stocks', unit='stock', leave=False):
            i = i + 1
            summary = f'https://finance.yahoo.com/quote/{stock}/'
            stats = f'https://finance.yahoo.com/quote/{stock}/key-statistics'
            analysis = f'https://finance.yahoo.com/quote/{stock}/analysis'
            r = requests.get(f'https://finance.yahoo.com/quote/{stock}')
            scrapped = bs(r.text, "html.parser")
            raw_data = scrapped.find_all('div', {'class': 'My(6px) Pos(r) smartphone_Mt(6px)'})[0]
            price = float(raw_data.find('span').text)
            try:
                summary_result = pd.read_html(summary, flavor='bs4')
                market_capital = summary_result[-1].iat[0, 1]
                pe_ratio = summary_result[-1].iat[2, 1]
                forward_dividend_yield = summary_result[-1].iat[5, 1]

                stats_result = pd.read_html(stats, flavor='bs4')
                high = stats_result[0].iat[3, 1]
                low = stats_result[0].iat[4, 1]
                profit_margin = stats_result[5].iat[0, 1]
                price_book_ratio = stats_result[0].iat[6, 1]
                return_on_equity = stats_result[6].iat[1, 1]

                analysis_result = pd.read_html(analysis, flavor='bs4')
                analysis_next_year = analysis_result[-1].iat[3, 1]
                analysis_next_5_years = analysis_result[-1].iat[4, 1]
                analysis_past_5_years = analysis_result[-1].iat[5, 1]

                n = n + 1

                self.worksheet.write(n, 0, f'{stock}')
                self.worksheet.write(n, 1, f'{market_capital}')
                self.worksheet.write(n, 2, f'{pe_ratio}')
                self.worksheet.write(n, 3, f'{forward_dividend_yield}')

                self.worksheet.write(n, 4, f'{price}')

                self.worksheet.write(n, 5, f'{high}')
                self.worksheet.write(n, 6, f'{low}')
                self.worksheet.write(n, 7, f'{profit_margin}')
                self.worksheet.write(n, 8, f'{price_book_ratio}')
                self.worksheet.write(n, 9, f'{return_on_equity}')

                self.worksheet.write(n, 10, f'{analysis_next_year}')
                self.worksheet.write(n, 11, f'{analysis_next_5_years}')
                self.worksheet.write(n, 12, f'{analysis_past_5_years}')

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
