import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import pandas as pd
import requests
import xlsxwriter
from bs4 import BeautifulSoup as bs

# import threading

# from tqdm import tqdm

logdir = os.path.isdir('logs')
datadir = os.path.isdir('data')
if not logdir:
    os.mkdir('logs')
if not datadir:
    os.mkdir('data')

from lib.helper_functions import nasdaq, logger

start_time = time.time()
current_year = int(datetime.today().year)


def initializer():
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    worksheet.write(0, 0, "Stock Ticker")
    worksheet.write(0, 1, "Capital")
    worksheet.write(0, 2, "PE Ratio")
    worksheet.write(0, 3, "Yield")
    worksheet.write(0, 4, "Current Price")
    worksheet.write(0, 5, "52 Week High")
    worksheet.write(0, 6, "52 Week Low")
    worksheet.write(0, 7, "Profit Margin")
    worksheet.write(0, 8, "Price Book Ratio")
    worksheet.write(0, 9, "Return on Equity")
    worksheet.write(0, 10, f"{current_year + 1} Analysis")
    worksheet.write(0, 11, f"{current_year} - {current_year + 5} Analysis")
    worksheet.write(0, 12, f"{current_year - 5} - {current_year} Analysis")


def launcher(stocks):
    n = 0
    i = 0
    total = len(stocks)
    for stock in stocks:
        try:
            i = i + 1
            summary = f'https://finance.yahoo.com/quote/{stock}/'
            stats = f'https://finance.yahoo.com/quote/{stock}/key-statistics/'
            analysis = f'https://finance.yahoo.com/quote/{stock}/analysis/'
            r = requests.get(f'https://finance.yahoo.com/quote/{stock}/')

            summary_result = pd.read_html(summary, flavor='bs4')
            market_capital = summary_result[-1].iat[0, 1]
            pe_ratio = summary_result[-1].iat[2, 1]
            forward_dividend_yield = summary_result[-1].iat[5, 1]

            scrapped = bs(r.text, "html.parser")
            raw_data = scrapped.find_all('div', {'class': 'My(6px) Pos(r) smartphone_Mt(6px)'})[0]
            price = float(raw_data.find('span').text)

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

            worksheet.write(n, 0, f'{stock}')
            worksheet.write(n, 1, f'{market_capital}')
            worksheet.write(n, 2, f'{pe_ratio}')
            worksheet.write(n, 3, f'{forward_dividend_yield}')

            worksheet.write(n, 4, f'{price}')

            worksheet.write(n, 5, f'{high}')
            worksheet.write(n, 6, f'{low}')
            worksheet.write(n, 7, f'{profit_margin}')
            worksheet.write(n, 8, f'{price_book_ratio}')
            worksheet.write(n, 9, f'{return_on_equity}')

            worksheet.write(n, 10, f'{analysis_next_year}')
            worksheet.write(n, 11, f'{analysis_next_5_years}')
            worksheet.write(n, 12, f'{analysis_past_5_years}')

        except KeyboardInterrupt:
            logger.error('Manual Override: Terminating session and saving the workbook')
            print('Manual Override: Terminating session and saving the workbook')
            workbook.close()
            logger.info(f'Stocks Analyzed: {n}')
            logger.info(f'Total Stocks looked up: {i}')
            print(f'Stocks Analyzed: {n}')
            print(f'Total Stocks looked up: {i}')
            null = i - n
            if null:
                logger.info(f'Number of stocks failed to analyze: {null}')
                print(f'Number of stocks failed to analyze: {null}')
            exec_time = time_converter(round(time.time() - start_time))
            logger.info(f'Total execution time: {exec_time}')
            print(f'Total execution time: {exec_time}')
            quit()

        except:
            logger.debug(f'Unable to analyze {stock}')

        # print(f'Task Executed by {threading.current_thread()}')
        display = (f'\rCurrent status: {i}/{total}\tProgress: [%s%s] %d %%' % (
            ('-' * int((i * 100 / total) / 100 * 30 - 1) + '>'),
            (' ' * (30 - len('-' * int((i * 100 / total) / 100 * 30 - 1) + '>'))), (float(i) * 100 / total)))
        sys.stdout.write(display)
        sys.stdout.flush()
    return i, n


def time_converter(seconds):
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


def worker():
    stocks = nasdaq()
    try:
        logger.info('Initializing Analysis on all NASDAQ stocks')
        print('Initializing Analysis on all NASDAQ stocks..')
        executor = ThreadPoolExecutor(max_workers=120)
        tasks = executor.submit(launcher, stocks)

        output1, output2 = tasks.result()
        workbook.close()
        return output1, output2
    except KeyboardInterrupt:
        workbook.close()
        print('\nManual Interruption')
        sys.exit(0)


if __name__ == '__main__':
    filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Results')
    initializer()
    overall, analyzed = worker()
    logger.info(f'Stocks Analyzed: {analyzed}')
    logger.info(f'Total Stocks looked up: {overall}')
    print(f'\nStocks Analyzed: {analyzed}')
    print(f'Total Stocks looked up: {overall}')
    left_overs = overall - analyzed
    if left_overs:
        logger.info(f'Number of stocks failed to analyze: {left_overs}')
        print(f'Number of stocks failed to analyze: {left_overs}')
    time_taken = time_converter(round(time.time() - start_time))
    logger.info(f'Total execution time: {time_taken}')
    print(f'Total execution time: {time_taken}')
