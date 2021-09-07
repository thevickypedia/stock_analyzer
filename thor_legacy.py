import os
import sys
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from urllib.error import HTTPError

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
from xlsxwriter import Workbook

from lib.helper_functions import nasdaq, logger

log_dir = os.path.isdir('logs')
data_dir = os.path.isdir('data')
if not log_dir:
    os.mkdir('logs')
if not data_dir:
    os.mkdir('data')


def worksheet_initializer():
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


def analyzer(stock):
    np = 0
    retries = 0
    # noinspection PyBroadException
    try:
        summary = f'{BASE_URL}/{stock}/'
        stats = f'{BASE_URL}/{stock}/key-statistics/'
        analysis = f'{BASE_URL}/{stock}/analysis/'
        r = requests.get(f'{BASE_URL}/{stock}/')

        summary_result = pd.read_html(summary, flavor='bs4')
        market_capital = summary_result[-1].iat[0, 1]
        pe_ratio = summary_result[-1].iat[2, 1]
        forward_dividend_yield = summary_result[-1].iat[5, 1]

        scrapped = BeautifulSoup(r.text, "html.parser")
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
        result = market_capital, pe_ratio, forward_dividend_yield, price, high, low, profit_margin, price_book_ratio, \
            return_on_equity, analysis_next_year, analysis_next_5_years, analysis_past_5_years
        stock_map.update({stock: result})
    except (ValueError, IndexError):
        np += 1
        logger.info(f'Unable to analyze {stock}')
    except HTTPError:
        retries += 1
        wait = retries * 30
        time.sleep(wait)
        ct = (str(threading.current_thread()))
        thread = (''.join([t for t in ct if t.isdigit()]))
        logger.info(f'WARNING: Waiting for {wait} secs for thread: {thread} on {stock}')
        stuck_thread.append(stock)
    except:
        print(f'ERROR: Unhandled Exception, Saving spreadsheet. See stacktrace below:\n'
              f'{traceback.print_exc(file=sys.stdout)}')
        writer()
        exit(1)
    return np, retries


def reprocess_threads():
    tnp = 0
    for pending in tqdm(stuck_thread, total=st_stocks, desc='Retrying Analysis', unit='stock', leave=True):
        try:
            summary = f'{BASE_URL}/{pending}/'
            stats = f'{BASE_URL}/{pending}/key-statistics/'
            analysis = f'{BASE_URL}/{pending}/analysis/'
            r = requests.get(f'{BASE_URL}/{pending}/')

            summary_result = pd.read_html(summary, flavor='bs4')
            market_capital = summary_result[-1].iat[0, 1]
            pe_ratio = summary_result[-1].iat[2, 1]
            forward_dividend_yield = summary_result[-1].iat[5, 1]

            scrapped = BeautifulSoup(r.text, "html.parser")
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
            result = market_capital, pe_ratio, forward_dividend_yield, price, high, low, profit_margin, \
                price_book_ratio, return_on_equity, analysis_next_year, analysis_next_5_years, \
                analysis_past_5_years
            stock_map.update({pending: result})
        except (ValueError, IndexError, HTTPError):
            tnp += 1
            logger.info(f'RETRY: Unable to analyze {pending}')
            pass
    return tnp


def writer():
    n = 0
    for ticker in stock_map:
        n = n + 1
        worksheet.write(n, 0, f'{ticker}')
        worksheet.write(n, 1, f'{stock_map[ticker][0]}')
        worksheet.write(n, 2, f'{stock_map[ticker][1]}')
        worksheet.write(n, 3, f'{stock_map[ticker][2]}')
        worksheet.write(n, 4, f'{stock_map[ticker][3]}')
        worksheet.write(n, 5, f'{stock_map[ticker][4]}')
        worksheet.write(n, 6, f'{stock_map[ticker][5]}')
        worksheet.write(n, 7, f'{stock_map[ticker][6]}')
        worksheet.write(n, 8, f'{stock_map[ticker][7]}')
        worksheet.write(n, 9, f'{stock_map[ticker][8]}')
        worksheet.write(n, 10, f'{stock_map[ticker][9]}')
        worksheet.write(n, 11, f'{stock_map[ticker][10]}')
        worksheet.write(n, 12, f'{stock_map[ticker][11]}')
    workbook.close()


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


if __name__ == '__main__':
    BASE_URL = 'https://finance.yahoo.com/quote'
    filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')
    workbook = Workbook(filename)
    worksheet = workbook.add_worksheet('Results')
    current_year = int(datetime.today().year)
    worksheet_initializer()
    stock_map = {}
    stuck_thread = []
    stocks = nasdaq()
    overall = len(stocks)
    logger.info('Threading initialized to analyze all NASDAQ stocks')
    print('Threading initialized to analyze all NASDAQ stocks')
    with ThreadPoolExecutor(max_workers=10) as executor:
        output = list(
            tqdm(executor.map(analyzer, stocks), total=overall, desc='Analyzing Stocks', unit='stock', leave=True))

    initial_analyzed = len(stock_map)

    retry, initial_unprocessed = 0, 0
    for ele in output:
        initial_unprocessed += (ele[0])
        retry += (ele[-1])

    st_stocks = len(stuck_thread)
    retry_unprocessed = 0
    if st_stocks:
        logger.info(f'Retrying {st_stocks} stocks that were not processed due to stuck threads.')
        print(f'\nRetrying {st_stocks} stocks that were not processed due to stuck threads.')
        retry_unprocessed = reprocess_threads()

    writer()

    analyzed = len(stock_map)

    unprocessed = initial_unprocessed + retry_unprocessed

    logger.info(f'Total Stocks looked up: {overall}')
    print(f'Total Stocks looked up: {overall}')
    logger.info(f'Stocks Analyzed: {analyzed}')
    print(f'Stocks Analyzed: {analyzed}')

    if unprocessed:
        print(f'Number of stocks failed to analyze: {unprocessed}')
        logger.info(f'Number of stocks failed to analyze: {unprocessed}')

    if retry:
        print(f'Retry count: {retry}')
        logger.info(f'Retry count: {retry}')

    retry_processed = analyzed - initial_analyzed
    if retry_processed:
        logger.info(f'Number of stocks re-processed: {retry_processed}')
        print(f'Number of stocks re-processed: {retry_processed}')

    time_taken = time_converter(round(time.perf_counter()))
    logger.info(f'Total execution time: {time_taken}')
    print(f'Total execution time: {time_taken}')
    logger.info(f'Spreadsheet stored as {filename}')
    print(f'Spreadsheet stored as {filename}')
