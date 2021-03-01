import os
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

from numerize import numerize
from tqdm import tqdm
from xlsxwriter import Workbook
from yfinance import Ticker

log_dir = os.path.isdir('logs')
data_dir = os.path.isdir('data')
if not log_dir:
    os.mkdir('logs')
if not data_dir:
    os.mkdir('data')


def worksheet_initializer():
    """Creates header in each column."""
    worksheet.write(0, 0, "Stock Ticker")
    worksheet.write(0, 1, "Stock Name")
    worksheet.write(0, 2, "Market Capital")
    worksheet.write(0, 3, "Dividend Yield")
    worksheet.write(0, 4, "PE Ratio")
    worksheet.write(0, 5, "PB Ratio")
    worksheet.write(0, 6, "Current Price")
    worksheet.write(0, 7, "Today's High")
    worksheet.write(0, 8, "Today's Low")
    worksheet.write(0, 9, "52W High")
    worksheet.write(0, 10, "52W Low")
    worksheet.write(0, 11, "5Y Dividend Yield")
    worksheet.write(0, 12, "Profit Margin")
    worksheet.write(0, 13, "Industry")
    worksheet.write(0, 14, "Employees")


def make_float(val):
    """return float value for each value received."""
    if val:
        return round(float(val), 2)
    else:
        return None


def analyzer(stock):
    """Gathers all the necessary details."""
    try:
        info = Ticker(stock).info
    except (ValueError, KeyError):
        info = None
        logger.info(f'Unable to analyze {stock}')
    if info:
        stock_name = info['shortName'] if 'shortName' in info.keys() and info['shortName'] else None
        capital = numerize.numerize(info['marketCap']) if 'marketCap' in info.keys() and info['marketCap'] else None
        dividend_yield = make_float(info['dividendYield']) if 'dividendYield' in info.keys() else None
        pe_ratio = make_float(info['forwardPE']) if 'forwardPE' in info.keys() else None
        pb_ratio = make_float(info['priceToBook']) if 'priceToBook' in info.keys() else None
        price = make_float(info['ask']) if 'ask' in info.keys() else None
        today_high = make_float(info['dayHigh']) if 'dayHigh' in info.keys() else None
        today_low = make_float(info['dayLow']) if 'dayLow' in info.keys() else None
        high_52_weeks = make_float(info['fiftyTwoWeekHigh']) if 'fiftyTwoWeekHigh' in info.keys() else None
        low_52_weeks = make_float(info['fiftyTwoWeekLow']) if 'fiftyTwoWeekLow' in info.keys() else None
        d_yield_5y = make_float(info['fiveYearAvgDividendYield']) if 'fiveYearAvgDividendYield' in info.keys() else None
        profit_margin = make_float(info['profitMargins']) if 'profitMargins' in info.keys() else None
        industry = info['industry'] if 'industry' in info.keys() else None
        employees = numerize.numerize(info['fullTimeEmployees']) if 'fullTimeEmployees' in info.keys() and \
                                                                    info['fullTimeEmployees'] else None

        result = stock_name, capital, dividend_yield, pe_ratio, pb_ratio, price, today_high, today_low, \
            high_52_weeks, low_52_weeks, d_yield_5y, profit_margin, industry, employees
        stock_map.update({stock: result})


def writer():
    """Writes the global variable value ({stock_map}) to an excel sheet."""
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
    return len(stock_map)


def time_converter(seconds):
    """Converts seconds to appropriate hours/minutes"""
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
    from lib.helper_functions import nasdaq, logger
    filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')  # creates filename with date and time
    workbook = Workbook(filename, {'strings_to_numbers': True})  # allows possible strings as numbers
    worksheet = workbook.add_worksheet('Results')  # sheet name in the workbook
    worksheet_initializer()  # initializes worksheet
    stocks = nasdaq()  # gets all the NASDAQ stock ticket values starting A to Z
    overall = len(stocks)
    stock_map = {}  # initiates stock_map as an empty dict
    logger.info('Threading initialized to analyze all NASDAQ stocks')
    print('Threading initialized to analyze all NASDAQ stocks')
    with ThreadPoolExecutor(max_workers=10) as executor:  # multi threaded to 10 workers for throttled processing
        output = list(
            tqdm(executor.map(analyzer, stocks), total=overall, desc='Analyzing Stocks', unit='stock', leave=True))
    analyzed = writer()  # gets the number of stocks analyzed as return value after writing to workbook

    # logs and prints some closure information
    logger.info(f'Total Stocks looked up: {overall}')
    print(f'Total Stocks looked up: {overall}')
    logger.info(f'Total Stocks analyzed: {analyzed}')
    print(f'Total Stocks analyzed: {analyzed}')
    logger.info(f'Total Stocks failed to analyze: {overall - analyzed}')
    print(f'Total Stocks failed to analyze: {overall - analyzed}')

    time_taken = time_converter(round(time.perf_counter()))
    logger.info(f'Total execution time: {time_taken}')
    print(f'Total execution time: {time_taken}')
    logger.info(f'Spreadsheet stored as {filename}')
    print(f'Spreadsheet stored as {filename}')
