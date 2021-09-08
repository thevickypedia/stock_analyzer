from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from os import getpid, mkdir, path, system
from time import perf_counter
from urllib.error import HTTPError

from _curses import error
from numerize.numerize import numerize
from pick import pick
from psutil import Process
from tqdm import tqdm
from xlsxwriter import Workbook
from yfinance import Ticker

log_dir = path.isdir('logs')
data_dir = path.isdir('data')
if not log_dir:
    mkdir('logs')
if not data_dir:
    mkdir('data')


def sort_by_value(data: dict, sort: int) -> dict:
    """Sorts the dictionary.

    Args:
        data: Dictionary which has to be sorted.
        sort: Index value of the list in the ``value`` of the dictionary.

    Returns:
        dict:
        Returns the dictionary which is sorted by a particular element in the list of values.
    """
    for key, value in data.items():
        if not value[sort]:
            value[sort] = 0
            data[key] = value
    return dict(sorted(data.items(), key=lambda e: e[1][sort], reverse=True))


def columns() -> list:
    """Names of the columns that needs to be on the header in the spreadsheet.

    Returns:
        list:
        List of headers for the spreadsheet.
    """
    return [
        "Stock Ticker",
        "Stock Name",
        "Market Capital",
        "Dividend Yield",
        "PE Ratio",
        "PB Ratio",
        "Current Price",
        "Today's High",
        "Today's Low",
        "52W High",
        "52W Low",
        "5Y Dividend Yield",
        "Profit Margin",
        "Industry",
        "Employees"
    ]


def worksheet_initializer() -> None:
    """Creates header in each column."""
    titles = columns()
    worksheet.write(0, 0, titles[0])
    worksheet.write(0, 1, titles[1])
    worksheet.write(0, 2, titles[2])
    worksheet.write(0, 3, titles[3])
    worksheet.write(0, 4, titles[4])
    worksheet.write(0, 5, titles[5])
    worksheet.write(0, 6, titles[6])
    worksheet.write(0, 7, titles[7])
    worksheet.write(0, 8, titles[8])
    worksheet.write(0, 9, titles[9])
    worksheet.write(0, 10, titles[10])
    worksheet.write(0, 11, titles[11])
    worksheet.write(0, 12, titles[12])
    worksheet.write(0, 13, titles[13])
    worksheet.write(0, 14, titles[14])


def make_float(val: int or float) -> float:
    """Return float value for each value received.

    Args:
        val: Takes integer or float as an argument and converts it to a proper decimal valued float number.

    Returns:
        float:
        Rounded float value to 2 decimal points.
    """
    return round(float(val), 2)


def extract_data(data: dict) -> list:
    """Extracts the necessary information of each stock from the data received.

    Args:
        data: Takes the information of each ticker value as an argument.

    Returns:
        list:
        A list of ``Stock Name``, ``Market Capital``, ``Dividend Yield``, ``PE Ratio``, ``PB Ratio``,
        ``Current Price``, ``Today's High Price``, ``Today's Low Price``, ``52 Week High``, ``52 Week Low``,
        ``5 Year Dividend Yield``, ``Profit Margin``, ``Industry``, ``Number of Employees``
    """
    stock_name = data.get('shortName')

    cap = data.get('marketCap')
    capital = numerize(cap) if cap else None

    div_yield = data.get('dividendYield')
    dividend_yield = make_float(div_yield) if div_yield else None

    fw_pe = data.get('forwardPE')
    pe_ratio = make_float(fw_pe) if fw_pe else None

    p2b = data.get('priceToBook')
    pb_ratio = make_float(p2b) if p2b else None

    price_ = data.get('ask')
    price = make_float(price_) if price_ else price_

    high = data.get('dayHigh')
    today_high = make_float(high) if high else None

    low = data.get('dayLow')
    today_low = make_float(low) if low else None

    high_52 = data.get('fiftyTwoWeekHigh')
    high_52_weeks = make_float(high_52) if high_52 else None

    low_52 = data.get('fiftyTwoWeekLow')
    low_52_weeks = make_float(low_52) if low_52 else None

    yield_5 = data.get('fiveYearAvgDividendYield')
    d_yield_5y = make_float(yield_5) if yield_5 else None

    pm = data.get('profitMargins')
    profit_margin = make_float(pm) if pm else None

    industry = data.get('industry')

    fte = data.get('fullTimeEmployees')
    employees = numerize(fte) if fte else None

    return [stock_name, capital, dividend_yield, pe_ratio, pb_ratio, price, today_high, today_low, high_52_weeks,
            low_52_weeks, d_yield_5y, profit_margin, industry, employees]


def analyzer(stock: str) -> None:
    """Gathers all the necessary details from each stock ticker. Calls ``extract_data()`` to get specifics.

    Args:
        stock: Takes stock ticker value as argument.
    """
    global count_404, printed
    try:
        info = Ticker(stock).info
    except (ValueError, KeyError, IndexError):
        info = None
        logger.error(f'Unable to analyze {stock}')
    except HTTPError as err:
        info = None
        # A 503 response or 50% 404 response with only 20% processed requests indicates an IP range denial
        if err.code == 503 or (count_404 > 50 * overall / 100 and len(stock_map) < 20 * overall / 100):
            print(f'\nNoticing repeated 404s, which indicates an IP range denial by {"/".join(err.url.split("/")[:3])}'
                  '\nPlease wait for a while before re-running this code. '
                  'Also, reduce number of max_workers in concurrency and consider switching to a new Network ID.') if \
                not printed else None
            printed = True  # makes sure the print statement happens only once
            ThreadPoolExecutor().shutdown()  # stop future threads to avoid progress bar on screen post print
            raise ConnectionRefusedError  # handle exception so that spreadsheet is created with existing stock_map dict
        elif err.code == 404:
            count_404 += 1  # increases count_404 for future handling
            logger.error(f'Failed to analyze {stock}. Faced error code {err.code} while requesting {err.url}. '
                         f'Reason: {err.reason}.')
        else:
            logger.error(f'Failed to analyze {stock}. Faced error code {err.code} while requesting {err.url}. '
                         f'Reason: {err.reason}.')
    if info:
        stock_map.update({stock: extract_data(data=info)})


def writer(mapping_dict: dict) -> int:
    """Writes the global variable value ``{stock_map}`` to a spreadsheet.

    Args:
        mapping_dict: Takes a dictionary as argument.

    Returns:
        int:
        Returns the number of elements in the dictionary.
    """
    n = 0
    for ticker in mapping_dict:
        n = n + 1
        worksheet.write(n, 0, f'{ticker}')
        worksheet.write(n, 1, f'{mapping_dict[ticker][0]}')
        worksheet.write(n, 2, f'{mapping_dict[ticker][1]}')
        worksheet.write(n, 3, f'{mapping_dict[ticker][2]}')
        worksheet.write(n, 4, f'{mapping_dict[ticker][3]}')
        worksheet.write(n, 5, f'{mapping_dict[ticker][4]}')
        worksheet.write(n, 6, f'{mapping_dict[ticker][5]}')
        worksheet.write(n, 7, f'{mapping_dict[ticker][6]}')
        worksheet.write(n, 8, f'{mapping_dict[ticker][7]}')
        worksheet.write(n, 9, f'{mapping_dict[ticker][8]}')
        worksheet.write(n, 10, f'{mapping_dict[ticker][9]}')
        worksheet.write(n, 11, f'{mapping_dict[ticker][10]}')
        worksheet.write(n, 12, f'{mapping_dict[ticker][11]}')
    workbook.close()
    return len(mapping_dict)


def time_converter(seconds: int) -> str:
    """Converts seconds to appropriate hours/minutes.

    Args:
        seconds: Takes the number of seconds as argument.

    Returns:
        str:
        Converted seconds to human readable values.
    """
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


def get_sort_key() -> int:
    """Displays a menu to the user, and prompts to choose how the user likes to sort the spreadsheet.

    Returns:
        int:
        Returns the index value using which the sorting has to be done.
    """
    title = "Please pick a value using which you'd like to sort the spreadsheet (Hit Ctrl+C to sort by stock ticker): "
    try:
        option, index = pick(columns()[1:], title, indicator='=>', default_index=0)
        return index
    except (error, KeyboardInterrupt):
        if not (run_env := Process(getpid()).parent().name()).endswith('sh'):
            logger.error(f"You're using {run_env} to run the script. "
                         f"Either use a terminal or enable 'Emulate terminal in output console' under\n"
                         f"Edit Configurations.. -> Execution in your {run_env}.")
        else:
            logger.error(error)


if __name__ == '__main__':
    # import in _main_ so that data and logs dir are created in advance
    from lib.helper_functions import logger, nasdaq

    filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')  # creates filename with date and time
    workbook = Workbook(filename, {'strings_to_numbers': True})  # allows possible strings as numbers
    worksheet = workbook.add_worksheet('Results')  # sheet name in the workbook
    worksheet_initializer()  # initializes worksheet
    stocks = nasdaq()  # gets all the NASDAQ stock ticket values starting A to Z
    overall = len(stocks)
    stock_map = {}  # initiates stock_map as an empty dict
    count_404 = 0
    printed = False
    logger.info('Threading initialized to analyze all NASDAQ stocks')
    print('Threading initialized to analyze all NASDAQ stocks')
    try:
        with ThreadPoolExecutor(max_workers=10) as executor:  # multi threaded to 10 workers for throttled processing
            output = list(
                tqdm(executor.map(analyzer, stocks), total=overall, desc='Analyzing Stocks', unit='stock', leave=True))
    except ConnectionRefusedError:
        pass
    sort_val = get_sort_key()
    if sort_val:
        stock_map = sort_by_value(data=stock_map, sort=sort_val)
    analyzed = writer(mapping_dict=stock_map)  # gets the number of stocks analyzed after writing to workbook

    # logs and prints some closure information
    logger.info(f'Total Stocks looked up: {overall}')
    print(f'Total Stocks looked up: {overall}')
    logger.info(f'Total Stocks analyzed: {analyzed}')
    print(f'Total Stocks analyzed: {analyzed}')
    logger.info(f'Total Stocks failed to analyze: {overall - analyzed}')
    print(f'Total Stocks failed to analyze: {overall - analyzed}')

    time_taken = time_converter(round(perf_counter()))
    logger.info(f'Total execution time: {time_taken}')
    print(f'Total execution time: {time_taken}')
    logger.info(f'Spreadsheet stored as {filename}')
    print(f'Spreadsheet stored as {filename}')
    system(f'open {filename}')  # opens spreadsheet post execution
