from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from os import getpid, mkdir, path, system
from time import perf_counter
from typing import Union
from urllib.error import HTTPError

from _curses import error
from numerize.numerize import numerize
from pick import pick
from psutil import Process
from requests.exceptions import ChunkedEncodingError, ConnectionError
from tqdm import tqdm
from urllib3.exceptions import ProtocolError
from xlsxwriter import Workbook
from yfinance import Ticker

if not path.isdir('logs'):
    mkdir('logs')
if not path.isdir('data'):
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
    column_name = headers[1:][sort]
    console_logger.info(f'Spreadsheet will be sorted by {column_name}')
    for key, value in data.items():
        if not value[sort]:
            value[sort] = 0
            data[key] = value
    reverse_flag = False if column_name == 'Rating' else True
    return dict(sorted(data.items(), key=lambda element: element[1][sort], reverse=reverse_flag))


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
        "Employees",
        "Rating"
    ]


def worksheet_initializer() -> None:
    """Creates header in each column."""
    worksheet.write(0, 0, headers[0])
    worksheet.write(0, 1, headers[1])
    worksheet.write(0, 2, headers[2])
    worksheet.write(0, 3, headers[3])
    worksheet.write(0, 4, headers[4])
    worksheet.write(0, 5, headers[5])
    worksheet.write(0, 6, headers[6])
    worksheet.write(0, 7, headers[7])
    worksheet.write(0, 8, headers[8])
    worksheet.write(0, 9, headers[9])
    worksheet.write(0, 10, headers[10])
    worksheet.write(0, 11, headers[11])
    worksheet.write(0, 12, headers[12])
    worksheet.write(0, 13, headers[13])
    worksheet.write(0, 14, headers[14])
    worksheet.write(0, 15, headers[15])


def make_float(val: int or float) -> float:
    """Return float value for each value received.

    Args:
        val: Takes integer or float as an argument and converts it to a proper decimal valued float number.

    Returns:
        float:
        Rounded float value to 2 decimal points.
    """
    return round(float(val), 2)


def extract_data(data: dict) -> Union[list, None]:
    """Extracts the necessary information of each stock from the data received.

    Args:
        data: Takes the information of each ticker value as an argument.

    Returns:
        list:
        A list of ``Stock Name``, ``Market Capital``, ``Dividend Yield``, ``PE Ratio``, ``PB Ratio``,
        ``Current Price``, ``Today's High Price``, ``Today's Low Price``, ``52 Week High``, ``52 Week Low``,
        ``5 Year Dividend Yield``, ``Profit Margin``, ``Industry``, ``Number of Employees``, ``Recommendation Rating``
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

    recommendation = data.get('recommendationMean')
    rating = float(recommendation) if recommendation else None

    stock_data = [stock_name, capital, dividend_yield, pe_ratio, pb_ratio, price, today_high, today_low, high_52_weeks,
                  low_52_weeks, d_yield_5y, profit_margin, industry, employees, rating]

    if stock_name and any(stock_data):
        return stock_data
    else:
        file_logger.error(f"Unable to extract necessary information for analyzing {data.get('symbol')}")


def analyzer(stock: str) -> None:
    """Gathers all the necessary details from each stock ticker. Calls ``extract_data()`` to get specifics.

    Args:
        stock: Takes stock ticker value as argument.

    See Also:
        - Captures the number of ``404`` responses in a global variable.
        - Exits if a ``503`` is received or if 50% requests returned a ``404``, with less than 20% of ``200`` response.
        - ``503`` response is received only during either of the following scenarios:

            - ``max_workers`` in ThreadPool is increased beyond 20.
            - Script is run repeatedly with short intervals.
    """
    global count_404, printed
    try:
        stock_data = extract_data(data=Ticker(stock).info)
        if stock_data:
            stock_map.update({stock: stock_data})
    except HTTPError as err:
        # A 503 response or 50% 404 response with only 20% processed requests indicates an IP range denial
        if err.code == 503 or (count_404 > 50 * overall / 100 and len(stock_map) < 20 * overall / 100):
            root_logger.error(f'\nNoticing repeated 404s, which indicates an IP range denial by '
                              f'{"/".join(err.url.split("/")[:3])}\nPlease wait for a while before re-running this '
                              'code. Also, reduce number of max_workers in concurrency and consider switching to a '
                              'new Network ID.') if not printed else None
            printed = True  # makes sure the print statement happens only once
            # stop future threads to avoid progress bar on screen post print
            raise ConnectionRefusedError  # handle exception so that spreadsheet is created with existing stock_map dict
        else:
            if err.code == 404:
                count_404 += 1  # increases count_404 for future handling
            file_logger.error(f'Failed to analyze {stock}. Faced error code {err.code} while requesting {err.url}. '
                              f'Reason: {err.reason}.')
    except (ConnectionError, ProtocolError, ConnectionResetError, ChunkedEncodingError) as conn_err:
        file_logger.error(f'Failed to analyze {stock}.\n{conn_err}')


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
        worksheet.write(n, 0, ticker)
        worksheet.write(n, 1, mapping_dict[ticker][0])
        worksheet.write(n, 2, mapping_dict[ticker][1])
        worksheet.write(n, 3, mapping_dict[ticker][2])
        worksheet.write(n, 4, mapping_dict[ticker][3])
        worksheet.write(n, 5, mapping_dict[ticker][4])
        worksheet.write(n, 6, mapping_dict[ticker][5])
        worksheet.write(n, 7, mapping_dict[ticker][6])
        worksheet.write(n, 8, mapping_dict[ticker][7])
        worksheet.write(n, 9, mapping_dict[ticker][8])
        worksheet.write(n, 10, mapping_dict[ticker][9])
        worksheet.write(n, 11, mapping_dict[ticker][10])
        worksheet.write(n, 12, mapping_dict[ticker][11])
        worksheet.write(n, 13, mapping_dict[ticker][12])
        worksheet.write(n, 14, mapping_dict[ticker][13])
        worksheet.write(n, 15, mapping_dict[ticker][14])
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
        option, index = pick(headers[1:], title, indicator='=>', default_index=0)
        return index
    except (error, KeyboardInterrupt):
        if not (run_env := Process(getpid()).parent().name()).endswith('sh'):
            root_logger.error(f"You're using {run_env} to run the script.")
            root_logger.error("Either use a terminal or enable 'Emulate terminal in output console' under "
                              f"Edit Configurations.. -> Execution in your {run_env}.")
            root_logger.error("Using default index to sort the spreadsheet.")
        else:
            root_logger.error(error)
        print('I AM HERE')
        exit()


def thread_executor() -> None:
    """Executes ``ThreadPool`` on all stock tickers with a max workers limit of 10.

    Warnings:
        - ``max_workers`` for ThreadPool is set to 10.
        - Increasing the number of workers will decrease the run time but may elevate the chances of a ``503`` response.
        - Shuts down the ThreadPool during either of the following:

            - KeyboardInterrupt (manual interrupt)
            - ConnectionRefusedError (raised by analyzer in case of a 503) exceptions.
    """
    console_logger.info(f'Instantiating multi threading to analyze {overall} NASDAQ stocks')
    try:
        with ThreadPoolExecutor(max_workers=10) as executor:  # multi threaded to 10 workers for throttled processing
            list(tqdm(executor.map(analyzer, stocks), total=overall, desc='Analyzing Stocks', unit='stock', leave=True))
    except ConnectionRefusedError:
        root_logger.error('Connection has been refused.')
    except KeyboardInterrupt:
        root_logger.error('Manual interrupt was received.')
    ThreadPoolExecutor().shutdown(wait=False, cancel_futures=True)


def finalizer() -> None:
    """Logs all the closure information and opens the spreadsheet."""
    console_logger.info(f'Total Stocks instantiated: {overall}')
    console_logger.info(f'Total Stocks analyzed: {analyzed}')
    console_logger.info(f'Total Stocks failed to analyze: {overall - analyzed}')

    time_taken = time_converter(round(perf_counter()))
    console_logger.info(f'Total execution time: {time_taken}')
    if analyzed:
        console_logger.info(f'Spreadsheet stored as {filename}')
        system(f'open {filename}')  # opens spreadsheet post execution


if __name__ == '__main__':
    # import in _main_ so that data and logs dir are created in advance
    from lib.helper_functions import logging_wrapper, nasdaq

    file_logger, console_logger, root_logger = logging_wrapper()

    headers = columns()  # stores all the titles into a variable
    filename = datetime.now().strftime('data/stocks_%H:%M_%d-%m-%Y.xlsx')  # creates filename with date and time
    workbook = Workbook(filename, {'strings_to_numbers': True})  # allows possible strings as numbers
    worksheet = workbook.add_worksheet('Results')  # sheet name in the workbook
    worksheet_initializer()  # initializes worksheet
    stocks = nasdaq()  # gets all the NASDAQ stock ticket values starting A to Z
    overall = len(stocks)  # stores the number of stock tickers in a variable

    # other variables initialization
    stock_map = {}  # initiates stock_map as an empty dict
    count_404 = 0  # 404 responses recorded to see if it is repeated
    printed = False  # initiates printed as False

    thread_executor()  # kicks off multi-threading

    sort_val = get_sort_key()
    if sort_val:
        stock_map = sort_by_value(data=stock_map, sort=sort_val)
    analyzed = writer(mapping_dict=stock_map)  # gets the number of stocks analyzed after writing to workbook
    finalizer()
