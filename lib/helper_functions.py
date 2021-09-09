import logging
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from importlib import reload
from string import ascii_uppercase

from bs4 import BeautifulSoup
from requests import get

stock_list = []


def logging_wrapper() -> tuple:
    """Wraps logging module to create multiple handlers for different purposes.

    See Also:
        - fileLogger: Writes the log information only to the log file.
        - consoleLogger: Writes the log information only in stdout.
        - rootLogger: Logs the entry in both stdout and log file.

    Returns:
        tuple:
        A tuple of classes logging.Logger for file, console and root logging.
    """
    log_file = datetime.now().strftime('logs/stock_logs_%H:%M_%d-%m-%Y.log')
    reload(logging)  # since the gmail-connector module uses logging, it is better to reload logging module before start
    log_formatter = logging.Formatter(
        fmt="%(asctime)s - [%(levelname)s] - %(name)s - %(funcName)s - Line: %(lineno)d - %(message)s",
        datefmt='%b-%d-%Y %H:%M:%S'
    )

    file_logger = logging.getLogger('FILE')
    console_logger = logging.getLogger('CONSOLE')
    root_logger = logging.getLogger("thor")

    file_handler = logging.FileHandler(filename=log_file)
    file_handler.setFormatter(fmt=log_formatter)
    file_logger.setLevel(level=logging.DEBUG)
    file_logger.addHandler(hdlr=file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(fmt=log_formatter)
    console_logger.setLevel(level=logging.DEBUG)
    console_logger.addHandler(hdlr=console_handler)

    root_logger.addHandler(hdlr=file_handler)
    root_logger.addHandler(hdlr=console_handler)
    root_logger.setLevel(level=logging.DEBUG)
    return file_logger, console_logger, root_logger


def ticker_gatherer(character: str) -> None:
    """Gathers the stock ticker in NASDAQ. Runs on ``multi-threading`` which drops run time by ~7 times.

    Args:
        character: ASCII character (alphabet) with which the stock ticker name starts.
    """
    url = f'https://www.eoddata.com/stocklist/NASDAQ/{character}.htm'
    response = get(url)
    scrapped = BeautifulSoup(response.text, "html.parser")
    d1 = scrapped.find_all('tr', {'class': 'ro'})
    d2 = scrapped.find_all('tr', {'class': 're'})
    for link in d1:
        stock_list.append(f"{(link.get('onclick').split('/')[-1]).split('.')[0]}")
    for link in d2:
        stock_list.append(f"{(link.get('onclick').split('/')[-1]).split('.')[0]}")


def nasdaq() -> list:
    """Spins up 26 threads, (one for each alphabet) and calls ``ticker_gatherer`` to get ticker values of that alphabet.

    Returns:
        list:
        List of stock tickers.
    """
    file_logger, console_logger, root_logger = logging_wrapper()
    alphabets = ascii_uppercase
    console_logger.info('Fetching tickers for all NASDAQ stocks')
    with ThreadPoolExecutor(max_workers=len(alphabets)) as executor:
        executor.map(ticker_gatherer, alphabets)
    return stock_list


if __name__ == '__main__':
    from pprint import pprint
    pprint(nasdaq())
