import string
import requests
from bs4 import BeautifulSoup as bs
import logging
from datetime import datetime

log_filename = datetime.now().strftime('logs/stock_logs_%H:%M_%d-%m-%Y.log')
logging.basicConfig(filename=log_filename, level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger('analyzer.py')


def nasdaq():
    char = string.ascii_uppercase
    stock_list = []
    logger.info('Fetching tickers for all NASDAQ stocks')
    print('Fetching tickers for all NASDAQ stocks')
    for x in char:
        url = f'http://www.eoddata.com/stocklist/NASDAQ/{x}.htm'
        r = requests.get(url)
        scrapped = bs(r.text, "html.parser")
        d1 = scrapped.find_all('tr', {'class': 'ro'})
        d2 = scrapped.find_all('tr', {'class': 're'})
        for link in d1:
            stock_list.append(f"{(link.get('onclick').split('/')[-1]).split('.')[0]}")
        for link in d2:
            stock_list.append(f"{(link.get('onclick').split('/')[-1]).split('.')[0]}")

    return stock_list


if __name__ == '__main__':
    from pprint import pprint
    pprint(nasdaq())
