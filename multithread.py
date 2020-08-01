from concurrent.futures import ThreadPoolExecutor
import requests
import pandas as pd
from bs4 import BeautifulSoup as bs
import time
from tqdm import tqdm


start_time = time.time()


def worker(stocks):
    a = ''
    for stock in stocks:
        try:
            summary = f'https://finance.yahoo.com/quote/{stock}/'
            stats = f'https://finance.yahoo.com/quote/{stock}/key-statistics/'
            analysis = f'https://finance.yahoo.com/quote/{stock}/analysis/'
            r = requests.get(f'https://finance.yahoo.com/quote/{stock}/')
            summary_result = pd.read_html(summary, flavor='bs4')
            market_capital = summary_result[-1].iat[0, 1]
            pe_ratio = summary_result[-1].iat[2, 1]
            forward_dividend_yield = summary_result[-1].iat[5, 1]
            a += market_capital
            a += pe_ratio
            a += forward_dividend_yield

            scrapped = bs(r.text, "html.parser")
            raw_data = scrapped.find_all('div', {'class': 'My(6px) Pos(r) smartphone_Mt(6px)'})[0]
            price = float(raw_data.find('span').text)
            a += price

            stats_result = pd.read_html(stats, flavor='bs4')
            high = stats_result[0].iat[3, 1]
            low = stats_result[0].iat[4, 1]
            profit_margin = stats_result[5].iat[0, 1]
            price_book_ratio = stats_result[0].iat[6, 1]
            return_on_equity = stats_result[6].iat[1, 1]
            a += high, low, profit_margin, price_book_ratio, return_on_equity

            analysis_result = pd.read_html(analysis, flavor='bs4')
            analysis_next_year = analysis_result[-1].iat[3, 1]
            analysis_next_5_years = analysis_result[-1].iat[4, 1]
            analysis_past_5_years = analysis_result[-1].iat[5, 1]
            a += analysis_next_5_years, analysis_next_year, analysis_past_5_years
        except:
            pass
    return a


def execute():
    with ThreadPoolExecutor(max_workers=50) as executor:
        # executor.map(worker, stocks)
        r = list(tqdm(executor.map(worker, stocks), total=len(stocks)))
    return r


if __name__ == '__main__':
    stocks = ['AACG', 'AAL', 'AAPL', 'ABCB', 'ABIO', 'ABTX', 'ACER', 'ACHC', 'ACIW', 'ACMR', 'ACOR', 'ACRX', 'ADES',
              'ADIL', 'ADMA', 'ADMS', 'ADPT', 'ADRO', 'ADTN', 'ADUS', 'AEGN', 'AEIS', 'AERI', 'AEY']
    print(execute())

print(f'\nExecution time = {round(float(time.time() - start_time), 2)} seconds')
