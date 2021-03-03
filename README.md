# Stock Analyzer
Analyze all NASDAQ stocks on [Yahoo Finance](https://finance.yahoo.com) using [YFinance](https://pypi.org/project/yfinance/)

### Libraries Used:
- ThreadPoolExecutor - Uses a pool of threads to execute calls asynchronously
- YFinance - Yahoo API to request stock information for each ticker value - [thor_api](thor_api.py)
- Pandas - Retrieve tables while using web calls - [thor](thor.py)
- BeautifulSoup - Retrieves information in non-tables
- tqdm - Progress bar
- xlsxwriter - Writes data into a spreadsheet

## License & copyright

&copy; Vignesh Sivanandha Rao, Stock Analyzer

Licensed under the [MIT License](LICENSE)
