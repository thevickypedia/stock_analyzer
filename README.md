# Stock Analyzer
Analyze all NASDAQ stocks on [Yahoo Finance](https://finance.yahoo.com) using [YFinance](https://pypi.org/project/yfinance/) API

### Libraries Used
- `ThreadPoolExecutor` - Uses a pool of threads to execute calls asynchronously
- `YFinance` - Yahoo API to request stock information for each ticker value
- `Tqdm` - Progress bar
- `Xlsxwriter` - Writes data into a spreadsheet
- `numerize` - Converts float value to understandable currency value (Example: `568153344` to `568.15M`)
- `pick` - Lets user to, choose a value to sort the source dictionary before writing to the spreadsheet

[Legacy:](https://github.com/thevickypedia/stock_analyzer/blob/master/thor_legacy.py)
- `Pandas` - Retrieve tables while using web calls
- `BeautifulSoup` - Retrieves information in non-tables

### Options
- [Web calls - legacy](https://github.com/thevickypedia/stock_analyzer/blob/master/thor_legacy.py) - Uses web calls to https://finance.yahoo.com
- [API](https://github.com/thevickypedia/stock_analyzer/blob/master/thor_api.py) - Uses Yahoo Finance API

### Instructions
1. `git clone https://github.com/thevickypedia/stock_analyzer.git`
2. `python3 -m venv venv`
3. `source venv/bin/activate`
4. `pip3 install -r requirements.txt`
5. `python3 thor_api.py`

### Linting
`PreCommit` will ensure linting, and the doc creation are run on every commit.

Requirement:
<br>
`pip install --no-cache --upgrade sphinx pre-commit recommonmark`

Usage:
<br>
`pre-commit run --all-files`

### Links
[Repository](https://github.com/thevickypedia/stock_analyzer)

[Runbook](https://thevickypedia.github.io/stock_analyzer/)

## License & copyright

&copy; Vignesh Sivanandha Rao, [Stock Analyzer](https://github.com/thevickypedia/stock_analyzer)

Licensed under the [MIT License](https://github.com/thevickypedia/stock_analyzer/blob/master/LICENSE)
