# -*- coding: utf-8 -*-

"""See LICENSE below for authorship, copyright and license information"""

import io
import os
import csv
import time
# import shutil
import datetime
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
from urllib.request import urlopen  # Request,
pd.options.mode.chained_assignment = None  # default='warn'
from pandas.plotting import register_matplotlib_converters
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
register_matplotlib_converters()

MAKE_PLOT = True
SHOW_PLOTS = False
WRITE_XL = True
UPDATE_FILES = False
DELETE_OBS_FILES = False
START_DATE = datetime.date(year=1993, month=1, day=1)
END_DATE = datetime.date.today() + datetime.timedelta(days=1)
GITHUB = 'https://github.com/Vettejeep/Financial_Stuff'

ABBR_FOR_AVG = 'months'
TERM_SHORT = 3
TERM_LONG = 42
ONE_MN_AVG = int(20)
SHORT_AVG = int(TERM_SHORT*4.33*5)
LONG_AVG = int(TERM_LONG*4.33*5)
RSI_SHORT = 14
RSI_LONG = 60
PLT_WIDTH = 7
PLT_HT = 4
PLT_HT_MAJOR = 3
PLT_HT_MINOR = 1
CELL_WIDTH = PLT_WIDTH * 13.625
CELL_HT = PLT_HT * 75

DATA_FOLDER = r'D:\WebScrape\raw'
TIMEOUT_SEC = 10
SLEEP_SEC = 1  # do not know yahoo policy, but a timeout may help to keep them from blocking us
COLS = ['Date', 'Open', 'High', 'Low', 'Close', 'Adj_Close', 'Volume']

__VERSION__ = '1.2.0'

INTRODUCTION = f"""Introduction & What is this?
Historical data charted for a list of asset ticker symbols from Yahoo finance, generally ETFs, futures and stocks.
Plots of the price ratio over time between two securities as sets of moving averages to smooth out noise.
The price ratio, plotted over time, tells you if the study security has performed better or worse than the base asset.
I believe that the charts here are valuable information for assisting in investment decisions.

ETFs can be bought or sold like stocks and can represent market indexes that otherwise cannot be directly invested in.
For ticker info just do an internet search for something like: <ticker> stock quote.
The different sheets of this Excel file contain plots, mostly organized by the base asset ticker and time frame.
Updated Excel output files are available at: {GITHUB}"""

DISCLAIMER = """
Disclaimer:
I am not a financial or investment professional, do not construe anything that I say or write as investment advice. 
This is not financial or investment advice, if you listen to me, quite possibly you will lose money.
I may have positions in some of the securities here.
I am not responsible for losses incurred by any user in any asset purchase or sale decision, long or short.
Validate all information herein, no guarantee of accuracy, use at your own risk.

Nothing in this file or any of my comments anywhere is fully inclusive as to all possible investment risks, so do your
own research and act accordingly. I cannot predict future events, which may cause losses in any investment.
The existence of an asset in these studies does not represent an endorsement of that security, it is simply something
deemed interesting to study. Some study may be done here on bad investments, just for curiosity.
There is no implication that any research has been performed on these securities beyond the plots given.
ETFs or commodities that invest in futures contracts or leverage can be especially toxic for many reasons, however, 
they may be included here for study purposes."""

CHARTING = """
Charting:
The charts here can be a portion of the useful information for assisting in investment decisions.
You also need information about world events, politics, macro economics and specific economic factors related to the 
asset classes and/or ETFs, futures or stocks that you might purchase or go short on, none of that is presented here.
It is also valuable to have and understand technically oriented charts for assets you are considering investing
in, these charts are not included in this study.
Moving averages may not be fully valid for the earliest dates in the plots, especially for the longest term average.
Plots may go back to 1993, but it depends on when the asset was created or Yahoo Finance data limitations"""

INTERPRETATION = """
Interpretation:
If the moving averages are going up, then the study security is out-performing the base security.
If the moving averages are going down, then the study security is under-performing the base security.
Some people consider cross-overs of the shorter term to longer term moving averages to be buy & sell signals,  
other trend interpretations exist.

The plots are based on moving averages, thus their information may be behind current market behavior.
Past performance is no guarantee of future performance, trends do change over time.
As an investor, one hopes to find trends that will remain in place for years, many, but not all, seem to do so.
These charts do not tell if asset prices are, or will be, increasing or decreasing, they may both be falling together, 
for example."""

BASES = """
Important Base Tickers:
SPY: ETF that is a proxy for the S&P 500
GC=F: Gold Front Month Futures"""

CONTACT = """
You are welcome to contact me:
Email: Vettejeep365@gmail.com
Twitter: @FlyingFish365 - please follow and like on Twitter and I will work to expand this work and keep this 
effort reasonably current. If you follow on Twitter, you will get notification of updates.

These charts don't change quickly, so every 2-4 weeks seems about right. Or, updates
will be posted when significant code changes are made.

Community discussion of these charts and investment ideas on Twitter is encouraged.
RSI has been a challenge for the longer term plots - specifically what time period to use.

Adding more asset classes (additional ticker symbols) is very easy to do. They need to be available on Yahoo Finance.
For added tickers it is best if they have existed for 5+ years so the longer term averages are somewhat reasonable.

Thanks for any kudos, or suggestions for the improvement of this study,
Kevin"""

CHANGES = """
Recent Changes:
1. Added RSI
2. Added short term plots (1 year)
3. Added some tickers
As of 2020-08-16"""

PYTHON_NOTES = f"""
Python: 
This Excel file was generated by a Python script: {os.path.basename(__file__)}, code version {__VERSION__}.
The source code is posted on Github as open source at {GITHUB}.

It is probably necessary to understand at least some Python in order to work with the code, plus the web scrape and
numeric library stack for Python.  This is still a rough piece of code, not all paths tested, so expect bugs.
No exception handling yet - for example when a file has insufficient data.
Designed for a reasonably up-to-date Anaconda Python 3 distribution.

It takes a very long time (hours) the first time it is run in order to obtain all needed historical data.
After historical data has been acquired, updates to the ticker files are reasonably fast if done at least monthly.

Usage (it helps to know Python to figure it out):
Create needed directories manually (to be fixed).
Adjust the settings on the ALL_CAPS variables to vary script operation.
Set your lists of tickers for acquisition and study, or use my settings.
Run the script using Python...
- not written for non-programmers to use, not user-friendly.
- it is pretty simple code, but, you will need to read the Python..."""

LICENSE = f"""
Licensing, the code is intended as open source code within the meaning of the license.

Code licensed under the GNU General Public License v3.0
Permissions of this strong copyleft license are conditioned on making available complete source code of licensed 
works and modifications, which include larger works using a licensed work, under the same license. Copyright and 
license notices must be preserved. Contributors provide an express grant of patent rights.
See: https://www.gnu.org/licenses/gpl-3.0.en.html.

Excel output, graphs, code and related project files: © 2020 Kevin Maher. 
Permission to use or copy any and all files or portions thereof granted as long as my Twitter handle 
is displayed or a citation in any normal academic format is provided which includes a link to 
{GITHUB}, and the above license is adhered to."""

CITATIONS = """
Sources used in writing this code.
RSI Calculation: # http://www.andrewshamlet.net/2017/06/10/python-tutorial-rsi/ (2020-08-16)
Fix RSI for newer Pandas: # https://stackoverflow.com/questions/57055894/has-pandas-stats-been-deprecated (2020-08-16)"""

OTHER_TICKERS = """
Other interesting tickers. Some do not work in ratio charts because of limited time data.
Not used in code or output, just a reference list.
No implication that any meaningful research has been performed on these tickers.
IVOL: bond fund playing fixed vs inflation protected bonds, maybe too new to plot, limited history, recently volatile.
TECB: high tech and biotech etf much like XT, limited history, low trading volume.
IAU: iShares version of GLD, divided into smaller shares, lower share price eases access for smaller investors.
OUNZ: VanEck gold etf like GLD that allows conversion to physical metal in coins or bars.
RING: iShares gold miner ETF like GDX
PPLT: Platinum physical metal ETF
FDN: Dow Jones Internet Index Fund
IGSB: iShares short-term ig corporate
SHY: iShares short-term treasury
IEI: iShares 3-7 yr treasury
SLQD: iShares 0-5 year ig corporate
FLOT: iShares floating rate bond
GDLC: bitcoin, etherium, etc
FMIL: Fidelity active ETF, all cap, opportunity fund, limited history, low trading volume.
FBCG: Fidelity active ETF, blue chip growth, limited history, low trading volume.
FBCV: Fidelity active ETF, blue chip value, limited history, low trading volume.
NACP: social good - diversity - investment bs or a trend?
MTUM: iShares momentum stocks
ISHG: int'l govt bond x-us, 1-3 year, has had negative yield, but seems to climb when the dollar falls
IGOV: int'l govt bond x-us, has had negative yield, but seems to climb when the dollar falls
"""

# https://www.marketwatch.com/story/heres-why-90-of-rich-people-squander-their-fortunes-2017-04-23

# TODO: try/except blocks so it can recover and do what it can, esp scraping tickers
# TODO: validate base & study tickers so it does not crash if the file for the ticker does not exist
# TODO: multi-panel plot with RSI, asset price
# TODO: warn/retry if a month comes up missing in web scrape, rare issue, data gap plots a hole now
# TODO: verify time zone issues generating time stamps
# TODO: go to 3 month internet data collection to speed up building the data files - ok with yahoo limitations
# TODO: automate sub directories
# TODO: crashes if ticker has inadequate data for RSI, need exception handler
# TODO: dynamically adjust plot legend and text based on data so data is not over-written
# TODO: font sizes on charts not working as expected, research into matplotlib axes needed

BASE_TKRS = ['SPY', 'GC=F']
PAIR_TKRS = {
    'SPY': 'RSP',
    'JKE': 'JKF',
    'SI=F': 'GC=F'
}

# tickers, along with a very short description that will fit in the chart title
# these are the ticker symbols that will be web scraped and updated in files
TKR_DICT = {
    'SPY': 'S&P 500',
    'GC=F': 'Gold Futures',
    'QQQ': 'Nasdaq 100',
    'RSP': 'Eq Wt S&P 500',
    'EFA': 'MSCI EAFE Ex-US',
    'IWM': 'Russel 2000',
    'IYT': 'Transportation',
    'GLD': 'Gold Metal',
    'GDX': 'Gold Miners',
    'GDXJ': 'Jr Gold Miners',
    'SI=F': 'Silver Futures',
    'SLV': 'Silver Metal',
    'SIL': 'Silver Miners',
    'SILJ': 'Jr Silver Miners',
    'JKF': 'Large Value',
    'JKE': 'Large Growth',

    'PPLT': 'Platinum Metal',  # comment out these & below for a short test run
    'CPER': 'Copper Metal',
    'EEM': 'Emerging Mkt',
    'IDEV': 'Dev Mkt Ex-US',
    'IJH': 'US Mid-Cap',
    'IJR': 'US Small Cap',
    'IYR': 'Commercial RE',
    'CCJ': 'Uranium Stock',
    'URPTF': 'Uranium Metal',  # or URPTF in the US
    'MOO': 'Agri-Business',
    'DBA': 'Ag commodities',
    'PAVE': 'Infrastructure',
    'ICLN': 'Clean Energy',
    'PBW': 'Clean Energy',
    'BOTZ': 'AI',
    'ROBO': 'AI',
    'IDRV': 'Self-Driving',
    'QQEW': 'Nasdaq 100 EW',
    'XT': 'Exponential Tech & Bio',
    'IBB': 'Bio-Tech',
    'IYW': 'US Tech',
    'ICVT': 'Convertible Bonds',
    'XLC': 'S&P Communication',
    'XLY': 'S&P Consumer Disc',
    'XLP': 'S&P Consumer Staples',
    'XLE': 'S&P Energy',
    'XLF': 'S&P Financials',
    'XLV': 'S&P Health Care',
    'XLI': 'S&P Industrial',
    'XLB': 'S&P Materials',
    'XLRE': 'S&P Real Estate',
    'XLK': 'S&P Technology',
    'XLU': 'S&P Utilities',
    'GBTC': 'Bitcoin ETF',
    # 'BTC-USD': 'Bitcoin',
    # 'ETH-USD': 'Etherium',  # odd results with crypto at the moment, date range issues in web scrape
    'BNO': 'Brent Oil ETF'
}

TICKERS = ['SPY', 'XLC', 'XLY', 'XLP', 'XLE', 'XLF', 'XLV', 'XLI', 'XLB', 'XLRE', 'XLK', 'XLU',  # S&P 500 + sectors
           'QQQ', 'IWM', 'IJH', 'IJR',   # NASDAQ, Russel 2000, US mid cap, US small cap
           'EFA', 'EEM', 'IDEV',   # Global outside US, emerging markets, core developed
           'LQD', 'GOVT', 'TIP', 'ICVT', 'FBND',  # corp bonds, US govt bonds, TIPS bonds, convertible bonds, mxd bonds
           'GLD', 'GDX', 'GDXJ',  # gold, miners, junior miners
           'SLV', 'SIL', 'SILJ',  # silver, miners, junior miners
           #'BNO', 'USO',  # brent oil, WTI oil; very messed up investments
           'IYR', 'IYT', 'IDU',  # commercial real estate, transportation sector, utilities
           'XT', 'IBB', 'IYW',    # Exponential Tech+Bio, Biotech, iShares tech; ETFs of interest
           'JKF', 'JKE',  # iShares Large Value, iShares Large Growth; compare value vs growth styles of investing
           'GC=F', 'SI=F',  # gold and silver futures, for info because prices go back further in time
           'FTGC', 'RING', 'IAU',  # active commodity fund, gold miners, gold bullion
           # 'NEM', 'AEM', 'GOLD',  # major gold mines with long history
           # 'ESGU', 'ICLN', 'DSI',  # sustainable investing, apparently a trend among younger investors
           # 'PPLT', 'NGLOY', 'CPER'  # platinum, ngloy is largest platinum mine, copper
           # 'EWJ', 'EWG', 'EWU', 'FXI', 'EWC', 'EPI', 'EWA' # Japan, Germany, UK, China, Canada, India, Aussie https://www.cnbc.com/country-etfs/
           'GBTC',  # bitcoin
           'RSP', 'PFF', 'FCPI', 'QQEW',  # S&P 500 equal weight, preferred and income stocks, fidelity inflation stocks, Nasdaq equal weight
           'ICLN', 'PBW',  # clean energy etfs - biden spend heavy?
           'PAVE',  # infrastructure - biden spend heavy?
           'BOTZ', 'ROBO', 'IDRV',  # artificial intelligence, self-driving etf
           'CCJ',  # uranium (stock)
           'MOO',  # agribusiness
]

# TICKERS = ['MTUM']

# platinum - some have poor financials
platinum = ['NGLOY', 'IMPUY', 'PLG', 'SBSW', 'NMPNF', 'IVPAF', 'ELRFF', 'NKORF', 'NMTLF']
uranium = ['CCJ', 'NXE', 'UEC', 'URG', 'ISENF', 'AZZUF', 'DNN', 'FCUUF', 'WSTRF']


def replace_comma(x):
    x = str(x)
    return x.replace(",", "")


def clean_files():
    files = [x for x in os.listdir(DATA_FOLDER)]

    for f in files:
        pd.read_csv(os.path.join(DATA_FOLDER, f)).iloc[0: -1].to_csv(os.path.join(DATA_FOLDER, f), index=False)

# TODO: check if start date in file, currently does not work if an earlier date specified
def update_tickers(remove=False):
    files = [x for x in os.listdir(DATA_FOLDER)]

    # remove if not in list
    if remove:
        for f in files:
            if f.split('.')[0] not in TICKERS:
                os.remove(os.path.join(DATA_FOLDER, f))

    for t in TKR_DICT.keys():
        print(f'starting work on {t}')
        fp = os.path.join(DATA_FOLDER, t + '.csv')
        start_date = START_DATE
        end_date = (start_date.replace(day=28) + datetime.timedelta(days=5)).replace(day=1)
        path_exists = os.path.exists(fp)
        df_tkr = None

        if path_exists:
            df_tkr = pd.read_csv(fp)
            df_tkr.Date = df_tkr.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())
            start_date = df_tkr.Date.iloc[-1] + datetime.timedelta(days=1)

            if start_date < END_DATE:
                end_date = (start_date.replace(day=28) + datetime.timedelta(days=5)).replace(day=1)

        good = True
        # print(start_date, type(start_date), end_date, type(end_date))
        # print()

        # maybe a better way, but yahoo returns a limited number of quotes, so for now this is dome one month at a time
        while good:
            start_ts = int(datetime.datetime.combine(start_date, datetime.datetime.min.time()).timestamp())
            end_ts = int(datetime.datetime.combine(end_date, datetime.datetime.min.time()).timestamp() - 10)
            url = url = f'https://finance.yahoo.com/quote/{t}/history?period1={start_ts}&period2={end_ts}&interval=1d&filter=history&frequency=1d'
            print(url)
            print(start_ts, end_ts)

            ct = 0
            done = False
            table = None
            while not done and ct < 6:
                try:
                    page = urlopen(url, timeout=TIMEOUT_SEC)
                    soup = BeautifulSoup(page, features="lxml")
                    table = soup.find('table')
                    if table is not None:
                        done = True
                except:
                    print(f'exception sleep {SLEEP_SEC} seconds, Start Date: {start_date}, End Date: {end_date}, Count: {ct}')
                    time.sleep(SLEEP_SEC)
                    table = None
                finally:
                    ct += 1

            if table is None:
                good = False
                break

            df = pd.DataFrame(columns=COLS)
            count = 0
            # parse the html table
            for row in table.find_all('tr'):
                # print('row: ', row)
                count += 1
                data = {}
                for column_idx, column in enumerate(row.find_all('td')):
                    # print(column)
                    # print(column.get_text().strip().lstrip('$'))
                    data[COLS[column_idx]] = column.get_text().strip().lstrip('$')
                df = df.append(data, ignore_index=True)

            df = df.dropna(axis=0, how='any', inplace=False)
            # print(df.shape, count)
            df['Date'] = df['Date'].apply(lambda x: datetime.datetime.strptime(x, '%b %d, %Y').date())
            df.sort_values(by='Date', inplace=True)
            # print(df.tail())

            # print(df.shape)
            if df.shape[0] > 0:
                if df_tkr is None:
                    df_tkr = df
                else:
                    df_tkr = df_tkr.append(df)

            start_date = (start_date.replace(day=28) + datetime.timedelta(days=5)).replace(day=1)
            end_date = (end_date.replace(day=28) + datetime.timedelta(days=5)).replace(day=1)

            print(f'sleep {SLEEP_SEC} seconds, Start Date: {start_date}, End Date: {end_date}, Good: {good}')
            time.sleep(SLEEP_SEC)

            if start_date > END_DATE or start_date > (datetime.date.today() - datetime.timedelta(days=1)):
                good = False

        # print(start_date, end_date)
        if df_tkr is not None:
            # even when past the last date available, Yahoo seems to return one value sometimes, so drop duplicates
            df_tkr.replace(to_replace='-', value=0, inplace=True)
            df_tkr.Open = df_tkr.Open.apply(lambda x: replace_comma(x)).astype(np.float32).fillna(0.0)
            df_tkr.High = df_tkr.High.apply(lambda x: replace_comma(x)).astype(np.float32).fillna(0.0)
            df_tkr.Low = df_tkr.Low.apply(lambda x: replace_comma(x)).astype(np.float32).fillna(0.0)
            df_tkr.Close = df_tkr.Close.apply(lambda x: replace_comma(x)).astype(np.float32).fillna(0.0)
            df_tkr['Adj_Close'] = df_tkr['Adj_Close'].apply(lambda x: replace_comma(x)).astype(np.float32).fillna(0.0)
            df_tkr.Volume = df_tkr.Volume.apply(lambda x: replace_comma(x)).astype(np.uint64).fillna(0)
            df_tkr = df_tkr.loc[df_tkr.Close != 0.0]
            df_tkr.drop_duplicates(subset=['Date'], inplace=True)
            df_tkr.reset_index(drop=True, inplace=True)
            df_tkr['Date'] = df_tkr['Date'].apply(lambda x: datetime.datetime.strftime(x, '%Y-%m-%d'))
            df_tkr.to_csv(fp, quoting=csv.QUOTE_NONE, index=False)
            print(f'data frame saved for {t}')
        else:
            print(f'{t} unable to be written')


def calc_ma(df, field):
    # df[f'ma_5day_{field}'] = df[field].rolling(window=5, min_periods=1).mean()
    df[f'ma_one_mo_{field}'] = df[field].rolling(window=ONE_MN_AVG, min_periods=1).mean()
    df[f'ma_short_{field}'] = df[field].rolling(window=SHORT_AVG, min_periods=1).mean()
    df[f'ma_long_{field}'] = df[field].rolling(window=LONG_AVG, min_periods=1).mean()
    return df


def cci(df, high, low, close, suffix):
    ave_p = (df[high] + df[low] + df[close]).rolling(window=20, min_periods=1).mean() / 3.0
    ave_p_sma = ave_p.rolling(window=20, min_periods=1).mean()
    md = (ave_p - ave_p_sma).rolling(window=20, min_periods=1).mean()
    df[f'cci_{suffix}'] = (ave_p - ave_p_sma) / (0.015 * md)
    return df

def chk_gap(t, df):
    if df.Low.iloc[-1] > df.High.iloc[-2]:
        gap = round(df.Low.iloc[-1]-df.High.iloc[-2], 2)
        lo = round(df.Low.iloc[-1], 2)
        hi = round(df.High.iloc[-2])
        print(f'Gap for {t}: {gap}, Low: {lo}, Hi: {hi}')

def ref_portfolio(tkers, allocs):
    date_now = datetime.date.today()
    date_start = (datetime.datetime.now() - datetime.timedelta(days=365)).date().replace(day=1)
    returns = []
    ytd_returns = []

    while date_start <= date_now:
        date_end = (date_start + datetime.timedelta(days=32)).replace(day=1)
        return_prop = 0.0
        ytd_return_prop = 0.0
        for tkr, alloc in zip(tkers, allocs):
            base_file = [x for x in os.listdir(DATA_FOLDER) if tkr in x][0]
            df = pd.read_csv(os.path.join(DATA_FOLDER, base_file))
            df.Date = df.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())
            df = df.loc[df.Date <= date_end, ].loc[df.Date >= date_start, ].reset_index(drop=True, inplace=False)
            base_pr = df.Adj_Close.iloc[0]
            end_pr = df.Adj_Close.iloc[-1]
            # =0.12*((320.98/310.52)-1)
            return_prop += alloc * (end_pr / base_pr)

            if date_start.year == date_now.year:
                ytd_return_prop += alloc * (end_pr / base_pr)

        returns.append(return_prop)
        if ytd_return_prop != 0.0:
            ytd_returns.append(ytd_return_prop)

        print(round(returns[-1], 3), date_start, date_end)
        date_start = date_end
    print(ytd_returns)
    total_ret = np.round(np.prod(returns[: -1]), 4)
    ytd_ret = np.round(np.prod(ytd_returns[: -1]), 4)
    cur_ytd_ret = np.round(np.prod(ytd_returns), 4)
    print(total_ret, ytd_ret, cur_ytd_ret)
    print()


# https://tcoil.info/compute-rsi-for-stocks-with-python-relative-strength-index/
def compute_rsi(data, time_window=RSI_SHORT):
    diff = data.diff(1).dropna()  # diff in one field(one day)

    # this preservers dimensions off diff values
    up_chg = 0 * diff
    down_chg = 0 * diff

    # up change is equal to the positive difference, otherwise equal to zero
    up_chg[diff > 0] = diff[diff > 0]

    # down change is equal to negative deifference, otherwise equal to zero
    down_chg[diff < 0] = diff[diff < 0]

    # check pandas documentation for ewm
    # https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.ewm.html
    # values are related to exponential decay
    # we set com=time_window-1 so we get decay alpha=1/time_window
    up_chg_avg = up_chg.ewm(com=time_window - 1, min_periods=time_window).mean()
    down_chg_avg = down_chg.ewm(com=time_window - 1, min_periods=time_window).mean()

    rs = abs(up_chg_avg / down_chg_avg)
    rsi = 100 - 100 / (1 + rs)

    # original function 1 data point short, complicates plotting, so corrected here - add: also reset index?
    rsi = pd.concat([rsi.iloc[0: 1], rsi])  # 0: 1 - concat needs a pd series, here of one item
    return rsi

# http://www.andrewshamlet.net/2017/06/10/python-tutorial-rsi/
# https://stackoverflow.com/questions/57055894/has-pandas-stats-been-deprecated
def RSI(series, period=RSI_SHORT):
    adj_length = series.shape[0]
    delta = series.diff().dropna()
    u = delta * 0
    d = u.copy()
    u[delta > 0] = delta[delta > 0]
    d[delta < 0] = -delta[delta < 0]
    u[u.index[period-1]] = np.mean( u[:period] ) #first value is sum of avg gains
    u = u.drop(u.index[:(period-1)])
    d[d.index[period-1]] = np.mean( d[:period] ) #first value is sum of avg losses
    d = d.drop(d.index[:(period-1)])
    # rs = pd.stats.moments.ewma(u, com=period-1, adjust=False) / pd.stats.moments.ewma(d, com=period-1, adjust=False)
    rs = u.ewm(com=period - 1, adjust=False).mean() / d.ewm(com=period - 1, adjust=False).mean()

    adj_length = adj_length - rs.shape[0]
    filler = adj_length * [rs.iloc[0]]

    rs = pd.concat([pd.Series(filler), rs])
    return 100 - 100 / (1 + rs)


def build_ratio_data(study_ticker, base_ticker):
    # print(study_ticker)
    base_file = [x for x in os.listdir(DATA_FOLDER) if base_ticker in x][0]
    study_file = [x for x in os.listdir(DATA_FOLDER) if study_ticker in x][0]

    base_df = pd.read_csv(os.path.join(DATA_FOLDER, base_file))
    # print(base_df.head())
    study_df = pd.read_csv(os.path.join(DATA_FOLDER, study_file))
    # print(study_df.head())

    base_df.dropna(axis=0, how='any', inplace=True)
    study_df.dropna(axis=0, how='any', inplace=True)

    # chk_gap(base_ticker, base_df)
    # chk_gap(study_ticker, study_df)

    base_df.Date = base_df.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())
    study_df.Date = study_df.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())

    start_date = max(base_df.Date.iloc[0], study_df.Date.iloc[0])
    base_df = base_df.loc[base_df.Date >= start_date]
    study_df = study_df.loc[study_df.Date >= start_date]

    end_date = min(base_df.Date.iloc[-1], study_df.Date.iloc[-1])
    base_df = base_df.loc[base_df.Date <= end_date]
    study_df = study_df.loc[study_df.Date <= end_date]

    base_df.reset_index(drop=True, inplace=True)
    study_df.reset_index(drop=True, inplace=True)
    # print(start_date, end_date, base_df.shape, study_df.shape)

    drops = []
    if base_df.shape[0] != study_df.shape[0]:
        for d in base_df.Date:
            if d not in drops and d not in study_df.Date.values:
                drops.append(d)

        # print('drops', len(drops))
        for d in study_df.Date:
            if d not in drops and d not in base_df.Date.values:
                drops.append(d)

        # seems inefficient, so do more research...
        # print('drops', len(drops))
        for d in drops:
            base_df = base_df.loc[base_df.Date != d]
            study_df = study_df.loc[study_df.Date != d]

    base_df.reset_index(drop=True, inplace=True)
    study_df.reset_index(drop=True, inplace=True)
    # print(start_date, end_date, base_df.shape, study_df.shape)

    base_df = calc_ma(base_df, field='Adj_Close')
    study_df = calc_ma(study_df, field='Adj_Close')

    # set up ratio df
    output_cols = ['dates', 'Adj_Close_ratio']
    df_out = pd.DataFrame(columns=output_cols)
    df_out['dates'] = base_df.Date
    df_out['Adj_Close_ratio'] = study_df.Adj_Close / base_df.Adj_Close
    df_out = calc_ma(df_out, field='Adj_Close_ratio')
    df_out.rename(mapper={'ma_one_mo_Adj_Close_ratio': 'one_mo_ratio',
                          'ma_short_Adj_Close_ratio': 'short_ratio',
                          'ma_long_Adj_Close_ratio': 'long_ratio'},
                  axis=1,
                  inplace=True)

    # df_out['rsi'] = compute_rsi(df_out['Adj_Close_ratio']).values  # values fixes indexing issue
    # df_out['rsi_long'] = compute_rsi(df_out['Adj_Close_ratio'], time_window=RSI_LONG).values

    df_out['rsi'] = RSI(series=df_out['Adj_Close_ratio'], period=RSI_SHORT).values
    df_out['rsi_long'] = RSI(df_out['Adj_Close_ratio'], period=RSI_LONG).values

    return df_out


def make_long_term_ratio_chart(df_out, study_ticker, study_desc, base_ticker, base_desc):
    fig, (ax0, ax1) = plt.subplots(2, 1, gridspec_kw={'height_ratios': [PLT_HT_MAJOR, PLT_HT_MINOR]}, figsize=(PLT_WIDTH, PLT_HT))

    ax0.plot(df_out.dates, df_out.one_mo_ratio, color='tab:grey', label=f'20 Day MA')
    ax0.plot(df_out.dates, df_out.short_ratio, color='tab:blue', label=f'{TERM_SHORT} {ABBR_FOR_AVG} MA')
    ax0.plot(df_out.dates, df_out.long_ratio, color='tab:orange', label=f'{TERM_LONG} {ABBR_FOR_AVG} MA')
    ax0.text(0.75, 0.03, 'Plot by @FlyingFish365 on Twitter', fontsize=6, transform=ax0.transAxes)
    ax0.set_xlabel('Date', fontsize=8)
    ax0.set_title(f'{study_ticker} ({study_desc}) vs {base_ticker} ({base_desc}) Moving Avg Price Ratios', fontsize=8)
    ax0.legend(loc='upper left', fontsize=6)
    ax0.margins(0.0)
    ax0.set_yscale(value='log')
    ax0.set_yticks(ticks=[])
    ax0.tick_params(axis='both', labelsize=6)

    ax1.plot(df_out.dates, df_out['rsi_long'], color='tab:blue')
    ax1.hlines((30, 70), xmin=df_out.dates.iloc[0], xmax=df_out.dates.iloc[-1], colors=('tab:grey', 'tab:grey'), linestyles='dashed')
    ax1.set_ylim(0, 100)
    ax1.tick_params(axis='y', labelsize=6)
    ax1.set_xticks(ticks=[])
    ax1.set_title(f'RSI ({RSI_LONG} trading days)', fontsize=8)
    ax1.margins(0.0)

    fig.tight_layout()

    if SHOW_PLOTS:
        plt.show()

    return fig


def make_short_term_ratio_chart(df_out, study_ticker, study_desc, base_ticker, base_desc):
    term_days = 250
    fig, (ax0, ax1) = plt.subplots(2, 1, gridspec_kw={'height_ratios': [PLT_HT_MAJOR, PLT_HT_MINOR]}, figsize=(PLT_WIDTH, PLT_HT))

    ax0.plot(df_out.dates.iloc[-term_days:], df_out['Adj_Close_ratio'].iloc[-250:], color='tab:green', label='Daily Price Ratio')
    ax0.plot(df_out.dates.iloc[-term_days:], df_out['one_mo_ratio'].iloc[-250:], color='tab:grey', label=f'20 Day MA')
    ax0.plot(df_out.dates.iloc[-term_days:], df_out['short_ratio'].iloc[-250:], color='tab:blue', label=f'{TERM_SHORT} {ABBR_FOR_AVG} MA')
    ax0.text(0.75, 0.03, 'Plot by @FlyingFish365 on Twitter', fontsize=6, transform=ax0.transAxes)
    ax0.set_xlabel('Date', fontsize=8)
    ax0.set_title(f'{study_ticker} ({study_desc}) vs {base_ticker} ({base_desc}) Moving Avg Price Ratios (1 yr)', fontsize=8)
    ax0.legend(loc='upper left', fontsize=6)
    ax0.margins(0.0)
    ax0.margins(0.0)
    ax0.set_yscale(value='log')
    ax0.set_yticks(ticks=[])
    ax0.tick_params(axis='both', labelsize=6)

    ax1.plot(df_out.dates.iloc[-term_days:], df_out['rsi'].iloc[-term_days:], color='tab:blue')
    ax1.hlines((30, 70), xmin=df_out.dates.iloc[-term_days], xmax=df_out.dates.iloc[-1], colors=('tab:grey', 'tab:grey'), linestyles='dashed')
    ax1.set_ylim(0, 100)
    ax1.tick_params(axis='y', labelsize=6)
    ax1.set_xticks(ticks=[])
    ax1.set_title(f'RSI ({RSI_SHORT} trading days)', fontsize=8)
    ax1.margins(0.0)

    fig.tight_layout()

    if SHOW_PLOTS:
        plt.show()

    return fig


    # make a plot
def make_basic_ratio_chart(df_out, study_ticker, study_desc, base_ticker, base_desc, term_days=250):
    fig = plt.figure(figsize=(5.5, 3))
    plt.plot(df_out.dates, df_out.one_mo_ratio, color='tab:grey', label=f'20 Day MA')
    plt.plot(df_out.dates, df_out.short_ratio, color='tab:blue', label=f'Short Term MA: {TERM_SHORT} {ABBR_FOR_AVG}')
    plt.plot(df_out.dates, df_out.long_ratio, color='tab:orange', label=f'Long Term MA: {TERM_LONG} {ABBR_FOR_AVG}')
    plt.figtext(0.7, 0.2, 'Plot by @FlyingFish365 on Twitter', fontsize=6)
    plt.title(f'{study_ticker} ({study_desc}) vs {base_ticker} ({base_desc}) Moving Avg Price Ratios', fontsize=8)
    plt.tick_params(axis='y', which='both', right=False, left=False, labelleft=False)
    plt.xlabel('Date', fontsize=8)
    plt.legend(loc='upper left', fontsize=6)
    plt.yscale(value='log')
    fig.tight_layout()

    if SHOW_PLOTS:
        plt.show()

    # print(df_out['dates'].iloc[0], df_out['dates'].iloc[-1])
    return fig

def write_intro_sheet(workbook):
    worksheet = workbook.add_worksheet('Description')
    worksheet.set_column(0, 0, 120)
    worksheet.write(0, 0, INTRODUCTION)
    worksheet.write(1, 0, DISCLAIMER)
    worksheet.write(2, 0, CHARTING)
    worksheet.write(3, 0, INTERPRETATION)
    worksheet.write(4, 0, BASES)
    worksheet.write(5, 0, CONTACT)
    worksheet.write(6, 0, CHANGES)
    worksheet.write(7, 0, PYTHON_NOTES)
    worksheet.write(8, 0, LICENSE)
    worksheet.write(9, 0, CITATIONS)

if __name__ == "__main__":
    filename = f'ETF_Study_ver_{__VERSION__}_{datetime.date.today().isoformat()}.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book

    if UPDATE_FILES:
        update_tickers(remove=DELETE_OBS_FILES)

    # ref_portfolio(tkers=['SPY', 'FBND'], allocs=[0.6, 0.4])  # not sure how well this works yet

    if MAKE_PLOT:
        if WRITE_XL:
            write_intro_sheet(workbook)

        # pages for the list of base tickers - long term
        for base in BASE_TKRS:
            base_desc = TKR_DICT[base]
            worksheet = workbook.add_worksheet(f'Base Tkr {base} Long Term')
            row = 0
            col = 0

            for ticker, desc in TKR_DICT.items():
                if base == ticker or (base == 'GC=F' and ticker in ['GLD', 'IAU', 'PHYS', 'OUNZ']):
                    continue

                print(f'Working {ticker} vs {base} Long Term')
                df_ratio = build_ratio_data(ticker, base)

                # return the chart - put in excel
                fig = make_long_term_ratio_chart(df_ratio, study_ticker=ticker, study_desc=desc, base_ticker=base, base_desc=base_desc)

                if WRITE_XL:
                    img_data = io.BytesIO()
                    fig.savefig(img_data, format="png")
                    img_data.seek(0)
                    worksheet.set_row(row=row, height=CELL_HT)
                    worksheet.set_column(first_col=col, last_col=col, width=CELL_WIDTH)
                    worksheet.insert_image(row, col, "", {'image_data': img_data})

                plt.close(fig)

                col += 1
                if col >= 3:
                    col = 0
                    row += 1

        # page for selected ticker pairs
        row = 0
        col = 0
        worksheet = workbook.add_worksheet(f'Ticker Pairs - Long Term')
        for base, ticker in PAIR_TKRS.items():
            base_desc = TKR_DICT[base]
            tkr_desc = TKR_DICT[ticker]
            print(f'Working {ticker} vs {base}')

            df_ratio = build_ratio_data(ticker, base)

            # return the chart - put in excel
            fig = make_long_term_ratio_chart(df_ratio, study_ticker=ticker, study_desc=tkr_desc, base_ticker=base,base_desc=base_desc)

            if WRITE_XL:
                img_data = io.BytesIO()
                fig.savefig(img_data, format="png")
                img_data.seek(0)
                worksheet.set_row(row=row, height=CELL_HT)
                worksheet.set_column(first_col=col, last_col=col, width=CELL_WIDTH)
                worksheet.insert_image(row, col, "", {'image_data': img_data})

            plt.close(fig)

            col += 1
            if col >= 2:
                col = 0
                row += 1

        # pages for the list of base tickers - short term
        for base in BASE_TKRS:
            base_desc = TKR_DICT[base]
            worksheet = workbook.add_worksheet(f'Base Tkr {base} Short Term')
            row = 0
            col = 0

            for ticker, desc in TKR_DICT.items():
                if base == ticker or (base == 'GC=F' and ticker in ['GLD', 'IAU', 'PHYS', 'OUNZ']):
                    continue

                print(f'Working {ticker} vs {base} Short Term')
                df_ratio = build_ratio_data(ticker, base)

                # return the chart - put in excel
                fig = make_short_term_ratio_chart(df_ratio, study_ticker=ticker, study_desc=desc, base_ticker=base, base_desc=base_desc)

                if WRITE_XL:
                    img_data = io.BytesIO()
                    fig.savefig(img_data, format="png")
                    img_data.seek(0)
                    worksheet.set_row(row=row, height=CELL_HT)
                    worksheet.set_column(first_col=col, last_col=col, width=CELL_WIDTH)
                    worksheet.insert_image(row, col, "", {'image_data': img_data})

                plt.close(fig)

                col += 1
                if col >= 3:
                    col = 0
                    row += 1

    if WRITE_XL and writer is not None:
        try:
            writer.save()
        except:
            print('!!!Cannot save Excel file!!!')
    else:
        print(f'!!!Cannot save Excel file, not set to save, or writer problem, WRITE_XL: {WRITE_XL}!!!')

    print('Done')
