# -*- coding: utf-8 -*-

"""See LICENSE below for authorship, copyright and license information"""
# https://github.com/Vettejeep/Financial_Stuff

# TODO: new percent change charts not fully validated, but the code appears correct

# play:
# qqq, but rsi divergence
# iyt maybe
# pave
# icln
# pbw
# smh
# botz
# robo
# qqew
# iyw
# icvt
# all gold and silver

import io
import os
import csv
import time
import datetime
import traceback
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
from urllib.request import urlopen  # Request,
pd.options.mode.chained_assignment = None  # default='warn'
pd.set_option('display.max_columns', 32)
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()

__VERSION__ = '2.1.0'

UPDATE_FILES = False  ### Set True to create the needed ticker files
MAKE_PLOT = True
SHOW_PLOTS = False
WRITE_XL = True
WRITE_PAIR_PLOTS = True
LOG_CHK = False
DELETE_OBS_FILES = False

# does not currently handle back dating or adjust to forward dating, if this is changed, delete all csv files
START_DATE = datetime.date(year=1993, month=1, day=1)
END_DATE = datetime.date.today() + datetime.timedelta(days=1)
GITHUB = 'https://github.com/Vettejeep/Financial_Stuff'

MIN_SIZE = 250
ABBR_FOR_AVG = 'months'
TERM_SHORT = 3
TERM_LONG = 42
ONE_MN_AVG = int(20)
SHORT_AVG = int(TERM_SHORT*4.33*5)
LONG_AVG = int(TERM_LONG*4.33*5)
RSI_SHORT = 14
RSI_LONG = 70
PLT_WIDTH = 7
PLT_HT = 4
PLT_HT_MAJOR = 3
PLT_HT_MINOR = 1
CELL_WIDTH = PLT_WIDTH * 13.625
CELL_HT = PLT_HT * 75
TAB = ['tab:blue', 'tab:orange', 'tab:green', 'tab:red', 'tab:purple', 'tab:brown', 'tab:pink', 'tab:gray', 'tab:olive', 'tab:cyan']

DATA_FOLDER = r'D:\WebScrape\raw'
TIMEOUT_SEC = 10
SLEEP_SEC = 1  # do not know yahoo policy, but a timeout may help to keep them from blocking us
COLS = ['Date', 'Open', 'High', 'Low', 'Close', 'Adj_Close', 'Volume']
DT = datetime.datetime.now()
DATE_TIME_NOW = DT.strftime('%Y-%m-%d %H:%M')
LOG_FILE_NAME = f'asset_compare_log_{DT.strftime("%Y_%m_%d_%H_%M")}.txt'
ERROR_LOG_FILE_LENGTH = 2500
ERROR_LOG_FILENAME = f'$asset_compare_error_log.txt'

INTRODUCTION = f"""Introduction & What is this?
This is a historical study and not investment advice. Provided here are:
Historical data charted for a list of asset ticker symbols from Yahoo finance, generally ETFs, futures and stocks.
Plots of the price ratio over time between two securities as sets of moving averages to smooth out noise.
The price ratio, plotted over time, tells you if the study security has performed better or worse than the base asset.
Charts showing percentage change over time for various sets of assets are also included.
Only assets traded on US exchanges are included, but funds may hold global assets in their portfolios, or be
primarily outside the US in their holdings.

I believe that the charts here are valuable information for assisting in investment decisions.
ETFs can be bought or sold like stocks and can represent market indexes that otherwise cannot be directly invested in.
For ticker info just do an internet search for something like: <ticker> stock quote.
The different sheets of this Excel file contain plots, mostly organized by the base asset ticker and time frame.
Updated Excel output files are available at: {GITHUB}

Open Office appears to be currently doing a better job of displaying the file. Not all table cell sizes are being
set in all versions of Excel. This is true even though the Python library in-use is designed for Excel. Users can
resize the cells while a solution is worked on.

This file is the output of a program run started at {DATE_TIME_NOW}."""

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
they may be included here for study purposes.

Past historical performance is no guarantee of future performance, investment return or asset price changes.

Problems with Yahoo Finance have been observed where the web site will not give historical data before the most
recent trading day. This has been observed both in a browser and in this code. At present, not all possible bugs
associated with this problem have been fixed. Check the ticker 'csv' files for very small file sizes (1 kb),
this is often a sign that there is a problem. Adequate error handling is still in progress, so this can cause
a program crash. This code is very much a work in progress and not fully debugged or complete in handling all error 
conditions, especially those coming from poor responses from Yahoo Finance. Currently, Yahoo appears to be doing
better and the bug has not been observed in the last few weeks, so it may have been fixed by Yahoo."""

CHARTING = """
Charting:
The charts here can be a portion of the useful information for assisting in investment decisions.
You also need information about world events, politics, macro economics and specific economic factors related to the 
asset classes and/or ETFs, futures or stocks that you might purchase or go short on, none of that is presented here.
It is also valuable to have and understand technically oriented charts for assets you are considering investing
in, these types of charts are not included in this study.

Moving averages may not be fully valid for the earliest dates in the plots, especially for the longest term average.
Plots may go back to 1996, but it depends on when the asset was created or Yahoo Finance data limitations

All charts use "adjusted close" as reported by Yahoo Finance. Yahoo Finance says this accounts for dividends and 
splits. Code to detect splits and dividends for new data has been added, in this case the whole timeline
for that specific security is re-created because the adjusted close has been recalculated for all dates by
Yahoo Finance."""

INTERPRETATION = """
Interpretation of Ratio Charts:
If the moving averages are going up, then the study security is out-performing the base security.
If the moving averages are going down, then the study security is under-performing the base security.
Some people consider cross-overs of the shorter term to longer term moving averages to be buy & sell signals,  
other trend interpretations exist.
Daily price ratio has been added to the short term plots, as a reference.

Many of the plots are based on moving averages, thus their information may be behind current market behavior.
Past performance is no guarantee of future performance, trends do change over time.
As an investor, one hopes to find trends that will remain in place for years, many, but not all, seem to do so.
Ratio charts do not tell if asset prices are, or will be, increasing or decreasing, they may both be falling together, 
for example."""

PCT_CHARTS = """
Charts of percentage change over time for selected asset groups have been added. These charts start at a time past,
normalize comparative asset prices to 100, and thus plot the percentage change from that date. These plots do show
which assets have been gaining or losing value. Currently these charts look approximately 6 and 12 months back, and a 
set of charts are provided from the March 2020 S&P 500 low."""

BASES = """
Important Base Tickers:
SPY: ETF that is a proxy for the S&P 500
GC=F: Gold Front Month Futures"""

CONTACT = """
You are welcome to contact me:
Email: Vettejeep365@gmail.com
Twitter: @FlyingFish365 - please follow and like on Twitter and I will work to expand this project and keep this 
effort reasonably current. If you follow on Twitter, you will get notification of updates.

The ratio charts don't change quickly, so updating the Excel output file every 2-4 weeks seems about right. My plan is 
to post updates from time to time and post them on the Github site.

Community discussion of these charts and investment ideas on Twitter is encouraged.
RSI has been a challenge for the longer term plots - specifically what time period to use.

Adding more asset classes (additional ticker symbols) is very easy to do. They need to be available on Yahoo Finance.
For added tickers it is best if they have existed for 5+ years so the longer term averages are somewhat reasonable.

Thanks for any kudos, or suggestions for the improvement of this study,
Kevin"""

PAST_CHANGES = """
As of 2020-08-16
1. Added RSI
2. Added short term plots (1 year)
3. Added some tickers
As of 2020-08-17
1. Added ability to turn off the pair plots, makes it easier to comment out tickers if desired.
2. Added more crypto ETFs
As of 2020-08-22
1. Add charts of percentage change over time for selected asset groups
"""

CHANGES = """
Recent Changes:
As of 2020-10-02
1. Code checks for dividends and splits, rebuilding the ticker file if needed due to the change in adjusted close.
As of 2020-10-11
1. Add Gold to Silver Ratio Chart - important for gold and silver investors"""

PYTHON_NOTES = f"""
Python: 
This Excel file was generated by a Python script: {os.path.basename(__file__)}, code version {__VERSION__}.
The source code is posted on Github as open source at {GITHUB}.

It is probably necessary to understand at least some Python in order to work with the code, plus the web scrape and
numeric library stack for Python.  This is still a rough piece of code, not all paths tested, so expect bugs.
Limited exception handling - will be adding more of this to prevent crashes.

Designed for a reasonably up-to-date Anaconda Python 3 distribution.

It takes a very long time (hours) the first time it is run in order to obtain all needed historical data.
After historical data has been acquired, updates to the ticker files are reasonably fast if done at least monthly.

Usage (it helps to know Python to figure it out):
Create needed directories manually (to be fixed).
Adjust the settings on the ALL_CAPS variables to vary script operation.
Set UPDATE_FILES = True to update or create the ticker files.
Set your lists of tickers for acquisition and study, or use my settings.
Run the script using Python...
- not written for non-programmers to use, not user-friendly.
- it is pretty simple code, but, you will need to read the Python...
- Non-programmers can usually find a reasonably up-to-date Excel file on the Github site"""

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
{GITHUB}, and the above license is adhered to for code and software."""

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
AIQ: AI and tech ETF
"""

# https://www.marketwatch.com/story/heres-why-90-of-rich-people-squander-their-fortunes-2017-04-23

# TODO: try/except blocks so it can recover and do what it can, esp scraping tickers
# TODO: validate base & study tickers so it does not crash if the file for the ticker does not exist
# TODO: warn/retry if a month comes up missing in web scrape, rare issue, data gap plots a hole now
# TODO: verify time zone issues generating time stamps
# TODO: go to 3 month internet data collection to speed up building the data files - ok with yahoo limitations
# TODO: automate sub directories
# TODO: crashes if ticker has inadequate data for RSI, need exception handler
# TODO: dynamically adjust plot legend and text based on data so data is not over-written
# TODO: font sizes and axes tick value eliminations on charts not working as expected, research matplotlib axes needed
# TODO: change percentage charts to specific 6 month and 1 year look back dates, is now 125 and 250 days
# TODO: add some comment to percent change charts for charts that reference things like the march 2020 low

BASE_TKRS = ['SPY', 'GC=F']
PAIR_TKRS = {
    'SPY': 'RSP',
    'JKE': 'JKF',
    'SI=F': 'GC=F',
    'GC=F': 'SPY',
}

# tickers, along with a very short description that will fit in the chart title
# these are the ticker symbols that will be web scraped and updated in files
TKR_DICT = {
    # 'TSLA': 'Tesla',
    'SPY': 'S&P 500',
    'GC=F': 'Gold Futures',
    'QQQ': 'Nasdaq 100',
    'RSP': 'Eq Wt S&P 500',

    'IEF': 'US 7-10 Yr Treasurys',
    'GOVT': 'US Treasury ETF',
    'LQD': 'Corporate Bond ETF',

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
    'FBND': 'US Bond Index',

    'PPLT': 'Platinum Metal',  # comment out these & below for a short test run
    'CPER': 'Copper Metal',
    'EEM': 'Emerging Mkt',
    'IDEV': 'Dev Mkt Ex-US',

    # 'EWJ': 'Japan',
    # 'EWG': 'Germany',  # RSI problem for EWG? needs debug
    # 'EWU': 'UK',

    'IJH': 'US Mid-Cap',
    'IJR': 'US Small Cap',
    'IYR': 'Commercial RE',
    'CCJ': 'Uranium Stock',  # URNM ETF not in existence long enough
    'URPTF': 'Uranium Metal',  # or URPTF in the US
    'MOO': 'Agri-Business',
    'DBC': 'Broad Commodities',
    'DBA': 'Ag commodities',
    'PAVE': 'Infrastructure',
    'ICLN': 'Clean Energy',
    'PBW': 'Clean Energy',
    'SMH': 'Semi-Conductor',
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
    'ETHE': 'Ethereum ETF',

    # 'GDLC': 'Large Cap Crypto ETF', # not enough data, part of the need for more checks and exception handlers
    # 'BTC-USD': 'Bitcoin',
    # 'ETH-USD': 'Ethereum',  # odd results with crypto at the moment, date range issues in web scrape

    'BNO': 'Brent Oil ETF',
}

# TKR_DICT = {
# 'SPY': 'S&P 500',
# 'FBND': 'US Bond Index',
# 'QQQ': 'Nasdaq 100',}

# platinum - some have poor financials
platinum = ['NGLOY', 'IMPUY', 'PLG', 'SBSW', 'NMPNF', 'IVPAF', 'ELRFF', 'NKORF', 'NMTLF']
uranium = ['CCJ', 'NXE', 'UEC', 'URG', 'ISENF', 'AZZUF', 'DNN', 'FCUUF', 'WSTRF']


# TODO: is doing some run logging in addition to error logs, split out run log.
def log_error(msg, filename=ERROR_LOG_FILENAME):
    """
    log errors to a file, call from most exception blocks, limits log file to 2000 lines
    :param msg: string, exception or other message to log
    :param filename: string, file name to log to
    :return: None
    """
    contents = []

    try:
        failure_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        msg = 'Time: %s' % failure_time + '\n' + msg
        print(msg)

        with open(filename, 'r') as f:
            line = f.readline()
            while line:
                contents.append(line)
                line = f.readline()

        if len(contents) > ERROR_LOG_FILE_LENGTH:
            st = len(contents) - ERROR_LOG_FILE_LENGTH
            contents = contents[st:]
            contents = [f'Log file Contents Trimmed to a Length of: {len(contents)} approximate lines at {failure_time}\n\n'] + contents

    except:
        pass

    try:
        contents.append(msg)
        contents.append('\n')
        with open(filename, 'w') as f:
            for entry in contents:
                f.write(entry)
    except:
        pass


def replace_comma(x):
    x = str(x)
    return x.replace(",", "")


def remove_obsolete_tickers():
    files = [x for x in os.listdir(DATA_FOLDER)]

    # remove if not in list
    for f in files:
        if f.split('.')[0] not in TKR_DICT.keys():
            os.remove(os.path.join(DATA_FOLDER, f))

def set_date_range(fp):
    start_date = START_DATE.replace(day=1)
    end_date = (start_date + datetime.timedelta(days=32)).replace(day=1)  # go three months in advance
    path_exists = os.path.exists(fp)

    try:
        # if it exists,
        if path_exists:
            df_tkr = pd.read_csv(fp)

            if df_tkr.shape[0] > 0:
                df_tkr.Date = df_tkr.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())
                start_date = df_tkr.Date.iloc[-1] + datetime.timedelta(days=1)

                if start_date < END_DATE:
                    end_date = (start_date + datetime.timedelta(days=32)).replace(day=1)

    except:
        msg = f'Error in set_date_range for {fp}\n'
        msg += traceback.format_exc()
        log_error(msg)

    return start_date, end_date, path_exists


def get_html_table(ticker, start_date, end_date):
    table = None

    try:
        # .replace(tzinfo=tz.gettz(ny_time_zone)). ?
        start_ts = int(datetime.datetime.combine(start_date, datetime.datetime.min.time()).timestamp())
        end_ts = int(datetime.datetime.combine(end_date, datetime.datetime.min.time()).timestamp() - 10)
        url = f'https://finance.yahoo.com/quote/{ticker}/history?period1={start_ts}&period2={end_ts}&interval=1d&filter=history&frequency=1d'
        print(url)
        print(start_ts, end_ts)

        ct = 0

        while table is None and ct < 6:
            try:
                page = urlopen(url, timeout=TIMEOUT_SEC)
                soup = BeautifulSoup(page, features="lxml")
                table = soup.find('table')

            except:
                print(
                    f'exception sleep {SLEEP_SEC} seconds, Start Date: {start_date}, End Date: {end_date}, Count: {ct}')
                time.sleep(SLEEP_SEC)
                table = None

            finally:
                ct += 1

    except:
        msg = f'Error in get_html_table for {ticker} for dates {start_date} to {end_date}\n'
        msg += traceback.format_exc()
        log_error(msg)

    return table

def table_to_dataframe(table, ticker, start_date, end_date):
    try:
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

        if df.shape[0] > 0:
            df = df.dropna(axis=0, how='any', subset=['Date', 'Open'], inplace=False)

            try:
                df['Date'] = df['Date'].apply(lambda x: datetime.datetime.strptime(x, '%b %d, %Y').date())

            # when it fails, need to go row by row?
            except:
                df.reset_index(inplace=True, drop=True)
                drops = []
                dates = df['Date'].values

                for i in df.index:
                    try:
                        dates[i] = datetime.datetime.strptime(dates[i], '%b %d, %Y').date()
                    except:
                        drops.append(i)

                df.drop(index=drops, axis=0, inplace=True)

            df.sort_values(by='Date', axis=0, ascending=True, inplace=True)
            df.reset_index(inplace=True, drop=True)

        else:
            df = None

    except:
        df = None
        msg = f'Error in table_to_dataframe for {ticker} for dates {start_date} to {end_date}\n'
        msg += traceback.format_exc()
        print(msg)
        log_error(msg)

    return df

def update_ticker(ticker, start_date, end_date, path_exists, remove=False):
    df_tkr = None

    try:
        fp = os.path.join(DATA_FOLDER, ticker + '.csv')
        if os.path.exists(fp):
            if remove:
                os.remove(fp)
            else:
                df_tkr = pd.read_csv(fp)
                df_tkr.Date = df_tkr.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())

        while start_date < END_DATE:
            table = get_html_table(ticker, start_date, end_date)
            df = None

            if table is not None:
                df = table_to_dataframe(table, ticker, start_date, end_date)
                if df is not None:
                    print(df.head(3))
                    print(df.tail(3))
                    print()

            if df is not None:
                has_new_dividends_splits = df.isna().any().any()

                if not remove and (start_date != START_DATE and has_new_dividends_splits):
                    start_date = START_DATE.replace(day=1)
                    end_date = (start_date + datetime.timedelta(days=32)).replace(day=1)  # go three months in advance
                    return update_ticker(ticker, start_date, end_date, path_exists, remove=True)

                else:
                    df = df.dropna(axis=0, how='any', inplace=False)
                    df.sort_values(by='Date', inplace=True)
                    df.reset_index(inplace=True, drop=True)

                    if df.shape[0] > 0:
                        if df_tkr is None:
                            df_tkr = df
                        else:
                            df_tkr = df_tkr.append(df)

            start_date = (start_date + datetime.timedelta(days=32)).replace(day=1)
            end_date = (end_date.replace(day=28) + datetime.timedelta(days=32)).replace(day=1)

            if df_tkr is not None:
                df.sort_values(by='Date', inplace=True)
                df.reset_index(inplace=True, drop=True)

    except:
        df_tkr = None
        msg = f'Error in update_ticker for {ticker} for dates {start_date} to {end_date}, remove: {remove}\n'
        msg += traceback.format_exc()
        log_error(msg)

    return df_tkr


# TODO: check if start date in file, currently does not work if an earlier date specified
def update_tickers(remove=False):
    t = ''
    try:
        # remove if not in list
        if remove:
            remove_obsolete_tickers()

        ran_tkr = []

        for t in TKR_DICT.keys():
            print(f'starting work on {t}')
            fp = os.path.join(DATA_FOLDER, t + '.csv')
            start_date, end_date, path_exists = set_date_range(fp)
            # print(start_date, type(start_date), end_date, type(end_date), path_exists)

            df_tkr = update_ticker(t, start_date, end_date, path_exists, remove=False)

            # print(start_date, end_date)
            if df_tkr is not None:
                df_tkr.reset_index(drop=True, inplace=True)
                # even when past the last date available, Yahoo seems to return one value sometimes, so drop duplicates
                df_tkr.replace(to_replace='-', value=0, inplace=True)
                df_tkr.Open = df_tkr.Open.apply(lambda x: replace_comma(x)).astype(np.float32)
                df_tkr.High = df_tkr.High.apply(lambda x: replace_comma(x)).astype(np.float32)
                df_tkr.Low = df_tkr.Low.apply(lambda x: replace_comma(x)).astype(np.float32)
                df_tkr.Close = df_tkr.Close.apply(lambda x: replace_comma(x)).astype(np.float32)
                df_tkr['Adj_Close'] = df_tkr['Adj_Close'].apply(lambda x: replace_comma(x)).astype(np.float32)
                df_tkr.Volume = df_tkr.Volume.apply(lambda x: replace_comma(x)).astype(np.uint64).fillna(0)
                df_tkr = df_tkr.loc[df_tkr.Close != 0.0,]  # futures get 0's on Sunday
                df_tkr.drop_duplicates(subset=['Date'], inplace=True)

                max_days_gap = df_tkr.Date.diff().max()
                # print(max_days_gap, type(max_days_gap))
                GAP_DAYS = 14
                ran_tkr.append(t)

                if not(type(max_days_gap) == datetime.timedelta or type(max_days_gap) == pd._libs.tslibs.timedeltas.Timedelta):
                    msg = f'max days gap problem for {t}, {max_days_gap}, {type(max_days_gap)}'
                    df_tkr.to_csv(t + '.csv', quoting=csv.QUOTE_NONE, index=False, float_format='%.2f')
                    log_error(msg)
                elif df_tkr.shape[0] >= MIN_SIZE and max_days_gap < datetime.timedelta(days=GAP_DAYS):  # 9/11/2001 closure seems to be 7 days
                    df_tkr.reset_index(drop=True, inplace=True)
                    df_tkr.to_csv(fp, quoting=csv.QUOTE_NONE, index=False, float_format='%.2f')
                    # print(df_tkr.tail(10))
                    msg = f'data frame saved for {t}'
                    log_error(msg)
                elif max_days_gap >= datetime.timedelta(days=GAP_DAYS):
                    msg = f'Gap Days: {max_days_gap}, Type: {type(max_days_gap)}\n'
                    msg += f'In, update_tickers, data frame not saved for {t}, gap too big: {max_days_gap} days.\n'
                    df_tkr.to_csv(t + '.csv', quoting=csv.QUOTE_NONE, index=False, float_format='%.2f')
                    log_error(msg)
                else:
                    msg = f'Gap Days: {max_days_gap}, Type: {type(max_days_gap)}\n'
                    msg += f'In, update_tickers, data frame not saved for {t}, too small, size is {df_tkr.shape[0]}\n'
                    log_error(msg)

            else:
                msg = f'In update_tickers, {t} unable to be written ({df_tkr})'
                log_error(msg)

        log_error(f'Actual run list: {ran_tkr}')

    except:
        msg = f'Error in update_tickers for {t}.\n'
        msg += traceback.format_exc() + '\n'
        log_error(msg)


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
    df0 = None
    df1 = None
    try:
        date_now = datetime.date.today()
        date_start = (datetime.datetime.now() - datetime.timedelta(days=366+32)).date().replace(day=1)
        returns = []
        ytd_returns = []
        pr = []
        dt = []
        date_start_month_end = (date_start + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)

        while date_start_month_end <= date_now:
            date_end_month_end = (date_start + datetime.timedelta(days=64)).replace(day=1) - datetime.timedelta(days=1)
            return_prop = 0.0
            ytd_return_prop = 0.0

            for tkr, alloc in zip(tkers, allocs):
                base_file = [x for x in os.listdir(DATA_FOLDER) if tkr in x][0]
                df = pd.read_csv(os.path.join(DATA_FOLDER, base_file))
                df.Date = df.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())

                df0 = df.loc[df.Date <= date_start_month_end, ].reset_index(drop=True, inplace=False)
                base_pr = df0.Adj_Close.iloc[-1]

                df1 = df.loc[df.Date <= date_end_month_end,].reset_index(drop=True, inplace=False)
                end_pr = df1.Adj_Close.iloc[-1]

                # =0.12*((320.98/310.52)-1)
                return_prop += alloc * (end_pr / base_pr)
                pr.append([base_pr, end_pr])

                if date_end_month_end.year == date_now.year:
                    ytd_return_prop += alloc * (end_pr / base_pr)

                # print(date_start_month_end, date_end_month_end, alloc, round(base_pr, 2), round(end_pr, 2), round(return_prop, 6))

            if date_end_month_end.year == date_now.year:
                ytd_returns.append(ytd_return_prop)
                # print(base_pr, end_pr)
                wkg_date = df0.Date.iloc[-1]
                if df1.shape[0] > 0:
                    wkg_date = df1.Date.iloc[0]
                dt.append([df0.Date.iloc[0], wkg_date])
                # print(df0.Date.iloc[0], wkg_date)
                # print()

            returns.append(return_prop)

            # 1 yr return: 22.0033%, ytd return: 9.806%, 3 month return: 15.456%, current ytd return: 3.8264% vs 4.24% at fidelity
            # print(round(returns[-1], 3), date_start, date_end)
            date_start = (date_start + datetime.timedelta(days=32)).replace(day=1)
            date_start_month_end = (date_start + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)

        returns = np.array(returns)
        ytd_returns = np.array(ytd_returns)

        # print(returns)
        # print(ytd_returns)
        # print(dt)
        # print(pr)
        total_ret = np.round(((np.prod(returns[: -1]) - 1.0) * 100.0), 4)
        ytd_ret = np.round(((np.prod(ytd_returns[: -1]) - 1.0) * 100.0), 4)
        cur_ytd_ret = np.round(((np.prod(ytd_returns) - 1.0) * 100.0), 4)
        mo3_ytd_ret = np.round(((np.prod(returns[-4: -1]) - 1.0) * 100.0), 4)

        print(f'Reference Portfolio for {tkers} in allocations {allocs}, all except current are to end of previous month.')
        print(f'1 yr return: {total_ret}%, ytd return: {ytd_ret}%, 3 month return: {mo3_ytd_ret}%, current ytd return: {cur_ytd_ret}%')
        print()

    except:
        msg = f'Error in ref_portfolio for {tkers}.\n'
        msg += traceback.format_exc() + '\n'
        log_error(msg)


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
def RSI(series, ticker, period=RSI_SHORT):

    try:
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

    except:
        msg = f'Error in RSI for {ticker}.\n'
        msg += traceback.format_exc() + '\n'
        log_error(msg)
        return np.zeros(series.shape[0])


def build_ratio_data(study_ticker, base_ticker):
    df_out = None
    try:
        # print(study_ticker)
        try:
            base_file = [x for x in os.listdir(DATA_FOLDER) if base_ticker in x][0]
            study_file = [x for x in os.listdir(DATA_FOLDER) if study_ticker in x][0]
        except:
            msg = f'Error in build build_ratio_data, either {study_ticker} or {base_ticker} does not exist.\n'
            log_error(msg)

        base_df = pd.read_csv(os.path.join(DATA_FOLDER, base_file))
        # print(base_df.head())
        study_df = pd.read_csv(os.path.join(DATA_FOLDER, study_file))
        # print(study_df.head())

        if base_df is not None and base_df.shape[0] >= MIN_SIZE and study_df is not None and study_df.shape[0] >= MIN_SIZE:
            base_df.dropna(axis=0, how='any', inplace=True)
            study_df.dropna(axis=0, how='any', inplace=True)

        if base_df.shape[0] >= MIN_SIZE and study_df.shape[0] >= MIN_SIZE:
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

        if base_df.shape[0] >= MIN_SIZE and study_df.shape[0] >= MIN_SIZE:
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

        if base_df.shape[0] >= MIN_SIZE and study_df.shape[0] >= MIN_SIZE:
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

            tkr_ratio = study_ticker + '/' + base_ticker

            # values fixes indexing issue
            df_out['rsi'] = RSI(series=df_out['Adj_Close_ratio'], ticker=tkr_ratio, period=RSI_SHORT).values
            df_out['rsi_long'] = RSI(df_out['Adj_Close_ratio'], ticker=tkr_ratio, period=RSI_LONG).values

        if df_out is not None and df_out.shape[0] < (MIN_SIZE / 2):
            df_out = None

    except:
        df_out = None
        msg = 'Error in build_ratio_data.\n'
        msg += traceback.format_exc() + '\n'
        # print(msg)
        log_error(msg)
    finally:
        return df_out


def build_simple_ratio_data(study_ticker, base_ticker, min_date, max_date):  # 'one_over_short', 'short_over_long'
    df = build_ratio_data(study_ticker, base_ticker)
    print(df.shape)
    df[f'one_over_short_{study_ticker}'] = (df['one_mo_ratio'] - df['short_ratio']).apply(lambda x: 1 if x >= 0 else 0)
    df[f'short_over_long_{study_ticker}'] = (df['short_ratio'] - df['long_ratio']).apply(lambda x: 1 if x >= 0 else 0)

    df = df[[f'one_over_short_{study_ticker}', f'short_over_long_{study_ticker}', 'dates']].loc[df['dates'] >= min_date]
    df = df[[f'one_over_short_{study_ticker}', f'short_over_long_{study_ticker}']].loc[df['dates'] <= max_date]
    return df.reset_index(drop=True)


def make_pct_gain_chart(tkr_list, period_days=250):
    # TODO: fails if it reads nothing into a fig...
    try:
        fig = plt.figure(figsize=(PLT_WIDTH, PLT_HT))
        has_plots = False

        for i, tkr in enumerate(tkr_list):
            try:
                df = pd.read_csv(os.path.join(DATA_FOLDER, tkr + '.csv')).iloc[-period_days:, ]
            except:
                msg = f'Error reading csv for {tkr} in {tkr_list} for {period_days} days, file may not exist.\n'
                log_error(msg)
                continue

            df.reset_index(inplace=True, drop=True)
            norm_vals = (df['Adj_Close'] / df['Adj_Close'].iloc[0] * 100.0).round(2)
            past_days = np.arange(start=-period_days+1, stop=1, step=1)
            pct_change = round(norm_vals.iloc[-1]-100.0, 1)
            plt.plot(past_days, norm_vals, color=TAB[i], label=tkr_list[i] + f': {pct_change}%')

            msg = f"{tkr}, {round(df['Adj_Close'].iloc[0], 2)}, {round(df['Adj_Close'].iloc[-1], 2)}, "
            msg += f"{round(pct_change, 1)}%\n"
            print(msg)
            has_plots = True

        if has_plots:
            plt.figtext(0.7, 0.05, 'Plot by @FlyingFish365 on Twitter', fontsize=7)
            plt.title(f'Percentage Chart: {period_days} trading days (From {df.Date.iloc[0]})', fontsize=8)
            plt.xlabel(f'Past {period_days} Trading Days', fontsize=8)
            plt.ylabel(f'Percent from 100% Base', fontsize=8)
            plt.xticks(np.arange(start=-period_days, stop=0, step=period_days/10, dtype=np.int32))
            plt.legend(loc='upper left', fontsize=7)
            fig.tight_layout()

            if SHOW_PLOTS:
                plt.show()

        else:
            fig = None

    except:
        fig = None
        msg = f'Error in make_pct_gain_chart for {tkr_list} for {period_days} days.\n'
        msg += traceback.format_exc() + '\n'
        # print(msg)
        log_error(msg)

    return fig


def make_long_term_ratio_chart(df_out, study_ticker, study_desc, base_ticker, base_desc):
    fig, (ax0, ax1) = plt.subplots(2, 1, gridspec_kw={'height_ratios': [PLT_HT_MAJOR, PLT_HT_MINOR]}, figsize=(PLT_WIDTH, PLT_HT))

    ax0.plot(df_out.dates, df_out.one_mo_ratio, color='tab:grey', label=f'20 Day MA')
    ax0.plot(df_out.dates, df_out.short_ratio, color='tab:blue', label=f'{TERM_SHORT} {ABBR_FOR_AVG} MA')
    ax0.plot(df_out.dates, df_out.long_ratio, color='tab:orange', label=f'{TERM_LONG} {ABBR_FOR_AVG} MA')
    ax0.text(0.75, 0.03, 'Plot by @FlyingFish365 on Twitter', fontsize=6, transform=ax0.transAxes)
    ax0.set_xlabel('Date', fontsize=8)
    ax0.set_title(f'{study_ticker} ({study_desc}) vs {base_ticker} ({base_desc}) Moving Avg Price Ratios', fontsize=8)
    ax0.legend(loc='upper left', fontsize=6)
    ax0.margins(0.02)
    ax0.set_yscale(value='log')
    ax0.set_yticks(ticks=[])
    ax0.tick_params(axis='both', labelsize=6)

    ax1.plot(df_out.dates, df_out['rsi_long'], color='tab:blue')
    ax1.hlines((30, 70), xmin=df_out.dates.iloc[0], xmax=df_out.dates.iloc[-1], colors=('tab:grey', 'tab:grey'), linestyles='dashed')
    ax1.set_ylim(0, 100)
    ax1.tick_params(axis='y', labelsize=6)
    ax1.set_xticks(ticks=[])
    ax1.set_title(f'RSI ({RSI_LONG} trading days)', fontsize=8)
    ax1.margins(0.02)

    fig.tight_layout()

    if SHOW_PLOTS:
        plt.show()

    return fig


def make_short_term_ratio_chart(df_out, study_ticker, study_desc, base_ticker, base_desc, alt_title=None):
    term_days = 250
    fig, (ax0, ax1) = plt.subplots(2, 1, gridspec_kw={'height_ratios': [PLT_HT_MAJOR, PLT_HT_MINOR]}, figsize=(PLT_WIDTH, PLT_HT))

    ax0.plot(df_out.dates.iloc[-term_days:], df_out['Adj_Close_ratio'].iloc[-250:], color='tab:green', label='Daily Price Ratio')
    ax0.plot(df_out.dates.iloc[-term_days:], df_out['one_mo_ratio'].iloc[-250:], color='tab:grey', label=f'20 Day MA')
    ax0.plot(df_out.dates.iloc[-term_days:], df_out['short_ratio'].iloc[-250:], color='tab:blue', label=f'{TERM_SHORT} {ABBR_FOR_AVG} MA')
    ax0.text(0.75, 0.03, 'Plot by @FlyingFish365 on Twitter', fontsize=6, transform=ax0.transAxes)
    ax0.set_xlabel('Date', fontsize=8)

    if alt_title is None:
        ax0.set_title(f'{study_ticker} ({study_desc}) vs {base_ticker} ({base_desc}) Moving Avg Price Ratios (1 yr)', fontsize=8)
    else:
        ax0.set_title(alt_title, fontsize=8)
    ax0.legend(loc='upper left', fontsize=6)
    ax0.margins(0.02)
    ax0.margins(0.02)
    ax0.set_yscale(value='log')
    ax0.set_yticks(ticks=[])
    ax0.tick_params(axis='both', labelsize=6)

    ax1.plot(df_out.dates.iloc[-term_days:], df_out['rsi'].iloc[-term_days:], color='tab:blue')
    ax1.hlines((30, 70), xmin=df_out.dates.iloc[-term_days], xmax=df_out.dates.iloc[-1], colors=('tab:grey', 'tab:grey'), linestyles='dashed')
    ax1.set_ylim(0, 100)
    ax1.tick_params(axis='y', labelsize=6)
    ax1.set_xticks(ticks=[])
    ax1.set_title(f'RSI ({RSI_SHORT} trading days)', fontsize=8)
    ax1.margins(0.02)

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
    plt.margins(x=0.02, y=0.02)
    fig.tight_layout()

    if SHOW_PLOTS:
        plt.show()

    # print(df_out['dates'].iloc[0], df_out['dates'].iloc[-1])
    return fig

def write_intro_sheet(workbook):
    worksheet = workbook.add_worksheet('Description')
    worksheet.set_column(0, 0, 120)
    cells = [INTRODUCTION, DISCLAIMER, CHARTING, INTERPRETATION, PCT_CHARTS, BASES, CONTACT, CHANGES, PYTHON_NOTES, LICENSE, CITATIONS]

    for i, cell in enumerate(cells):
        worksheet.write(i, 0, cell)

    return worksheet


if __name__ == "__main__":
    if UPDATE_FILES:
        update_tickers(remove=DELETE_OBS_FILES)

    if WRITE_XL:
        filename = f'ETF_Study_ver_{__VERSION__}_{DT.strftime("%Y_%m_%d_%H_%M")}.xlsx'
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        workbook = writer.book

        if MAKE_PLOT:
            if WRITE_XL:
                worksheet = write_intro_sheet(workbook)

            # pages for the list of base tickers - long term
            for base in BASE_TKRS:
                base_desc = TKR_DICT[base]
                worksheet = workbook.add_worksheet(f'Base Tkr {base} Long Term')
                row = 0
                col = 0

                if os.path.exists(os.path.join(DATA_FOLDER, base + '.csv')):
                    df_test_b = pd.read_csv(os.path.join(DATA_FOLDER, base + '.csv'))
                    if base == 'GC=F' and df_test_b is not None and os.path.exists(os.path.join(DATA_FOLDER, 'GLD' + '.csv')):
                        df_test_b2 = pd.read_csv(os.path.join(DATA_FOLDER, 'GLD' + '.csv'))
                        if df_test_b2 is not None and df_test_b2.shape[0] > df_test_b.shape[0]:
                            base = 'GLD'

                elif os.path.exists(os.path.join(DATA_FOLDER, 'GLD' + '.csv')):
                    base = 'GLD'

                else:
                    continue

                for ticker, desc in TKR_DICT.items():
                    if base == ticker or (base == 'GC=F' and ticker in ['GLD', 'IAU', 'PHYS', 'OUNZ']) or (base == 'GLD' and ticker in ['GC=F', 'IAU', 'PHYS', 'OUNZ']):
                        continue

                    print(f'Working {ticker} vs {base} Long Term')
                    df_ratio = build_ratio_data(ticker, base)

                    if df_ratio is not None:
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
            if WRITE_PAIR_PLOTS:
                row = 0
                col = 0
                worksheet = workbook.add_worksheet(f'Ticker Pairs - Long Term')
                for base, ticker in PAIR_TKRS.items():
                    base_desc = TKR_DICT[base]
                    tkr_desc = TKR_DICT[ticker]
                    print(f'Working {ticker} vs {base}')

                    df_ratio = build_ratio_data(ticker, base)

                    if df_ratio is not None:
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

                if WRITE_XL:
                    worksheet = workbook.add_worksheet(f'Base Tkr {base} Short Term')

                    row = 0
                    col = 0

                    for ticker, desc in TKR_DICT.items():
                        if base == ticker or (base == 'GC=F' and ticker in ['GLD', 'IAU', 'PHYS', 'OUNZ']):
                            continue

                        print(f'Working {ticker} vs {base} Short Term')
                        df_ratio = build_ratio_data(ticker, base)

                        if df_ratio is not None:
                            # return the chart - put in excel
                            fig = make_short_term_ratio_chart(df_ratio, study_ticker=ticker, study_desc=desc, base_ticker=base, base_desc=base_desc)

                            if WRITE_XL:
                                img_data = io.BytesIO()
                                fig.savefig(img_data, format="png")
                                img_data.seek(0)
                                worksheet.write(0, 1, '<<< Gold to Silver Ratio.\n    Important for gold and silver investors.')
                                worksheet.set_row(row=row, height=CELL_HT)
                                worksheet.set_column(first_col=col, last_col=col, width=CELL_WIDTH)
                                worksheet.insert_image(row, col, "", {'image_data': img_data})

                                plt.close(fig)

                            col += 1
                            if col >= 3:
                                col = 0
                                row += 1

            ticker_grps = [
                ['SPY', 'QQQ', 'GLD', 'GDX', 'GDXJ'],
                ['GLD', 'GDX', 'GDXJ', 'SLV', 'SIL', 'SILJ'],
                ['SPY', 'QQQ', 'RSP', 'IWM', 'IYT', 'XLF'],
                ['SPY', 'QQQ', 'EFA', 'EEM'],
                ['SPY', 'QQQ', 'LQD', 'GOVT'],
            ]

            df = pd.read_csv(os.path.join(DATA_FOLDER, 'SPY.csv'))
            df.Date = df.Date.apply(lambda x: datetime.datetime.strptime(x, '%Y-%m-%d').date())
            df0 = df.loc[df.Date >= datetime.date(year=2020, month=3, day=23), ]
            df1 = df.loc[df.Date >= datetime.date(year=2019, month=12, day=31), ]
            periods = [20, df0.shape[0], df1.shape[0], 250]
            del df

            if WRITE_XL:
                row = 0
                col = 0
                worksheet = workbook.add_worksheet(f'Other Charts')

                # worksheet.write(0, 1, 'Gold to Silver Ratio (Important for Precious Metal Markets)')
                df_ratio = build_ratio_data(study_ticker='GC=F', base_ticker='SI=F')

                if df_ratio is not None:
                    fig = make_short_term_ratio_chart(df_ratio, study_ticker='GC=F', study_desc=TKR_DICT['GC=F'],
                                                      base_ticker='SI=F', base_desc=TKR_DICT['SI=F'],
                                                      alt_title='Gold to Silver Ratio (1 yr)')

                    if fig is not None:
                        img_data = io.BytesIO()
                        fig.savefig(img_data, format="png")
                        img_data.seek(0)
                        worksheet.set_row(row=row, height=CELL_HT)
                        worksheet.set_column(first_col=col, last_col=col, width=CELL_WIDTH)
                        worksheet.insert_image(row, col, "", {'image_data': img_data, })
                        # plt.show()
                        plt.close(fig)
                        row += 1

                for tkr_grp in ticker_grps:
                    col = 0
                    has_fig = False
                    for period in periods:
                        fig = make_pct_gain_chart(tkr_list=tkr_grp, period_days=period)

                        if fig:
                            if WRITE_XL:
                                has_fig = True
                                img_data = io.BytesIO()
                                fig.savefig(img_data, format="png")
                                img_data.seek(0)
                                worksheet.set_row(row=row, height=CELL_HT)
                                worksheet.set_column(first_col=col, last_col=col, width=CELL_WIDTH)
                                worksheet.insert_image(row, col, "", {'image_data': img_data})

                            plt.close(fig)

                            col += 1
                    if has_fig:
                        row += 1

        if WRITE_XL and writer is not None:
            try:
                writer.save()
                print('!!!Saved Excel file!!!')
            except:
                print('!!!Cannot save Excel file!!!')
        else:
            print(f'!!!Cannot save Excel file, not set to save, or writer problem, WRITE_XL: {WRITE_XL}!!!')

        # allocs must balance tickers and sum to 1.0
        ref_portfolio(tkers=['SPY', 'FBND'], allocs=[0.6, 0.4])
        ref_portfolio(tkers=['SPY', 'QQQ','FBND'], allocs=[0.4, 0.2, 0.4])
        ref_portfolio(tkers=['FBND'], allocs=[1.0])
        ref_portfolio(tkers=['SPY'], allocs=[1.0])
        ref_portfolio(tkers=['SPY', 'LQD', 'GOVT'], allocs=[0.6, 0.2, 0.2])

    # first dates:
    print('SPY', pd.read_csv(os.path.join(DATA_FOLDER, 'SPY.csv')).Date.iloc[0])
    print('RSP', pd.read_csv(os.path.join(DATA_FOLDER, 'RSP.csv')).Date.iloc[0])
    print('JKE', pd.read_csv(os.path.join(DATA_FOLDER, 'JKE.csv')).Date.iloc[0])
    print('JKF', pd.read_csv(os.path.join(DATA_FOLDER, 'JKF.csv')).Date.iloc[0])
    print('IWM', pd.read_csv(os.path.join(DATA_FOLDER, 'IWM.csv')).Date.iloc[0])
    print('EFA', pd.read_csv(os.path.join(DATA_FOLDER, 'EFA.csv')).Date.iloc[0])
    print('GDX', pd.read_csv(os.path.join(DATA_FOLDER, 'EFA.csv')).Date.iloc[0])
    #
    # print('LQD', pd.read_csv(os.path.join(DATA_FOLDER, 'LQD.csv')).Date.iloc[0])
    # print('IEF', pd.read_csv(os.path.join(DATA_FOLDER, 'IEF.csv')).Date.iloc[0])

    # adaptive_portfolio()

    print(f'Done')  # , log file (if saved): {LOG_FILE_NAME}')
