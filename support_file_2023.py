from pprint import pprint
import zrd_login
import datetime
import pandas as pd
from pprint import pprint
import pdb


step_values = {'AARTIIND': 20, 'ACC': 20, 'ADANIENT': 20, 'ADANIPORTS': 10, 'ALKEM': 20, 'AMARAJABAT': 10, 'AMBUJACEM': 5, 'APLLTD': 10, 'APOLLOHOSP': 50, 'APOLLOTYRE': 5, 'ASHOKLEY': 2.5, 'ASIANPAINT': 20, 'AUBANK': 20, 'AUROPHARMA': 10, 'AXISBANK': 10, 'BAJAJ-AUTO': 50, 'BAJAJFINSV': 100, 'BAJFINANCE': 100, 'BALKRISIND': 20, 'BANDHANBNK': 5, 'BANKBARODA': 2.5, 'BATAINDIA': 20, 'BEL': 2.5, 'BERGEPAINT': 10, 'BHARATFORG': 10, 'BHARTIARTL': 10, 'BHEL': 1, 'BIOCON': 5, 'BOSCHLTD': 250, 'BPCL': 10, 'BRITANNIA': 20, 'CADILAHC': 5, 'CANBK': 5, 'CHOLAFIN': 10, 'CIPLA': 10, 'COALINDIA': 2.5, 'COFORGE': 50, 'COLPAL': 20, 'CONCOR': 10, 'CUB': 2.5, 'CUMMINSIND': 20, 'DABUR': 5, 'DEEPAKNTR': 20, 'DIVISLAB': 50, 'DLF': 5, 'DRREDDY': 50, 'EICHERMOT': 50, 'ESCORTS': 20, 'EXIDEIND': 2.5, 'FEDERALBNK': 1, 'GAIL': 2.5, 'GLENMARK': 5, 'GMRINFRA': 1, 'GODREJCP': 10, 'GODREJPROP': 20, 'GRANULES': 5, 'GRASIM': 20, 'GUJGASLTD': 10, 'HAVELLS': 20, 'HCLTECH': 10, 'HDFC': 50, 'HDFCAMC': 50, 'HDFCBANK': 20, 'HDFCLIFE': 10, 'HEROMOTOCO': 50, 'HINDALCO': 5, 'HINDPETRO': 5, 'HINDUNILVR': 20, 'IBULHSGFIN': 5, 'ICICIBANK': 10, 'ICICIGI': 20, 'ICICIPRULI': 5, 'IDEA': 1, 'IDFCFIRSTB': 1, 'IGL': 10,
               'INDIGO': 20, 'INDUSINDBK': 20, 'INDUSTOWER': 5, 'INFY': 20, 'IOC': 1, 'IRCTC': 20, 'ITC': 2.5, 'JINDALSTEL': 10, 'JSWSTEEL': 5, 'JUBLFOOD': 50, 'KOTAKBANK': 20, 'L&TFH': 2.5, 'LALPATHLAB': 50, 'LICHSGFIN': 10, 'LT': 20, 'LTI': 100, 'LTTS': 50, 'LUPIN': 20, 'M&M': 10, 'M&MFIN': 5, 'MANAPPURAM': 2.5, 'MARICO': 5, 'MARUTI': 100, 'MCDOWELL-N': 5, 'MFSL': 20, 'MGL': 20, 'MINDTREE': 20, 'MOTHERSUMI': 5, 'MPHASIS': 20, 'MRF': 500, 'MUTHOOTFIN': 20, 'NAM-INDIA': 5, 'NATIONALUM': 1, 'NAUKRI': 100, 'NAVINFLUOR': 50, 'NESTLEIND': 100, 'NMDC': 2.5, 'NTPC': 1, 'ONGC': 2.5, 'PAGEIND': 500, 'PEL': 50, 'PETRONET': 5, 'PFC': 2.5, 'PFIZER': 50, 'PIDILITIND': 20, 'PIIND': 20, 'PNB': 1, 'POWERGRID': 2.5, 'PVR': 20, 'RAMCOCEM': 10, 'RBLBANK': 5, 'RECLTD': 2.5, 'RELIANCE': 20, 'SAIL': 2.5, 'SBILIFE': 10, 'SBIN': 5, 'SHREECEM': 500, 'SIEMENS': 20, 'SRF': 100, 'SRTRANSFIN': 20, 'SUNPHARMA': 10, 'SUNTV': 10, 'TATACHEM': 10, 'TATACONSUM': 10, 'TATAMOTORS': 5, 'TATAPOWER': 2.5, 'TATASTEEL': 10, 'TCS': 50, 'TECHM': 20, 'TITAN': 20, 'TORNTPHARM': 20, 'TORNTPOWER': 5, 'TRENT': 20, 'TVSMOTOR': 10, 'UBL': 20, 'ULTRACEMCO': 100, 'UPL': 10, 'VEDL': 5, 'VOLTAS': 20, 'WIPRO': 5, 'ZEEL': 5}
watchlist = ['ADANIPORTS', 'ASIANPAINT', 'INFRATEL', 'AXISBANK', 'BAJAJ-AUTO', 'BAJFINANCE', 'BAJAJFINSV', 'BPCL', 'BHARTIARTL', 'BRITANNIA', 'CIPLA', 'COALINDIA', 'DIVISLAB', 'DRREDDY', 'EICHERMOT', 'GAIL', 'GRASIM', 'HCLTECH', 'HDFCBANK', 'HDFCLIFE', 'HEROMOTOCO', 'HINDALCO', 'HINDUNILVR', 'HDFC', 'ICICIBANK', 'ITC', 'IOC', 'INDUSINDBK', 'INFY', 'JSWSTEEL', 'KOTAKBANK', 'LT', 'M&M', 'MARUTI', 'NTPC', 'NESTLEIND', 'ONGC', 'POWERGRID', 'RELIANCE', 'SBILIFE', 'SHREECEM', 'SBIN', 'SUNPHARMA', 'TCS', 'TATAMOTORS', 'TATASTEEL', 'TECHM', 'TITAN', 'UPL', 'ULTRACEMCO', 'WIPRO']

kite = zrd_login.kite


def get_good_values(name):

    zrd_name = 'NSE:' + name
    data = kite.quote([zrd_name])

    ltp = data[zrd_name]['last_price']
    openx = data[zrd_name]['ohlc']['open']
    high = data[zrd_name]['ohlc']['high']
    low = data[zrd_name]['ohlc']['low']
    close = data[zrd_name]['ohlc']['close']
    volume = data[zrd_name]['volume']

    return ltp, openx, high, low, close, volume


def option_name_finder(ltp, step_value, multiplier, name, exipry, ce_pe):
    step_value = step_values[name]
    atm_strike = round(ltp/step_value)*step_value + multiplier*step_value
    option_name = name + exipry + str(atm_strike) + 'CE'
    return option_name


def get_data(name, segment, delta, interval, continuous, oi):

    token = kite.ltp([segment + name])[segment + name]['instrument_token']
    to_date = datetime.datetime.now().date()
    from_date = to_date - datetime.timedelta(days=delta)

    data = kite.historical_data(instrument_token=token, from_date=from_date, to_date=to_date, interval=interval, continuous=False, oi=True)
    df = pd.DataFrame(data)
    # df = df.set_index(df['date'])
    return df