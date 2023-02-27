# import api_login as login
# kite = login.kite
from pprint import pprint
import pdb
import pandas as pd
import support_file as get
import datetime
from datetime import datetime
import os 
from kiteconnect import KiteConnect
# from kiteconnect import KiteTicker
import time, json, datetime, sys 
import xlwings as xw 
import ast


# 2021 ---------------------------------------------------------2021---------------------------------------

def LTP(name):

	zrd_name = 'NSE:'+ name
	data = kite.quote([zrd_name])
	ltp = data[zrd_name]['last_price']
	return ltp

def OPENN(name):
	openn = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['open']
	return openn

def HIGH(name):
	high = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['high']
	return high

def LOW(name):
	low = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['low']
	return low

def CLOSE(name):
	close = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['close']
	return close

def BID(name):
	bid_price = kite.quote(['NSE:'+ name])['NSE:'+ name]['depth']['buy'][0]['price']
	return bid_price

def ASK(name):
	Ask_price = kite.quote(['NSE:'+ name])['NSE:'+ name]['depth']['sell'][0]['price']
	return Ask_price

def VOLUME(name):
	volume = kite.quote(['NSE:'+ name])['NSE:'+ name]['volume']
	return volume

def l_ohlc_v(name):
	zrd_name = 'NSE:'+ name
	data = kite.quote([zrd_name])
	ltp = data[zrd_name]['last_price']
	openn = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['open']
	high = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['high']
	low = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['low']
	close = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['close']
	volume = kite.quote(['NSE:'+ name])['NSE:'+ name]['volume']
	return ltp,openn,high,low,close,volume

def ob():
	margins = kite.margins()
	ob = margins['equity']['available']['opening_balance']
	return ob

def ohlc(name):
	
	openn = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['open']
	high = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['high']
	low = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['low']
	close = kite.ohlc(['NSE:' + name])['NSE:' + name]['ohlc']['close']
	
	return openn,high,low,close,

def BID_1(name):
	data = kite.quote('NSE:' + name)
	bid_price = data['NSE:'+name]['depth']['buy'][0]['price']
	return bid_price

def lb():
	margins = kite.margins()
	lb = margins['equity']['available']['live_balance']
	return lb

def pnl(name):
	sym = name
	pos = kite.positions()
	net_pos = pos['net']
	len_net_pos = len(net_pos)

	for x in range(len_net_pos):
		if net_pos[x]['tradingsymbol']==sym:
			pnl = (round(pos['net'][x]['pnl'],2))
			return pnl

def net_pnl():
	traded_stocks_pnl = []
	pos = kite.positions()
	net_pos = pos['net']
	len_net_pos = len(net_pos)

	for x in range(len_net_pos):
		a = round(net_pos[x]['pnl'],2)
		traded_stocks_pnl.append(a)
	value = (sum(traded_stocks_pnl))
	return value

def day_pnl():
	traded_stocks_pnl = []
	pos = kite.positions()
	net_pos = pos['day']
	len_net_pos = len(net_pos)

	for x in range(len_net_pos):
		a = round(net_pos[x]['pnl'],2)
		traded_stocks_pnl.append(a)
	value = (sum(traded_stocks_pnl))
	return value

def LTP_NFO(name):

	zrd_name = 'NFO:'+ name
	data = kite.quote([zrd_name])
	ltp = data[zrd_name]['last_price']
	return ltp

def BID_NFO(name):
	bid_price = kite.quote(['NFO:'+ name])['NFO:'+ name]['depth']['buy'][0]['price']
	return bid_price

def False_expiry():
	Expiry = datetime.datetime.now().strftime("%y") + datetime.datetime.now().strftime("%b").upper()
	return Expiry

def option_name_atm(name,ce_pe):
	ltp = LTP(name)
	sv = step_value[name]
	atm_strike = round(ltp/sv)*sv
	option_name = name + expiry()+ str(atm_strike)+ce_pe
	return option_name

def option_name(name,ce_pe,multiplier):
	ltp = LTP(name)
	sv = step_value[name]
	atm_strike = round(ltp/sv)*sv+sv*multiplier
	option_name = name + expiry()+ str(atm_strike)+ce_pe
	return option_name

def get_fno_data(name, delta, interval,oi):

	token =  kite.ohlc([name])[name]['instrument_token']
	to_date = datetime.datetime.now().date()
	from_date = to_date - datetime.timedelta(days=delta)
	data = kite.historical_data(instrument_token=token, from_date=from_date, to_date=to_date, interval=interval,  oi=True)
	df = pd.DataFrame(data)	
	return df

def get_equity_data(name, segment, delta, interval, continuous, oi):

	token = kite.ltp([segment + name])[segment + name]['instrument_token']
	to_date = datetime.datetime.now().date()
	from_date = to_date - datetime.timedelta(days=delta)

	data = kite.historical_data(instrument_token=token, from_date=from_date, to_date=to_date, interval=interval, continuous=False, oi=False)
	df = pd.DataFrame(data)
	# df = df.set_index(df['date'])
	return df

def read_data(name):

	this_add = '21JUNFUT'
	next_add = '21JULFUT'

	this_name = name + this_add + '.csv'
	next_name =  name + next_add + '.csv'

	this = pd.read_csv('this' + '\\'+ this_name)
	nextt = pd.read_csv('next' + '\\'+ next_name)

	this =  this.set_index(this['date'])
	nextt = nextt.set_index(nextt['date'])

	return this,nextt

# 2022 ---------------------------------------------------------2022---------------------------------------

def get_good_values(name):

	zrd_name = 'NSE:' + name
	data = kite.quote([zrd_name])

	ltp = data[zrd_name]['last_price']
	openx = data[zrd_name]['ohlc']['open']
	high = data[zrd_name]['ohlc']['high']
	low = data[zrd_name]['ohlc']['low']
	close = data[zrd_name]['ohlc']['close']
	# volume = data[zrd_name]['volume']

	return ltp, openx, high, low, close

def wd():

	trade_day = {'Monday' : {'stop_loss': 0.06, 'target': 0.84},'Tuesday' : {'stop_loss': 0.60, 'target': 0.70},'Wednessday' : {'stop_loss': 0.20, 'target': 0.95},'Thrusday' : {'stop_loss': 0.14, 'target': 0.84},'Friday' : {'stop_loss': 0.06, 'target': 0.84},}
	week_day = datetime.today().strftime('%A')
	sl = trade_day[week_day]['stop_loss']
	tgt = trade_day[week_day]['target']
	print(week_day)
	return sl ,tgt

def convert(date_time):
    format = "%d-%b-%Y"# The format
    datetime_str = datetime.datetime.strptime(date_time, format)
 
    return datetime_str

def get_data_range(name,segment,delta,delta1,interval,continuous,oi):
	token = kite.ltp([segment + name])[segment + name]['instrument_token']
	to_date = datetime.datetime.now().date()
	to_date1 = to_date - datetime.timedelta(days = delta1)
	from_date = to_date - datetime.timedelta(days = delta)
	hd = kite.historical_data(instrument_token = token, from_date = from_date, to_date = to_date1, interval = interval, continuous = False, oi = False)
	hd = pd.DataFrame(hd)
	return h

def current_expiry():
	

	expiry_status = {}
	year = datetime.datetime.now().year
	month =datetime.datetime.now().month
	day = datetime.datetime.now().day

	test_date = datetime.datetime(year, month, day)
	nxt_mnth = test_date.replace(day=28) + datetime.timedelta(days=4)
	res = nxt_mnth - datetime.timedelta(days=nxt_mnth.day)
	last_day_month = (res.day)

	index_list_from_nse = (indices)

	data_from_nse = expiry_list(index_list_from_nse[2])

	for name in data_from_nse:

		date_time = name
		soso = get.convert(date_time)
		soso = str(soso)
		month_cut_soso =soso[6:7]
		year_cut_soso = soso[2:4]

		new_name = name[0:2]
		expiry_date = int(new_name)

		if (last_day_month - expiry_date) > 7:
			# print('weekly_expiry')

			# year = str(year)
			# year = year[2:4]
			# month = str(month)
			expiry_weekly = ( year_cut_soso + month_cut_soso + name[0:2])
			expiry_status[name] = expiry_weekly
			year = datetime.datetime.now().year
			# print(expiry_status)



		if (last_day_month - expiry_date) < 7:
			# print('Monthly_expiry')

			# year = str(year)
			# year = year[2:4]
			month_first_3 = str(datetime.datetime.now().strftime("%b").upper())
			
			expiry_monthly = ( year_cut_soso + month_first_3)
			expiry_status[name] = expiry_monthly

			year = datetime.datetime.now().year


	current_expiry = data_from_nse[0]


	return expiry_status[current_expiry]

def next_expiry():
	

	expiry_status = {}
	year = datetime.datetime.now().year
	month =datetime.datetime.now().month
	day = datetime.datetime.now().day

	test_date = datetime.datetime(year, month, day)
	nxt_mnth = test_date.replace(day=28) + datetime.timedelta(days=4)
	res = nxt_mnth - datetime.timedelta(days=nxt_mnth.day)
	last_day_month = (res.day)

	index_list_from_nse = (indices)

	data_from_nse = expiry_list(index_list_from_nse[2])

	for name in data_from_nse:

		date_time = name
		soso = get.convert(date_time)
		soso = str(soso)
		month_cut_soso =soso[6:7]
		year_cut_soso = soso[2:4]

		new_name = name[0:2]
		expiry_date = int(new_name)

		if (last_day_month - expiry_date) > 7:
			# print('weekly_expiry')

			# year = str(year)
			# year = year[2:4]
			# month = str(month)
			expiry_weekly = ( year_cut_soso + month_cut_soso + name[0:2])
			expiry_status[name] = expiry_weekly
			year = datetime.datetime.now().year
			# print(expiry_status)



		if (last_day_month - expiry_date) < 7:
			# print('Monthly_expiry')

			# year = str(year)
			# year = year[2:4]
			month_first_3 = str(datetime.datetime.now().strftime("%b").upper())
			
			expiry_monthly = ( year_cut_soso + month_first_3)
			expiry_status[name] = expiry_monthly

			year = datetime.datetime.now().year


	current_expiry = data_from_nse[1]


	return expiry_status[current_expiry]

# 2023 ---------------------------------------------------------2023---------------------------------------

def convert_ohlc(ticks: list[dict] ,name , candle,ltp,watchlist):
	ctime = datetime.datetime.now()
	ltp = kite.ltp(watchlist)[name]['last_price']
	ticks.append({'time' : ctime, 'close' : ltp})
	df = pd.DataFrame(ticks)
	df = df.set_index(df['time'])
	df = df[['close']]
	ohlc = df.resample(candle).ohlc()
	return ohlc

def get_login_credentials():
	global login_credential

	def login_credentials():
		print("---- Enter you Zerodha Login Credentials ----") 
		login_credential = {"api_key": str(input("Enter API Key :")), 
							"api_secret": str(input("Enter API Secret :")) 
							} 

		if input("Press Y to save login credential and any key to bypass : ").upper() == "Y":
			with open(f"Login credentials.txt", "w") as f: 
				json.dump(login_credential, f) 
			print("Data Saved...") 
		else: 
			print("Data Save canceled!!!!!") 




	while True:
		try:
			with open(f"Login credentials.txt", "r") as f: 
				login_credential = json.load(f) 
			break 

		except: 
				login_credentials() 
	return login_credential 

def get_access_token(): 
	global access_token 




	def login(): 
		global login_credential 
		print("Trying Log In...") 
		kite = KiteConnect(api_key = login_credential['api_key']) 
		print("Login url : ", kite.login_url()) 
		request_tkn = input("Login and enter your request token here : ") 
		try: 
			access_token = kite.generate_session(request_token=request_tkn, api_secret=login_credential["api_secret"])['access_token'] 
			os.makedirs(f"AccessToken", exist_ok=True)
			
			with open(f"AccessToken/{datetime.datetime.now().date()}.json", "w") as f: 
				json.dump(access_token, f) 
				# pdb.set_trace() 
			print("Login successful...") 
		except Exception as e: 
			print(f"Login Failed {{{e}}}") 


	print("Already Logged in for today.......")
	while True:
		if os.path.exists(f"AccessToken/{datetime.datetime.now().date()}.json"):
			with open(f"AccessToken/{datetime.datetime.now().date()}.json", "r") as f: 
				access_token = json.load(f)
			break
		else:
			login()
	return access_token

def get_kite(): 
	global kite, login_credential, access_token 

	try: 
		kite = KiteConnect(api_key = login_credential["api_key"]) 
		kite. set_access_token(access_token) 
	except Exception as e: 
		print(f"Error : {e}") 
		os.remove(f"AccessToken/{datetime.datetime.now().date()}.json") if os.path.exists(f"AccessToken/{datetime.datetime.now().date()}.json") else None 
		sys.exit()

def get_live_data(instruments):
	global kite, live_data
	try:
		live_data
	except:
		live_data = {}
	try:
		live_data = kite.quote(instruments)
	except Exception as e:
		print(f"Get live data Failed {{{e}}}") 
		pass
	return live_data

def get_live():
	subs_lst = []
	while True:
		try:
			time.sleep(0.25)
			get_live_data(subs_lst)
			symbols = dt.range(f"b{2}:b{500}").value
			

			for i in subs_lst:
				if i not in symbols:
					subs_lst.remove(i)
					try:
						del live_data[i]
					except Exception as e:
						pass
			main_list = []
			
			for i in symbols:
				lst = [None] 
				if i:
					if i not in subs_lst:
						subs_lst.append(i)
					if i in subs_lst:
						try:
							lst = [live_data[i]["last_price"]]
							# pdb.set_trace()

						except Exception as e:
							pass
				main_list.append(lst)
				
				# pdb.set_trace()
			dt.range("c2").value = main_list
			# pdb.set_trace()
		except Exception as e:
			print(e)
			pass

def start_excel():
	global kite, live_data 
	print("Opening Your Excel.......")
	if not os.path.exists("Master_excel.xlsx"):
		try:
			wb = xw.Book()
			wb.save('Master_excel.xlsx')
			wb.close()
		except Exception as e: 
			print(f"Error : {e}") 
			sys.exit()
	wb = xw.Book('Master_excel.xlsx')
	for i in ["Data", "Exchange" , "Selected"]:
		try:
			wb.sheets(i)
		except:
			wb.sheets.add(i)
	dt = wb.sheets("Data")
	ex = wb.sheets("Exchange") 
	se = wb.sheets("Selected") 
	ex.range("a:j").value = dt.range("p:q").value = se.range("a:n").value  = None
	dt.range(f"a1:q1").value = ["Sr/No", "Symbol", "LTP"]


	subs_lst = []
	while True:
		try:
			df = pd.DataFrame(kite.instruments())
			df = df.drop(["instrument_token","exchange_token","last_price","tick_size"],axis = 1)
			df["watchlist_symbol"] = df["exchange"] + ":" + df["tradingsymbol"] 
			df.columns = df.columns.str.replace("_", " ")
			df.columns = df.columns.str.title()
			ex.range("a1").value = df
			ndf = df[df['Exchange'].str.contains("NFO")]
			ndf = ndf[ndf['Segment'].str.contains("NFO-OPT")]
			ndf = ndf[ndf['Name'].str.match("NIFTY")]
			ndf = ndf[ndf['Expiry'].astype(str).str.contains('2023')]
			ndf = ndf.set_index(ndf['Strike'])
			ndf = ndf[ndf.Strike.between(lower_range,upper_range)]
			ndf = ndf[['Watchlist Symbol' , 'Instrument Type' , 'Expiry' , 'Tradingsymbol']]
			se.range("a1").value = ndf
			break

		except Exception as e:
			time.sleep(1)
			# pdb.set_trace()
	
def operations():
	global lower_range,upper_range ,wb ,wk,tick_list

	try:
		wb = xw.Book('Operations.xlsx')
		wk = wb.sheets['Work']
		wk.range("A1").value = 'SCRIPS'
	except Exception as e:
		pass
	lower_range =  wk.range("B14").value
	upper_range = wk.range("B15").value
	wk.range("C2:BG500").value  = None
	tick_list = []

def user_input():

	print('Please Select Your Symbols in Work Tab of Operations Excel......\n' 'Put DONE in B1 Cell')
	while True:
		try:
			time.sleep(0.10)
			aa = wk.range("B1").value
			if aa.upper() == "DONE":
				break
		except Exception as e:
			pass

def make_lists():
	global symbols ,watchlist,final ,name 
	symbols = wk.range(f"A{2}:A{9}").value
	watchlist = ['NSE:NIFTY 50'] + symbols

	final = {}

	for name in watchlist:
		try:
			final[name[4:]] = tick_list.copy()

		except Exception as e:
			pass

def ohlc_Forming():
	global ohlc,ctime
	
	print('OHLC is forming...........see Excel Operations ')
	while True:
		try:
			
			for name in watchlist:
				ctime = datetime.datetime.now()
				ltp = kite.ltp(watchlist)[name]['last_price']
				final[name[4:]].append({'time' : ctime, 'close' : ltp})
				df = pd.DataFrame(final[name[4:]])
				df = df.set_index(df['time'])
				df = df[['close']]
				ohlc = df.resample('1Min').ohlc()
				wk.range("B12").value = kite.ltp(watchlist)['NSE:NIFTY 50']['last_price']

				if name==watchlist[0]:
					wk.range("C2").value = ohlc

				if name==watchlist[1]:
					wk.range("I2").value = ohlc

				if name==watchlist[2]:
					wk.range("O2").value = ohlc

				if name==watchlist[3]:
					wk.range("U2").value = ohlc

				if name==watchlist[4]:
					wk.range("AB2").value = ohlc

				if name==watchlist[5]:
					wk.range("AH2").value = ohlc

				if name==watchlist[6]:
					wk.range("AN2").value = ohlc

				if name==watchlist[7]:
					wk.range("AR2").value = ohlc	
			
		except Exception as e:
			pass