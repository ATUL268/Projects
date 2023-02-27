import os 
from kiteconnect import KiteConnect
# from kiteconnect import KiteTicker
import time, json, datetime, sys 
import xlwings as xw 
import pandas as pd
import pdb
import ast


def LTP(name):

	
	data = kite.quote([name])
	ltp = data[name]['last_price']
	return ltp

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
				curr_time = datetime.datetime.now().time()
				ltp = kite.ltp(watchlist)[name]['last_price']
				final[name[4:]].append({'time' : ctime, 'close' : ltp})
				df = pd.DataFrame(final[name[4:]])
				df = df.set_index(df['time'])
				df = df[['close']]
				ohlc = df.resample('1Min').ohlc()
				completed_candle = pd.Series(datetime.datetime.now()).dt.floor('1min')[0]- datetime.timedelta(minutes=1)
				completed_candle = completed_candle.strftime("%Y-%m-%d %H:%M:%S")
				wk.range("B12").value = kite.ltp(watchlist)['NSE:NIFTY 50']['last_price']

				if name==watchlist[0]:
					wk.range("C2").value = ohlc.loc[:completed_candle]

				if name==watchlist[1] and wk.range("B2").value == True:
					wk.range("I2").value = ohlc.loc[:completed_candle]

				if name==watchlist[2] and wk.range("B3").value == True:
					wk.range("O2").value = ohlc.loc[:completed_candle]

				if name==watchlist[3] and wk.range("B4").value == True:
					wk.range("U2").value = ohlc.loc[:completed_candle]

				if name==watchlist[4] and wk.range("B5").value == True:
					wk.range("AA2").value = ohlc.loc[:completed_candle]

				if name==watchlist[5] and wk.range("B6").value == True:
					wk.range("AG2").value = ohlc.loc[:completed_candle]

				if name==watchlist[6] and wk.range("B7").value == True:
					wk.range("AM2").value = ohlc.loc[:completed_candle]

				if name==watchlist[7] and wk.range("B8").value == True:
					wk.range("AS2").value = ohlc.loc[:completed_candle]	

				if name==watchlist[8] and wk.range("B9").value == True:
					wk.range("AY2").value = ohlc.loc[:completed_candle]	
			
			
		except Exception as e:
			pass

if __name__ == '__main__' :
	operations()
	get_login_credentials()
	get_access_token()
	get_kite()
	start_excel()
	user_input()
	make_lists()
	ohlc_Forming()

		





# pdb.set_trace()