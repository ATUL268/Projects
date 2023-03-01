from NorenRestApiPy.NorenApi import  NorenApi
import logging
import pdb,time
from pprint import pprint
import pyotp,time,os
# import zrd_login
import pandas as pd
from support_file_shoonya import ShoonyaApiPy
from datetime import datetime
import xlwings as xw


def start():
	global api ,ret, feed_opened

	key = "4ZOVKF6K647ACIR37T6LT45K4W6L57Q3"
	totp = pyotp.TOTP(key).now()
	totp = str(totp)
	user    = "FA98926"
	pwd     = "Share@0077"
	vc      = "FA98926_U"
	app_key = "323e466b6be0274ec1853ca3e0581e48"
	imei    = "abc1234" 

	api = ShoonyaApiPy()

	#make the api call
	ret = api.login(userid=user, password=pwd, twoFA=totp, vendor_code=vc, api_secret=app_key, imei=imei)
	feed_opened = False
	api.start_websocket( order_update_callback=event_handler_order_update,subscribe_callback=event_handler_feed_update, socket_open_callback=open_callback)

	while(feed_opened==False):
		pass


def event_handler_feed_update(tick_data):
	if 'lp' in tick_data and 'tk' in tick_data :
		timset = datetime.fromtimestamp(int(tick_data['ft'])).isoformat()
		rest_api[tick_data['tk']] = {'LTP' : float(tick_data['lp']) , 'tt' : timset}
		# print(f"feed update {tick_data}")

def event_handler_order_update(tick_data):
	print(f"Order update {tick_data}")

def open_callback():
	global feed_opened
	feed_opened = True

def start_excel():
	global wb,dt,ex,ob,rest_api,get_watchlist,main_dict,watchlist
	print("Excel Starting...")
	if not os.path.exists("ticks.xlsx"):
		try:
			wb = xw.Book()
			wb.save('ticks.xlsx')
			wb.close()
		except Exception as e: 
			print(f"Error : {e}") 
			sys.exit()
	wb = xw.Book('ticks.xlsx')
	for i in ["Data", "Exchange_NFO", "orderbook"]:
		try:
			wb.sheets(i)
		except:
			wb.sheets.add(i)
	dt = wb.sheets("Data")
	ex = wb.sheets("Exchange_NFO") 
	ob = wb.sheets("OrderBook")
	ex.range("a:j").value = ob.range("a:h").value = dt.range("p:q").value = None
	dt.range('B1').value  = 'Symbols'

	while True:
		try:
			df  =  pd.read_csv('https://api.shoonya.com/NFO_symbols.txt.zip')
			df['Trading_Symbol'] = df['Exchange'] + ':' + df["TradingSymbol"]
			ndf = df[['Trading_Symbol','LotSize','Expiry', 'Instrument','OptionType','StrikePrice','TickSize' , "TradingSymbol"]]
			ex.range('A1').value = ndf

			break
		except Exception as e:
			print(e)

def get_watch():
	global watchlist , rest_api ,get_watchlist
	watchlist = dt.range('B2').expand('down').value

	rest_api = {}
	main_dict = {}
	get_watchlist = []
	for name in watchlist:

		val = api.searchscrip(exchange=name[:3], searchtext=name[4:])
		token = val['values'][0]['token']
		new_name  = val['values'][0]['exch'] + '|' +  val['values'][0]['token']
		get_watchlist.append(new_name)
		main_dict[name] = token

	# pdb.set_trace()

	api.subscribe(get_watchlist)
	print(get_watchlist)
	time.sleep(1)


def Live_Feed():
	global pdf

	print('Live_Data is running..........Sucessful')
	subs_lst = []
	while True:

		time.sleep(0.02)
		try:
			
			pdf = pd.DataFrame(rest_api).T
			pdf = pdf.drop(['tt'], axis=1)
			dt.range('E1').value  = pdf
			# pdb.set_trace()
		
		except Exception as e:
			continue




def user_input():

	print('Please Select Your Symbols in Work Tab of Ticks Excel......\n' 'Put DONE in A1 Cell \n' 'Select Your Symbols In B cell.....')
	print('----------------------------------------Finbaba--------------------------------------')
	print('----------------------------------------Finbaba--------------------------------------')
	print('----------------------------------------Finbaba--------------------------------------')
	print('--------------------------------------8871446294-------------------------------------')
	print('----------------------------------------Finbaba--------------------------------------')
	print('----------------------------------------Finbaba--------------------------------------')
	print('----------------------------------------Finbaba--------------------------------------')
	print('Please Select Your Symbols in Work Tab of Ticks Excel......\n' 'Put DONE in A1 Cell \n' 'Select Your Symbols In B cell.....')

	while True:
		try:
			time.sleep(0.10)
			wb = xw.Book('ticks.xlsx')
			dt = wb.sheets("Data")
			aa = dt.range("A1").value
			if aa.upper() == "DONE":
				print(aa)
				break
		except Exception as e:
			print(e)

#---------------------------------Start--------------------------------





# read_symbols_NSE  =  pd.read_csv('https://api.shoonya.com/NSE_symbols.txt.zip')
# read_symbols_NFO  =  pd.read_csv('https://api.shoonya.com/NFO_symbols.txt.zip')
# pdb.set_trace()


if __name__ == '__main__' :
	start()
	start_excel()
	user_input()
	get_watch()
	Live_Feed()



