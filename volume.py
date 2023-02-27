from pprint import pprint
import pdb
import os 
from kiteconnect import KiteConnect
import pandas as pd
import datetime
from datetime import datetime
import time, json, datetime, sys 
import ast
import xlwings as xw
import support_file as get



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

def get_fno_data(name, delta, interval,oi):

	token =  kite.ltp([name])[name]['instrument_token']
	to_date = datetime.datetime.now().date()
	from_date = to_date - datetime.timedelta(days=delta)
	data = kite.historical_data(instrument_token=token, from_date=from_date, to_date=to_date, interval=interval,  oi=True)
	dff = pd.DataFrame(data)	
	return dff





if __name__ == '__main__' :
	get_login_credentials()
	get_access_token()
	get_kite()


print("VOLUME---------------------------FINBABA-----------------------------VOLUME")
print("VOLUME---------------------------FINBABA-----------------------------VOLUME")
print("VOLUME---------------------------FINBABA-----------------------------VOLUME")
print("VOLUME---------------------------FINBABA-----------------------------VOLUME")
print("VOLUME-------------------------8871446294----------------------------VOLUME")
print("VOLUME---------------------------FINBABA-----------------------------VOLUME")
print("VOLUME---------------------------FINBABA-----------------------------VOLUME")
print("VOLUME---------------------------FINBABA-----------------------------VOLUME")

lst = []
lst_av = []
lst_c = []
wb = xw.Book('Operations.xlsx')
sht = wb.sheets['Volume']
name = sht.range("A1").value
name  = 'NFO:' + name
# pdb.set_trace()

while True:
	try:
		
		# time.sleep(0.08)
		name = sht.range("A1").value


		df  = get_fno_data(name=name , delta=0,interval = 'minute', oi=True)
		df = df.set_index(df['date'])
		df['S_vol'] = df['volume'].shift(1)
		df['C_vol'] = df['volume'] - df['S_vol']
		df['S_OI'] = df['oi'].shift(1)

		df['c_oi'] = df['S_OI'] - df['oi']
		df['d_close'] = df['close'].shift(1)
		df['LTP_CHG'] = df['close'] - df['d_close']

		df = df[['open' , 'high','low', 'close','volume','C_vol','LTP_CHG', 'c_oi']]

		# pdb.set_trace()
		sht.range('B1').value = df

			
	except Exception as e:
		pass


