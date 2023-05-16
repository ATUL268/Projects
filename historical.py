from pprint import pprint
import pdb 
import pandas as pd
import time
import zrd_login
import datetime
import support_file_2023 as sf
import xlwings as xw
kite = zrd_login.kite



def LTP(name):
	last_price = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
	return last_price

status = {}
final = {}
Scrip = 'NIFTY BANK'
name1 = "BANKNIFTY"
no_of_days = 2
time_frame = '5minute'
step_value = 100
multiplier = 0
expiry = "23511" #Expiry 
lot_Size = 25

status = { 'UA_close' : None , 'ATM' : None ,  'close_CE': None , 'volume_CE': None , 'oi_CE': None , 'close_PE': None , 'volume_PE': None , 'oi_PE' : None}

ltp_bn = LTP(Scrip)
idf = sf.get_data(name = 'NIFTY BANK', segment = 'NSE:', delta = no_of_days, interval= time_frame, continuous= False, oi=True)
idf = idf.set_index(idf['date'])
# idf = idf[["close" ]]
time.sleep(2)

for index, ohlc in idf.iterrows():
	close = idf.loc[index]['close']
	atm = round(close/step_value)* step_value 
	atm_name_CE = name1 + expiry + str(atm) +'CE'
	atm_name_PE = name1 + expiry + str(atm) +'PE'


	opdf_CE = sf.get_data(name = atm_name_CE, segment = 'NFO:', delta = no_of_days, interval= time_frame, continuous= False, oi=True)
	opdf_CE = opdf_CE.set_index(opdf_CE['date'])
	opdf_CE['previous_OI_CE'] = opdf_CE['oi'].shift(1)

	close_CE = opdf_CE.loc[index]['close']
	volume_CE = opdf_CE.loc[index]['volume']
	oi_CE = opdf_CE.loc[index]['oi']


	opdf_PE = sf.get_data(name = atm_name_PE, segment = 'NFO:', delta = no_of_days, interval= time_frame, continuous= False, oi=True)
	opdf_PE = opdf_PE.set_index(opdf_PE['date'])
	pdb.set_trace()

	close_PE = opdf_PE.loc[index]['close']
	volume_PE = opdf_PE.loc[index]['volume']
	oi_PE = opdf_PE.loc[index]['oi']

	status['UA_close'] = close
	status['ATM'] = atm

	status['close_CE'] = close_CE
	status['volume_CE'] = volume_CE
	status['oi_CE'] = oi_CE

	status['close_PE'] = close_PE
	status['volume_PE'] = volume_PE
	status['oi_PE'] = oi_PE

	final[index] = status
	df = pd.DataFrame(final).T
	status = { 'UA_close' : None , 'ATM' : None ,  'close_CE': None , 'volume_CE': None , 'oi_CE': None , 'close_PE': None , 'volume_PE': None , 'oi_PE' : None}
	print(index)


df['previous_OI_CE'] = df['oi_CE'].shift(1)
df['previous_OI_PE'] = df['oi_PE'].shift(1)
df['COI_CE'] = df['oi_CE'] - df['previous_OI_CE']
df['COI_IN_LOT_CE'] = df['COI_CE'] / lot_Size
df['COI_PE'] = df['oi_PE'] - df['previous_OI_PE']
df['COI_IN_LOT_PE'] = df['COI_PE'] / lot_Size


	
df['COI_Percentage_CE'] = abs(round(((df['previous_OI_CE'] - df['oi_CE'])/df['oi_CE'])*100,0))
df['COI_Percentage_PE'] = abs(round(((df['previous_OI_PE'] - df['oi_PE'])/df['oi_PE'])*100,0))

df['PUT/CALL'] = abs(round(df['COI_PE'] / df['COI_CE'],2))
df['CALL/PUT'] = abs(round(df['COI_CE'] / df['COI_PE'],2))











