from pprint import pprint
import pdb 
import pandas as pd
import time
import zrd_login
import datetime
import support_file_2023 as sf
import xlwings as xw

def ATM_finder_CE_BN(name,name1):
    expiry = L_expiry
    ltp = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
    step_value = 100
    multiplier = 0
    CE_atm_strike = round(ltp/step_value)* step_value + multiplier*step_value
    CE_ATM = (name1 + expiry + str(CE_atm_strike ) + 'CE' )
    return (CE_ATM)

def ATM_finder_PE_BN(name,name1):
    expiry = L_expiry
    ltp = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
    step_value = 100
    multiplier = 0
    PE_atm_strike = round(ltp/step_value)* step_value + multiplier*step_value
    PE_ATM = (name1 + expiry + str(PE_atm_strike ) + 'PE' )
    return (PE_ATM)

def ATM_finder_CE_N(name,name1):
    expiry = L_expiry
    ltp = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
    step_value = 50
    multiplier = 0
    CE_atm_strike = round(ltp/step_value)* step_value + multiplier*step_value
    CE_ATM = (name1 + expiry + str(CE_atm_strike ) + 'CE' )
    return (CE_ATM)

def ATM_finder_PE_N(name,name1):
    expiry = L_expiry
    ltp = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
    step_value = 50
    multiplier = 0
    PE_atm_strike = round(ltp/step_value)* step_value + multiplier*step_value
    PE_ATM = (name1 + expiry + str(PE_atm_strike ) + 'PE' )
    return (PE_ATM)

def LTP(name):
    last_price = kite.ltp(['NSE:'+ name])['NSE:'+ name]['last_price']
    return last_price


def LTP_NFO(name):
    last_price = kite.ltp(['NFO:'+ name])['NFO:'+ name]['last_price']
    return last_price


def init():
    global kite,wb,L_expiry,sht_BN,sht_N,sht_1_N,sht_2_BN,signal_n,signal_bn,Temp_n,Temp_bn,Status_n,Status_bn,tradeno_n,tradeno_bn,Qty_n,Qty_bn,cc_lst_n,cc_lst_bn,multiplier,step_value_n,step_value_bn
    L_expiry = '23511'
    kite = zrd_login.kite
    wb = xw.Book('oi_oi_change_buying.xlsx')
    sht_1_N = wb.sheets['Sheet1_N']
    sht_2_BN = wb.sheets['Sheet2_BN']
    sht_1_N.range("a1:az500").value = None
    sht_2_BN.range("a1:az500").value = None
    sht_N = wb.sheets['NIFTY']
    sht_BN = wb.sheets['BANKNIFTY']
    sht_N.range("a1:az500").value = None
    sht_BN.range("a1:az500").value = None
    Temp_n = {'signal_n':None,'ce_atm_n':None,'pe_atm_n':None,'ltp_ce_n':None,'ltp_pe_n':None,'entry_level_ce_n':None,'entry_level_pe_n':None,'traded_n':None,'entry_time':None,'date':None,'traded_ce_n':None,'traded_pe_n':None,'atm':None,'buy_price':None,'new_buy_price':None,'target':None,'sl':None,'qty_n':None,'new_sl':None,'sell_price':None,'pnl':None,'remark':None,'traded_ok_n':None}
    Temp_bn = {'signal_bn':None,'ce_atm_bn':None,'pe_atm_bn':None,'ltp_ce_bn':None,'ltp_pe_bn':None,'entry_level_ce_bn':None,'entry_level_pe_bn':None,'traded_bn':None,'entry_time':None,'date':None,'traded_ce_bn':None,'traded_pe_bn':None,'atm':None,'buy_price':None,'new_buy_price':None,'target':None,'sl':None,'qty_bn':None,'new_sl':None,'sell_price':None,'pnl':None,'remark':None,'traded_ok_bn':None}
    Status_n = {}
    Status_bn = {}
    tradeno_n = 0
    tradeno_bn = 0
    Status_n[tradeno_n] = Temp_n.copy()
    Status_bn[tradeno_bn] = Temp_bn.copy()
    Qty_n = 50
    cc_lst_n =[]
    step_value_n = 50
    Qty_bn = 25
    cc_lst_bn =[]
    step_value_bn = 100
    multiplier = 0
    
   
init()
while True:
    completed_candle = pd.Series(datetime.datetime.now()).dt.floor('5min')[0]- datetime.timedelta(minutes=5)
    completed_candle = completed_candle.strftime("%Y-%m-%d %H:%M:%S+05:30")
    
    try:
        ctime1 = datetime.datetime.now().time()
        ctime = datetime.datetime.now()
        if completed_candle not in cc_lst_bn:
            ctime1 = datetime.datetime.now().time()
            ctime = datetime.datetime.now()
# ........................//////////////////////Excel data for banknifty/////////////////////................................
            ltp_bn = LTP('NIFTY BANK')
            idf = sf.get_data(name = 'NIFTY BANK', segment = 'NSE:', delta = 2, interval= '5minute', continuous= False, oi=False)
            idf = idf.set_index(idf['date'])
            

            atm_strike = round(ltp_bn/step_value_bn)* step_value_bn + multiplier*step_value_bn
            CE_ATM = ('BANKNIFTY' + L_expiry + str(atm_strike ) + 'CE' )
            PE_ATM = ('BANKNIFTY' + L_expiry + str(atm_strike ) + 'PE' )
            cdf = sf.get_data(name = CE_ATM, segment = 'NFO:', delta = 2, interval= '5minute', continuous= False, oi=True)
            cdf['previous_OI'] = cdf['oi'].shift(1)
            cdf['COI'] = cdf['previous_OI'] - cdf['oi']
            cdf['COI_IN_LOT'] = (cdf['previous_OI'] - cdf['oi'])/Qty_bn
            cdf['COI_Percentage'] = abs(round(((cdf['previous_OI'] - cdf['oi'])/cdf['oi'])*100,0))
            cdf = cdf.set_index(cdf['date'])
            cdf = cdf.loc[:completed_candle]
            call_COI = cdf.loc[completed_candle]['COI']
            call_COI_Percentage = cdf.loc[completed_candle]['COI_Percentage']
            cdf['OI_IN_LOT'] = cdf['oi'] / Qty_bn


            pdf = sf.get_data(name = PE_ATM, segment = 'NFO:', delta = 2, interval= '5minute', continuous= False, oi=True)
            pdf['previous_OI'] = pdf['oi'].shift(1)
            pdf['COI'] = pdf['previous_OI'] - pdf['oi']
            pdf['COI_IN_LOT'] = (pdf['previous_OI'] - pdf['oi'])/Qty_bn
            pdf['COI_Percentage'] = abs(round(((pdf['previous_OI'] - pdf['oi'])/pdf['oi'])*100,0))
            pdf = pdf.set_index(pdf['date'])
            pdf = pdf.loc[:completed_candle]
            put_COI = pdf.loc[completed_candle]['COI']
            put_COI_Percentage = pdf.loc[completed_candle]['COI_Percentage']
            cdf['index_close'],pdf['index_close'] = [idf['close'] , idf['close']]
            cdf['ATM'] = [round(row['index_close']/step_value_bn)* step_value_bn + multiplier*step_value_bn for index, row in cdf.iterrows()]
            pdf['ATM'] = [round(row['index_close']/step_value_bn)* step_value_bn + multiplier*step_value_bn for index, row in cdf.iterrows()]
             
            cdf['Put_COI'] = pdf['COI']
            cdf['PUT/CALL'] = abs(round(cdf['Put_COI'] / cdf['COI'],2))
            pdf['Call_COI'] = cdf['COI']
            pdf['CALL/PUT'] = abs(round(pdf['Call_COI']/pdf['COI'],2))
            pdf['OI_IN_LOT'] = pdf['oi'] / Qty_bn
            pdf = pdf[[ "close", "oi",'OI_IN_LOT', "COI","COI_Percentage",'COI_IN_LOT', 'Call_COI' , 'CALL/PUT' , 'index_close' , 'ATM']]
            cdf = cdf[[ "close", "oi",'OI_IN_LOT', "COI","COI_Percentage",'COI_IN_LOT','Put_COI','PUT/CALL','index_close' , 'ATM']]

            df_banknifty = pd.concat([cdf, pdf], axis=1)

            sht_2_BN.range("A1").value  =  df_banknifty

            # pdb.set_trace()

            

            ltp_CE_ATM = LTP_NFO(CE_ATM)
            ltp_PE_ATM = LTP_NFO(PE_ATM)
            cc_lst_bn.append(completed_candle)
            wb.save()

            signal_bn = None
            if df_banknifty.loc[completed_candle]['COI_Percentage'][0] > 9  or df_banknifty.loc[completed_candle]['COI_Percentage'][1] > 9:# change dimension
                if df_banknifty.loc[completed_candle]['CALL/PUT']  > 3 or df_banknifty.loc[completed_candle]['PUT/CALL'] > 3:# change dimension
                    signal_bn = 'yes'
                    print('yes_bn') 

        if signal_bn == 'yes' and Status_bn[tradeno_bn]['signal_bn'] is None:
            Status_bn[tradeno_bn]['signal_bn'] = 'yes'
            print('hello_bn')
            Status_bn[tradeno_bn]['ce_atm_bn'] = ATM_finder_CE_BN('NIFTY BANK','BANKNIFTY')
            Status_bn[tradeno_bn]['pe_atm_bn'] = ATM_finder_PE_BN('NIFTY BANK','BANKNIFTY')
            Status_bn[tradeno_bn]['ltp_ce_bn'] = kite.ltp(['NFO:'+ Status_bn[tradeno_bn]['ce_atm_bn']])['NFO:'+ Status_bn[tradeno_bn]['ce_atm_bn']]['last_price']
            Status_bn[tradeno_bn]['ltp_pe_bn'] = kite.ltp(['NFO:'+ Status_bn[tradeno_bn]['pe_atm_bn']])['NFO:'+ Status_bn[tradeno_bn]['pe_atm_bn']]['last_price']
            Status_bn[tradeno_bn]['entry_level_ce_bn'] = Status_bn[tradeno_bn]['ltp_ce_bn'] + 0.15*Status_bn[tradeno_bn]['ltp_ce_bn']# change dimension
            Status_bn[tradeno_bn]['entry_level_pe_bn'] = Status_bn[tradeno_bn]['ltp_pe_bn'] + 0.15*Status_bn[tradeno_bn]['ltp_pe_bn']# change dimension
            sht_BN.range('A1').value = pd.DataFrame(Status_bn).T

        if Status_bn[tradeno_bn]['signal_bn'] == 'yes':
            # print(tradeno_bn)
            ltp_CE_BN = kite.ltp(['NFO:'+ Status_bn[tradeno_bn]['ce_atm_bn']])['NFO:'+ Status_bn[tradeno_bn]['ce_atm_bn']]['last_price']
            ltp_PE_BN = kite.ltp(['NFO:'+ Status_bn[tradeno_bn]['pe_atm_bn']])['NFO:'+ Status_bn[tradeno_bn]['pe_atm_bn']]['last_price']
              
            if ((ltp_CE_BN > Status_bn[tradeno_bn]['entry_level_ce_bn']) or (ltp_PE_BN > Status_bn[tradeno_bn]['entry_level_pe_bn'])) and Status_bn[tradeno_bn]['traded_bn'] is None:
                Status_bn[tradeno_bn]['traded_bn'] = 'yes'
                Status_bn[tradeno_bn]['entry_time'] = str(ctime.time())
                Status_bn[tradeno_bn]['date'] = str(ctime.date())
                sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                if (ltp_CE_BN > Status_bn[tradeno_bn]['entry_level_ce_bn']) and Status_bn[tradeno_bn]['traded_ce_bn'] is None:
                    print('hii_bn_ce')
                    Status_bn[tradeno_bn]['traded_ce_bn'] = 'yes'
                    Status_bn[tradeno_bn]['atm'] = Status_bn[tradeno_bn]['ce_atm_bn']
                    Status_bn[tradeno_bn]['buy_price'] = ltp_CE_BN
                    Status_bn[tradeno_bn]['new_buy_price'] = Status_bn[tradeno_bn]['buy_price']
                    Status_bn[tradeno_bn]['target'] = 4 * (Status_bn[tradeno_bn]['buy_price'])# change dimension 
                    Status_bn[tradeno_bn]['sl'] = (Status_bn[tradeno_bn]['buy_price']) - 0.25*(Status_bn[tradeno_bn]['buy_price'])#change dimension
                    Status_bn[tradeno_bn]['qty_bn'] = Qty_bn
                    sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                if (ltp_PE_BN > Status_bn[tradeno_bn]['entry_level_pe_bn']) and Status_bn[tradeno_bn]['traded_pe_bn'] is None:
                    print('hii_bn_pe')                        
                    Status_bn[tradeno_bn]['traded_pe_bn'] = 'yes'
                    Status_bn[tradeno_bn]['atm'] = Status_bn[tradeno_bn]['pe_atm_bn']
                    Status_bn[tradeno_bn]['buy_price'] = ltp_PE_BN
                    Status_bn[tradeno_bn]['new_buy_price'] = Status_bn[tradeno_bn]['buy_price']
                    Status_bn[tradeno_bn]['target'] = 4 * (Status_bn[tradeno_bn]['buy_price'])# change dimension
                    Status_bn[tradeno_bn]['sl'] = (Status_bn[tradeno_bn]['buy_price']) - 0.25*(Status_bn[tradeno_bn]['buy_price'])# change dimension
                    Status_bn[tradeno_bn]['qty_bn'] = Qty_bn
                    sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
            if (Status_bn[tradeno_bn]['traded_bn'] == 'yes') and (Status_bn[tradeno_bn]['traded_ce_bn'] == 'yes') :
                x = (Status_bn[tradeno_bn]['new_buy_price']) + 0.10*(Status_bn[tradeno_bn]['new_buy_price'])# change dimension
                if (ltp_CE_BN > x):
                    Status_bn[tradeno_bn]['new_sl'] = Status_bn[tradeno_bn]['sl'] + (0.05*x)# change dimension
                    Status_bn[tradeno_bn]['new_buy_price'] = x
                    Status_bn[tradeno_bn]['sl'] = Status_bn[tradeno_bn]['new_sl']
                    sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
            if (Status_bn[tradeno_bn]['traded_bn'] == 'yes') and (Status_bn[tradeno_bn]['traded_pe_bn'] == 'yes'):
                x = (Status_bn[tradeno_bn]['new_buy_price']) + 0.10*(Status_bn[tradeno_bn]['new_buy_price'])# change dimension
                if (ltp_PE_BN > x):
                    Status_bn[tradeno_bn]['new_sl'] = Status_bn[tradeno_bn]['sl'] + (0.05*x)# change dimension
                    Status_bn[tradeno_bn]['new_buy_price'] = x
                    Status_bn[tradeno_bn]['sl'] = Status_bn[tradeno_bn]['new_sl']
                    sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
            if (Status_bn[tradeno_bn]['traded_bn'] == 'yes') and Status_bn[tradeno_bn]['traded_ok_bn'] is None:
                if (Status_bn[tradeno_bn]['traded_ce_bn'] == 'yes') and ((ltp_CE_BN > Status_bn[tradeno_bn]['target']) or (ltp_CE_BN < Status_bn[tradeno_bn]['sl'])):
                    print('hello_sell_bce')
                    Status_bn[tradeno_bn]['traded_ok_bn'] = 'yes'
                    Status_bn[tradeno_bn]['sell_price'] = ltp_CE_BN
                    Status_bn[tradeno_bn]['pnl'] = (Status_bn[tradeno_bn]['sell_price'] - Status_bn[tradeno_bn]['buy_price'])* Status_bn[tradeno_bn]['qty_bn']
                    sht_BN.range('A1').value = pd.DataFrame(Status_bn).T                        
                    if (ltp_CE_BN > Status_bn[tradeno_bn]['target']) and Status_bn[tradeno_bn]['remark'] is None:
                        Status_bn[tradeno_bn]['remark'] = 'target_hit'
                        # Status_bn[tradeno_bn] = Temp_bn.copy()
                        sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                        # pdb.set_trace()
                        tradeno_bn = tradeno_bn + 1
                        Status_bn[tradeno_bn] = Temp_bn.copy()
                        print('exit_bce_tgt')

                    if (ltp_CE_BN < Status_bn[tradeno_bn]['sl']) and Status_bn[tradeno_bn]['remark'] is None:
                        Status_bn[tradeno_bn]['remark'] = 'sl_hit'
                        # Status_bn[tradeno_bn] = Temp_bn.copy()
                        sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                        # pdb.set_trace()
                        tradeno_bn = tradeno_bn + 1
                        Status_bn[tradeno_bn] = Temp_bn.copy()
                        print('exit_bce_sl')

                if (Status_bn[tradeno_bn]['traded_pe_bn'] == 'yes') and ((ltp_PE_BN > Status_bn[tradeno_bn]['target']) or (ltp_PE_BN < Status_bn[tradeno_bn]['sl'])):
                    print('hello_sell_bpe')
                    Status_bn[tradeno_bn]['traded_ok_bn'] = 'yes'
                    Status_bn[tradeno_bn]['sell_price'] = ltp_PE_BN
                    Status_bn[tradeno_bn]['pnl'] = (Status_bn[tradeno_bn]['sell_price'] - Status_bn[tradeno_bn]['buy_price'])* Status_bn[tradeno_bn]['qty_bn']
                    sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                    if (ltp_PE_BN > Status_bn[tradeno_bn]['target']) and Status_bn[tradeno_bn]['remark'] is None:
                        Status_bn[tradeno_bn]['remark'] = 'target_hit'
                        # Status_bn[tradeno_bn] = Temp_bn.copy()
                        sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                        # pdb.set_trace()
                        tradeno_bn = tradeno_bn + 1
                        Status_bn[tradeno_bn] = Temp_bn.copy()
                        print('exit_bpe_sl')

                    if (ltp_PE_BN < Status_bn[tradeno_bn]['sl']) and Status_bn[tradeno_bn]['remark'] is None:
                        Status_bn[tradeno_bn]['remark'] = 'sl_hit'
                        # Status_bn[tradeno_bn] = Temp_bn.copy()
                        sht_BN.range('A1').value = pd.DataFrame(Status_bn).T
                        # pdb.set_trace()
                        tradeno_bn = tradeno_bn + 1
                        Status_bn[tradeno_bn] = Temp_bn.copy()
                        print('exit_bpe_tgt')


        if completed_candle not in cc_lst_n:
            ctime1 = datetime.datetime.now().time()
            ctime = datetime.datetime.now()
# ........................//////////////////////Excel data for nifty/////////////////////................................

            ltp_n = LTP('NIFTY 50')
            idf = sf.get_data(name = 'NIFTY 50', segment = 'NSE:', delta = 2, interval= '5minute', continuous= False, oi=False)
            idf = idf.set_index(idf['date'])
            

            atm_strike = round(ltp_n/step_value_n)* step_value_n + multiplier*step_value_n
            CE_ATM = ('NIFTY' + L_expiry + str(atm_strike ) + 'CE' )
            PE_ATM = ('NIFTY' + L_expiry + str(atm_strike ) + 'PE' )
            cdf = sf.get_data(name = CE_ATM, segment = 'NFO:', delta = 2, interval= '5minute', continuous= False, oi=True)
            cdf['previous_OI'] = cdf['oi'].shift(1)
            cdf['COI'] = cdf['previous_OI'] - cdf['oi']
            cdf['COI_IN_LOT'] = (cdf['previous_OI'] - cdf['oi'])/Qty_bn
            cdf['COI_Percentage'] = abs(round(((cdf['previous_OI'] - cdf['oi'])/cdf['oi'])*100,0))
            cdf = cdf.set_index(cdf['date'])
            cdf = cdf.loc[:completed_candle]
            call_COI = cdf.loc[completed_candle]['COI']
            call_COI_Percentage = cdf.loc[completed_candle]['COI_Percentage']
            cdf['OI_IN_LOT'] = cdf['oi'] / Qty_bn



            pdf = sf.get_data(name = PE_ATM, segment = 'NFO:', delta = 2, interval= '5minute', continuous= False, oi=True)
            pdf['previous_OI'] = pdf['oi'].shift(1)
            pdf['COI'] = pdf['previous_OI'] - pdf['oi']
            pdf['COI_IN_LOT'] = (pdf['previous_OI'] - pdf['oi'])/Qty_bn
            pdf['COI_Percentage'] = abs(round(((pdf['previous_OI'] - pdf['oi'])/pdf['oi'])*100,0))
            pdf = pdf.set_index(pdf['date'])
            pdf = pdf.loc[:completed_candle]
            put_COI = pdf.loc[completed_candle]['COI']
            put_COI_Percentage = pdf.loc[completed_candle]['COI_Percentage']
            cdf['index_close'],pdf['index_close'] = [idf['close'] , idf['close']]
            cdf['ATM'] = [round(row['index_close']/step_value_n)* step_value_n + multiplier*step_value_n for index, row in cdf.iterrows()]
            pdf['ATM'] = [round(row['index_close']/step_value_n)* step_value_n + multiplier*step_value_n for index, row in cdf.iterrows()]
             
            cdf['Put_COI'] = pdf['COI']
            cdf['PUT/CALL'] = abs(round(cdf['Put_COI'] / cdf['COI'],2))
            pdf['Call_COI'] = cdf['COI']
            pdf['CALL/PUT'] = abs(round(pdf['Call_COI']/pdf['COI'],2))
            pdf['OI_IN_LOT'] = pdf['oi'] / Qty_bn
            pdf = pdf[[ "close", "oi",'OI_IN_LOT', "COI","COI_Percentage",'COI_IN_LOT', 'Call_COI' , 'CALL/PUT' , 'index_close' , 'ATM']]
            cdf = cdf[[ "close", "oi",'OI_IN_LOT', "COI","COI_Percentage",'COI_IN_LOT','Put_COI','PUT/CALL','index_close' , 'ATM']]

            df_nifty = pd.concat([cdf, pdf], axis=1)

            sht_1_N.range("A1").value  =  df_nifty


            

            ltp_CE_ATM = LTP_NFO(CE_ATM)
            ltp_PE_ATM = LTP_NFO(PE_ATM)
            cc_lst_n.append(completed_candle)
            wb.save()

            signal_n = None
            if df_nifty.loc[completed_candle]['COI_Percentage'][0] > 9  or df_nifty.loc[completed_candle]['COI_Percentage'][1] > 9:# change dimension
                if df_nifty.loc[completed_candle]['CALL/PUT']  > 3 or df_nifty.loc[completed_candle]['PUT/CALL'] > 3:# change dimension
                    signal_n = 'yes'
                    print('yes_n')


        if signal_n == 'yes' and Status_n[tradeno_n]['signal_n'] is None:
            Status_n[tradeno_n]['signal_n'] = 'yes'
            print('hello_n')
            # pdb.set_trace()
            Status_n[tradeno_n]['ce_atm_n'] = ATM_finder_CE_N('NIFTY 50','NIFTY')
            Status_n[tradeno_n]['pe_atm_n'] = ATM_finder_PE_N('NIFTY 50','NIFTY')
            Status_n[tradeno_n]['ltp_ce_n'] = kite.ltp(['NFO:'+ Status_n[tradeno_n]['ce_atm_n']])['NFO:'+ Status_n[tradeno_n]['ce_atm_n']]['last_price']
            Status_n[tradeno_n]['ltp_pe_n'] = kite.ltp(['NFO:'+ Status_n[tradeno_n]['pe_atm_n']])['NFO:'+ Status_n[tradeno_n]['pe_atm_n']]['last_price']
            Status_n[tradeno_n]['entry_level_ce_n'] = Status_n[tradeno_n]['ltp_ce_n'] + 0.15*Status_n[tradeno_n]['ltp_ce_n']# change dimension
            Status_n[tradeno_n]['entry_level_pe_n'] = Status_n[tradeno_n]['ltp_pe_n'] + 0.15*Status_n[tradeno_n]['ltp_pe_n']# change dimension
            sht_N.range('A1').value = pd.DataFrame(Status_n).T
                
        if Status_n[tradeno_n]['signal_n'] == 'yes':
            # print('hii_n')
            ltp_CE_N = kite.ltp(['NFO:'+ Status_n[tradeno_n]['ce_atm_n']])['NFO:'+ Status_n[tradeno_n]['ce_atm_n']]['last_price']
            ltp_PE_N = kite.ltp(['NFO:'+ Status_n[tradeno_n]['pe_atm_n']])['NFO:'+ Status_n[tradeno_n]['pe_atm_n']]['last_price']
              
            if ((ltp_CE_N > Status_n[tradeno_n]['entry_level_ce_n']) or (ltp_PE_N > Status_n[tradeno_n]['entry_level_pe_n'])) and Status_n[tradeno_n]['traded_n'] is None:
                Status_n[tradeno_n]['traded_n'] = 'yes'
                Status_n[tradeno_n]['entry_time'] = str(ctime.time())
                Status_n[tradeno_n]['date'] = str(ctime.date())
                sht_N.range('A1').value = pd.DataFrame(Status_n).T
                if (ltp_CE_N > Status_n[tradeno_n]['entry_level_ce_n']) and Status_n[tradeno_n]['traded_ce_n'] is None:
                    print('hii_n_ce')
                    Status_n[tradeno_n]['traded_ce_n'] = 'yes'
                    Status_n[tradeno_n]['atm'] = Status_n[tradeno_n]['ce_atm_n']
                    Status_n[tradeno_n]['buy_price'] = ltp_CE_N
                    Status_n[tradeno_n]['new_buy_price'] = Status_n[tradeno_n]['buy_price']
                    Status_n[tradeno_n]['target'] = 4 * (Status_n[tradeno_n]['buy_price'])# change dimension 
                    Status_n[tradeno_n]['sl'] = (Status_n[tradeno_n]['buy_price']) - 0.25*(Status_n[tradeno_n]['buy_price'])# change dimension
                    Status_n[tradeno_n]['qty_n'] = Qty_n
                    sht_N.range('A1').value = pd.DataFrame(Status_n).T
                if (ltp_PE_N > Status_n[tradeno_n]['entry_level_pe_n']) and Status_n[tradeno_n]['traded_pe_n'] is None:
                    print('hii_n_pe')                        
                    Status_n[tradeno_n]['traded_pe_n'] = 'yes'
                    Status_n[tradeno_n]['atm'] = Status_n[tradeno_n]['pe_atm_n']
                    Status_n[tradeno_n]['buy_price'] = ltp_PE_N
                    Status_n[tradeno_n]['new_buy_price'] = Status_n[tradeno_n]['buy_price']
                    Status_n[tradeno_n]['target'] = 4 * (Status_n[tradeno_n]['buy_price'])# change dimension
                    Status_n[tradeno_n]['sl'] = (Status_n[tradeno_n]['buy_price']) - 0.25*(Status_n[tradeno_n]['buy_price'])# change dimension
                    Status_n[tradeno_n]['qty_n'] = Qty_n
                    sht_N.range('A1').value = pd.DataFrame(Status_n).T
            if (Status_n[tradeno_n]['traded_n'] == 'yes') and (Status_n[tradeno_n]['traded_ce_n'] == 'yes') :
                x = (Status_n[tradeno_n]['new_buy_price']) + 0.10*(Status_n[tradeno_n]['new_buy_price'])# change dimension
                if (ltp_CE_N > x):
                    Status_n[tradeno_n]['new_sl'] = Status_n[tradeno_n]['sl'] + (0.05*x)# change dimension
                    Status_n[tradeno_n]['new_buy_price'] = x
                    Status_n[tradeno_n]['sl'] = Status_n[tradeno_n]['new_sl']
                    sht_N.range('A1').value = pd.DataFrame(Status_n).T
            if (Status_n[tradeno_n]['traded_n'] == 'yes') and (Status_n[tradeno_n]['traded_pe_n'] == 'yes'):
                x = (Status_n[tradeno_n]['new_buy_price']) + 0.10*(Status_n[tradeno_n]['new_buy_price'])# change dimension
                if (ltp_PE_N > x):
                    Status_n[tradeno_n]['new_sl'] = Status_n[tradeno_n]['sl'] + (0.05*x)# change dimension
                    Status_n[tradeno_n]['new_buy_price'] = x
                    Status_n[tradeno_n]['sl'] = Status_n[tradeno_n]['new_sl']
                    sht_N.range('A1').value = pd.DataFrame(Status_n).T
            if (Status_n[tradeno_n]['traded_n'] == 'yes') and Status_n[tradeno_n]['traded_ok_n'] is None:
                if (Status_n[tradeno_n]['traded_ce_n'] == 'yes') and ((ltp_CE_N > Status_n[tradeno_n]['target']) or (ltp_CE_N < Status_n[tradeno_n]['sl'])):
                    print('hello_sell_nce')
                    Status_n[tradeno_n]['traded_ok_n'] = 'yes'
                    Status_n[tradeno_n]['sell_price'] = ltp_CE_N
                    Status_n[tradeno_n]['pnl'] = (Status_n[tradeno_n]['sell_price'] - Status_n[tradeno_n]['buy_price'])* Status_n[tradeno_n]['qty_n']
                    sht_N.range('A1').value = pd.DataFrame(Status_n).T                        
                    if (ltp_CE_N > Status_n[tradeno_n]['target']) and Status_n[tradeno_n]['remark'] is None:
                        Status_n[tradeno_n]['remark'] = 'target_hit'
                        # Status_n[tradeno_n] = Temp_n.copy()
                        sht_N.range('A1').value = pd.DataFrame(Status_n).T
                        # pdb.set_trace()
                        tradeno_n = tradeno_n + 1
                        Status_n[tradeno_n] = Temp_n.copy()
                        print('exit_nce_tgt')

                    if (ltp_CE_N < Status_n[tradeno_n]['sl']) and Status_n[tradeno_n]['remark'] is None:
                        Status_n[tradeno_n]['remark'] = 'sl_hit'
                        # Status_n[tradeno_n] = Temp_n.copy()
                        sht_N.range('A1').value = pd.DataFrame(Status_n).T
                        # pdb.set_trace()
                        tradeno_n = tradeno_n + 1
                        Status_n[tradeno_n] = Temp_n.copy()
                        print('exit_nce_sl')

                if (Status_n[tradeno_n]['traded_pe_n'] == 'yes') and ((ltp_PE_N > Status_n[tradeno_n]['target']) or (ltp_PE_N < Status_n[tradeno_n]['sl'])):
                    print('hello_sell_npe')
                    Status_n[tradeno_n]['traded_ok_n'] = 'yes'
                    Status_n[tradeno_n]['sell_price'] = ltp_PE_N
                    Status_n[tradeno_n]['pnl'] = (Status_n[tradeno_n]['sell_price'] - Status_n[tradeno_n]['buy_price'])* Status_n[tradeno_n]['qty_n']
                    sht_N.range('A1').value = pd.DataFrame(Status_n).T
                    if (ltp_PE_N > Status_n[tradeno_n]['target']) and Status_n[tradeno_n]['remark'] is None:
                        Status_n[tradeno_n]['remark'] = 'target_hit'
                        # Status_n[tradeno_n] = Temp_n.copy()
                        sht_N.range('A1').value = pd.DataFrame(Status_n).T
                        # pdb.set_trace()
                        tradeno_n = tradeno_n + 1
                        Status_n[tradeno_n] = Temp_n.copy()
                        print('exit_npe_tgt')

                    if (ltp_PE_N < Status_n[tradeno_n]['sl']) and Status_n[tradeno_n]['remark'] is None:
                        Status_n[tradeno_n]['remark'] = 'sl_hit'
                        # Status_n[tradeno_n] = Temp_n.copy()
                        sht_N.range('A1').value = pd.DataFrame(Status_n).T
                        # pdb.set_trace()s
                        tradeno_n = tradeno_n + 1
                        Status_n[tradeno_n] = Temp_n.copy()
                        print('exit_npe_sl')



        

    except Exception as e:
        print(e)
        continue
    
    
               