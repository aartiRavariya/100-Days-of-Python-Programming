import logging
import pandas as pd
import datetime
import math
import numpy as np
from tkinter import filedialog
from tkinter import *
import time; time.sleep(.5)
from dateutil.relativedelta import relativedelta

print('Enter Annuity file')
filename =  filedialog.askopenfilename(initialdir = "/",title = "Select annuity file",filetypes = (("csv","*.csv"),("all files","*.*")))
print(filename)
sheet = pd.read_csv(filename) 
df1 = pd.DataFrame(sheet)
dict1 = df1.to_dict(orient='records')



sheet2 = pd.read_excel('annuity_rates_series_from_1.1-1.10.xlsx')
df3 = pd.DataFrame(sheet2)
dict3 = df3.to_dict(orient='records')

sheet3 = pd.read_excel('annuity_rates_option2.1.xlsx')
df4 = pd.DataFrame(sheet3)
dict4 = df4.to_dict(orient='records')

sheet4 = pd.read_excel('annuity_rates_option2.2.xlsx')
df5 = pd.DataFrame(sheet4)
dict5 = df5.to_dict(orient='records')

sheet5 = pd.read_excel('annuity_rates_option2.3.xlsx')
df6 = pd.DataFrame(sheet5)
dict6 = df6.to_dict(orient='records')

sheet6 = pd.read_excel('annuity_rates_option2.4.xlsx')
df7 = pd.DataFrame(sheet6)
dict7 = df7.to_dict(orient='records')

sheet7 = pd.read_excel('annuity_rates_option2.5.xlsx')
df8 = pd.DataFrame(sheet7)
dict8 = df8.to_dict(orient='records')

sheet8 = pd.read_excel('annuity_rates_option2.5_double_life.xlsx')
df9 = pd.DataFrame(sheet8)
dict9 = df9.to_dict(orient='records')

if __name__ == '__main__':
        
        proposal_no = []
        uin = []
        Ia_quote = []
        sbi_discount_code = []
        distance_marketing_flag = []
        Annuitant1_date_of_birth = []
        Annuitant2_date_of_birth = []
        source_of_business =[]
        agent_id =[]
        purchase_price = []
        Purchase_price_excluding_service = []
        annuity_inst_amt=[]
        pay_mode =[]
        annuity_option =[]
        annuity_advance_flag = []
        annuity_policy_cat = []
        annuity_type = []
        Diff_quotedate_annuitant1 = []
        Diff_quotedate_annuitant2 = []
        Servicetax = []
        Servicetax_amount = []
        Net_purchase_price = []
        Annuity_rate =[]
        Hpp=[]
        plan_id = []
        gst_region = []
        Frequency_Loading = []
        Frequency_Factor = []
        Annuity_final_rates =[]
        Divisor_base_on_frequency=[]
        Net_Annuity_Amount=[]
        Nps_special_Incentive = []
        annuity_Amount_periodically = []
        result = []
        res = []
        acct_hist_amt = []
        Pol_iss_loc_cd = []


        


for i in dict1:
        print('-------------PROPOSAL_NO ---------',i['PROPOSAL_NO'])
        age = ''
        hpp = 0.0
        NPS_special_Incentive = 0
        frequency_Loading = 0.0
        frequency_Factor = 0.0
        diff_quotedate_annuitant2_conv = ''
        diff_quotedate_annuitant1_conv = ''
        servicetax_conversion = 0.0
        divisor_base_on_frequency = 0.0
        purchase_price_excluding_service = 0.0
        servicetax = 0.0
        annuity_rate = 0.0
        servicetax_amount = 0.0
        UIN = str(i['UIN'])
        U = UIN[-3:]
        agent = str(i['AGENT_ID']).strip()
        

        Ia_quote_date = datetime.datetime.strptime(i['IA_QUOTE_DATE'],"%d/%m/%Y").date()
        Annuitant1_dob = datetime.datetime.strptime(i['ANNUITANT1_DOB'],"%d/%m/%Y").date()
        Annuitant2_dob = datetime.datetime.strptime(i['ANNUITANT2_DOB'],"%d/%m/%Y").date()
        annuitant2_age =Annuitant2_dob.strftime("%d/%m/%Y")
        if annuitant2_age == '01/01/1940':
                age = i['IA_QUOTE_DATE']
                Age = datetime.datetime.strptime(age,"%d/%m/%Y").date()

        else:
                age = Annuitant2_dob
                Age = Annuitant2_dob




        diff_date_annuitant1 = Ia_quote_date - Annuitant1_dob
        diff_in_days = diff_date_annuitant1.days
        diff_quotedate_annuitant1 = (diff_in_days/365.25)
        diff_quotedate_annuitant1 = int(diff_quotedate_annuitant1)
        
		diff_date_annuitant2 = Ia_quote_date - Age
        diff_in_annuitant2_date = diff_date_annuitant2.days
        diff_quotedate_annuitant2 =   (diff_in_annuitant2_date/365.25)
        diff_quotedate_annuitant2 =  int(diff_quotedate_annuitant2)
        if U in 'V04':
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V4'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V4'
                
        elif U in 'V05':
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V5'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V5'
        elif U in 'V06':
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V6'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V6'
        elif U in 'V07':
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V7'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V7'
        elif U in 'V08':
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V8'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V8'
        elif U in 'V09':
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V9'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V9'
        else:
                diff_quotedate_annuitant2_conv = str(diff_quotedate_annuitant2)+'_V10'
                diff_quotedate_annuitant1_conv = str(diff_quotedate_annuitant1)+'_V10'
        
        
        


        start_dat = '01/06/2016'
        start_date = datetime.datetime.strptime(start_dat,"%d/%m/%Y").date()
        end_dat = '30/06/2017'
        end_date = datetime.datetime.strptime(end_dat,"%d/%m/%Y").date()
        start_dat_1 = '01/07/2017'
        start_date_1 = datetime.datetime.strptime(start_dat_1,"%d/%m/%Y").date()
        end_dat_1 = '30/06/2020'
        end_date_1 = datetime.datetime.strptime(end_dat_1,"%d/%m/%Y").date()
        

        if Ia_quote_date >=start_date and Ia_quote_date<=end_date:
                servicetax = 1.5
        elif Ia_quote_date>=start_date_1 and Ia_quote_date<=end_date_1:
                servicetax = 1.8
        else:
                servicetax = 0
        if i['GST_REGION'] == 15 and i['POL_ISS_LOC_CD'] == 15:
                servicetax = 1.9
        if i['ANNUITY_POLICY_CAT'] == 'N' or  i['SOURCE_OF_BUSINESS'] == 'N':
                servicetax = 0
        



        servicetax_conversion = (1+servicetax/100)
        
        servicetax_amount = i['PURCHASE_PRICE']  - (i['PURCHASE_PRICE']/servicetax_conversion)
        
        purchase_price_excluding_service =  i['PURCHASE_PRICE'] -  servicetax_amount 

        for l in dict3:
                for key,value in l.items():
                        if l['Ages']==diff_quotedate_annuitant1 and i['PLAN_ID'] == key:
                                annuity_rate= value
                
        if i['ANNUITY_OPTION'] == 2.1:
                for m in dict4:
                        for key,value in m.items():
                                if ((diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 4'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 5'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 6 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 7'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 8'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 9 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and m['Annuitant2 age version 10'] == diff_quotedate_annuitant2_conv)):
                                        annuity_rate = value
                                        
                                        
        if i['ANNUITY_OPTION'] == 2.2:
                for n in dict5:
                        for key,value in n.items():
                                if ((diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 4'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 5'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 6 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 7'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 8'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 9 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and n['Annuitant2 age version 10'] == diff_quotedate_annuitant2_conv)):
                                        annuity_rate = value
                                       

        if i['ANNUITY_OPTION'] ==2.3:
                for o in dict6:
                        for key,value in o.items():
                                if ((diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 4'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 5'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 6 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 7'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 8'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 9 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and o['Annuitant2 age version 10'] == diff_quotedate_annuitant2_conv)):
                                        annuity_rate = value
                                              
        if i['ANNUITY_OPTION'] ==2.4:
                for p in dict7:
                        for key,value in p.items():
                                if ((diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 4'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 5'] == diff_quotedate_annuitant2_conv) or
                                (diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 6 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 7'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 8'] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 9 '] == diff_quotedate_annuitant2_conv)  or
                                (diff_quotedate_annuitant1_conv == key and p['Annuitant2 age version 10'] == diff_quotedate_annuitant2_conv)):
                                        annuity_rate = value
                                      

        if i['ANNUITY_OPTION'] ==2.5:
                if diff_quotedate_annuitant2==0:
                        for q in dict8:
                                for key,value in q.items():
                                        if diff_quotedate_annuitant1 == q['Annuitant1 Age'] and i['PLAN_ID'] == key:
                                                annuity_rate = value
              
                                                      

        if i['ANNUITY_OPTION'] == 2.5:
                if diff_quotedate_annuitant2 != 0:
                        for r in dict9:
                                for key,value in r.items():
                                        if ((diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 4'] == diff_quotedate_annuitant2_conv) or
                                            (diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 5'] == diff_quotedate_annuitant2_conv) or
                                            (diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 6 '] == diff_quotedate_annuitant2_conv)  or
                                            (diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 7 '] == diff_quotedate_annuitant2_conv)  or
                                            (diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 8 '] == diff_quotedate_annuitant2_conv)  or
                                            (diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 9 '] == diff_quotedate_annuitant2_conv)  or
                                            (diff_quotedate_annuitant1_conv == key and r['Annuitant2 age version 10 '] == diff_quotedate_annuitant2_conv)):
                                                annuity_rate = value
                                              



        


        
        if U in 'V04':
                if purchase_price_excluding_service<=50000:
                        HPP=0
                elif purchase_price_excluding_service>=150000 and purchase_price_excluding_service<=299999:
                        HPP=2.50
                elif purchase_price_excluding_service>=300000 and  purchase_price_excluding_service<=499999:
                        HPP=3.50
                elif purchase_price_excluding_service>=500000:
                        HPP=4.25
                else:
                        HPP=0
        else:
                if purchase_price_excluding_service>=1000000 and  purchase_price_excluding_service<=1499999:
                        HPP=0.5
                elif purchase_price_excluding_service>=1499999 :
                        HPP=1.00
                else:
                        HPP=0
                
        if i['ANNUITY_PAYOUT_FREQUENCY'] == 1:
                frequency_loading = 1
                frequency_factor = 12
        elif i['ANNUITY_PAYOUT_FREQUENCY'] == 3:
                frequency_loading = 1.0075
                frequency_factor = 4
        elif i['ANNUITY_PAYOUT_FREQUENCY'] == 6:
                frequency_loading = 1.0175
                frequency_factor = 2
        else:
                frequency_loading = 1.035
                frequency_factor = 1
        
        Annuity_final_r = (annuity_rate * frequency_loading)
        annuity_final_rates = Annuity_final_r + HPP
        
        if i['SOURCE_OF_BUSINESS']=='N' or i['SOURCE_OF_BUSINESS'] == 'S' or i['DISTANCE_MARKETING_FLAG']=='Y' or i['ANNUITY_POLICY_CAT']=='N' or agent=='DMAIL' or i['SBI_DISCOUNT_CODE']=='Y' or agent=='DONLINE' or  i['ANNUITY_POLICY_CAT']=='S':
                divisor_base_on_frequency = 980
        else:
                divisor_base_on_frequency = 1000
                
        Net_Annuity_Amt = (annuity_final_rates/frequency_factor)
        net_amt = Net_Annuity_Amt / divisor_base_on_frequency
        net_Annuity_Amount = net_amt * purchase_price_excluding_service
        
        if i['SOURCE_OF_BUSINESS']=='N' or i['ANNUITY_POLICY_CAT']=='N':
                NPS_special_Incentive  = net_Annuity_Amount * 0.75 / 100
        else:
                NPS_special_Incentive = 0
        
        Annuity_Amount_periodically = net_Annuity_Amount + NPS_special_Incentive
        

        final_result = i['ANNUITY_INST_AMT'] - Annuity_Amount_periodically
                        

        if final_result>=-5 and final_result<=5:
                result.append('PASS')
        else:
                result.append('FAIL')

        temp = i['ANNUITY_INST_AMT'] - i['ACCT_HIST_AMT']
        if temp>=-5 and temp<=5:
                res.append('PASS')
        else:
                res.append('FAIL')
                
        proposal_no.append(i['PROPOSAL_NO'])
        Ia_quote.append(i['IA_QUOTE_DATE'])
        uin.append(i['UIN'])
        sbi_discount_code.append(i['SBI_DISCOUNT_CODE'])
        distance_marketing_flag.append(i['DISTANCE_MARKETING_FLAG'])
        Annuitant1_date_of_birth.append(i['ANNUITANT1_DOB'])
        Annuitant2_date_of_birth.append(i['ANNUITANT2_DOB'])
        source_of_business.append(i['SOURCE_OF_BUSINESS'])
        agent_id.append(agent)
        purchase_price.append(i['PURCHASE_PRICE'])
        Net_purchase_price.append(i['NET_PURCHASE_PRICE'])
        Purchase_price_excluding_service.append(purchase_price_excluding_service)
        Servicetax.append(servicetax)
        annuity_inst_amt.append(i['ANNUITY_INST_AMT'])
        pay_mode.append(i['ANNUITY_PAYOUT_FREQUENCY'])
        annuity_option.append(i['ANNUITY_OPTION'])
        annuity_advance_flag.append(i['ANNUITY_ADVANCE_FLAG'])
        annuity_policy_cat.append(i['ANNUITY_POLICY_CAT'])
        Diff_quotedate_annuitant1.append(diff_quotedate_annuitant1)
        Diff_quotedate_annuitant2.append(diff_quotedate_annuitant2)
        Annuity_rate.append(annuity_rate)
        Hpp.append(HPP)
        plan_id.append(i['PLAN_ID'])
        Frequency_Loading.append(frequency_loading)
        Frequency_Factor.append(frequency_factor)
        Annuity_final_rates.append(annuity_final_rates)
        Divisor_base_on_frequency.append(divisor_base_on_frequency)
        Net_Annuity_Amount.append(net_Annuity_Amount)
        Nps_special_Incentive.append(NPS_special_Incentive)
        gst_region.append(i['GST_REGION'])
        annuity_Amount_periodically.append(Annuity_Amount_periodically)
        acct_hist_amt.append(i['ACCT_HIST_AMT'])
        Pol_iss_loc_cd.append(i['POL_ISS_LOC_CD'])

                        



                


o1 = pd.Series(proposal_no,dtype=str,name ='PROPOSAL_NO')
o2 = pd.Series(Ia_quote,dtype=str,name='IA_QUOTE_DATE')
o3 = pd.Series(pay_mode,dtype=str,name='PAY_MODE')
o4 = pd.Series(distance_marketing_flag,dtype=str,name='DISTANCE_MARKETING_FLAG')
o5 = pd.Series(uin,dtype=str,name='UIN')
o6 = pd.Series(Annuitant1_date_of_birth,dtype=str,name='SYS_ANNUITANT1_DOB')
o7 = pd.Series(Annuitant2_date_of_birth,dtype=str,name='SYS_ANNUITANT2_DOB')
o8 = pd.Series(source_of_business,dtype=str,name='SOURCE_OF_BUSINESS')
o9 = pd.Series(agent_id,dtype=str,name='AGENT_ID')
o10 = pd.Series(purchase_price,dtype=str,name='SYS_PURCHASE_PRICE')
o11 = pd.Series(Net_purchase_price,dtype=str,name='SYS_NET_PURCHASE_PRICE')
o12 = pd.Series(Servicetax,dtype=str,name='SERVICE_TAX')
o13 = pd.Series(Purchase_price_excluding_service,dtype=str,name='PURCHASE_PRICE_EXCLUDING_SERVICE')
o14 = pd.Series(annuity_option,dtype=str,name='SYS_ANNUITY_OPTION')
o15 = pd.Series(annuity_advance_flag,dtype=str,name='SYS_ANNUITY_ADVANCE_FLAG')
o16 = pd.Series(annuity_policy_cat,dtype=str,name='SYS_ANNUITY_POLICY_CAT')
o17 = pd.Series(sbi_discount_code,dtype=str,name='SBI_DISCOUNT_CODE')
o18 = pd.Series(Diff_quotedate_annuitant1,dtype=str,name='ANNUITANT1_AGE')
o19 = pd.Series(Diff_quotedate_annuitant2,dtype=str,name='ANNUITANT2_AGE')
o20 = pd.Series(gst_region,dtype=str,name='GST_REGION')
o21 = pd.Series(Pol_iss_loc_cd,dtype=str,name='POL_ISS_LOC_CD')
o22 = pd.Series(Annuity_rate,dtype=str,name='ANNUITY_RATE')
o23 = pd.Series(Hpp,dtype=str,name='HPP_FACTOR')
o24 = pd.Series(plan_id,dtype=str,name='PLAN_ID')
o25 = pd.Series(Frequency_Loading,dtype=str,name='FREQUENCY_LOADING')
o26 = pd.Series(Frequency_Factor,dtype=str,name='FREQUENCY_FACTOR')
o27 = pd.Series(Annuity_final_rates,dtype=str,name='ANNUITY_FINAL_RATES')
o28 = pd.Series(Divisor_base_on_frequency,dtype=str,name='DIVISOR_BASE_ON_FREQUENCY')
o29 = pd.Series(Net_Annuity_Amount,dtype=str,name='NET_ANNUITY_AMOUNT')
o30 = pd.Series(Nps_special_Incentive,dtype=str,name='NPS_SPECIAL_INCENTIVE')
o31 = pd.Series(acct_hist_amt,dtype=str,name='ACCT_HIST_AMT')
o32 = pd.Series(res,dtype=str,name='RESULT_ACCT_HIST_AMT_ANNUITY_INST_AMT')
o33 = pd.Series(annuity_Amount_periodically,dtype=str,name='ANNUITY_AMOUNT_PERIODICALLY')
o34 = pd.Series(annuity_inst_amt,dtype=str,name='SYS_ANNUITY_INST_AMT')
o35 = pd.Series(result,dtype=str,name='AUTOMATED_RESULTS')

data_concat = [o1,o2,o3,o4,o5,o6,o7,o8,o9,o10,o11,o12,o13,o14,o15,o16,o17,o18,o19,o20,o21,o22,o23,o24,o25,o26,o27,o28,o29,o30,o31,o32,o33,o34,o35]

df = pd.concat(data_concat, axis=1) 
print('Succes , choose path to save') 
path_to_save = file_path = filedialog.asksaveasfilename(filetypes = (("csv files", "*.csv"),
                ("all files","*.*")))
try:
        df.to_csv(path_to_save, index=False)
        logging.info('SUCCESS-excel file is generated') 
except:
        logging.info("File is open ERROR")
        exit()
print("csv file is generated at path ")
print(path_to_save)
        
        



                
