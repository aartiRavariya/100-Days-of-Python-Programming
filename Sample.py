#Allocation
import pandas as pd
import datetime
from datetime import timedelta
from dateutil import parser
import math
from itertools import groupby
from collections import OrderedDict
import numpy as np
import logging
import tkinter as tk
from tkinter import filedialog
from tkinter import *
root = tk.Tk()
root.withdraw()
from dateutil.relativedelta import relativedelta
import time

start_time = time.time()
print('Start time:',datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
print("---------------------------------")

df_data1 = pd.read_csv(r"Z:\Aarti\ALLOCATION _CHARGES\Input-files\product-admin-4policies.csv")
file_data1 = df_data1.to_dict(into= OrderedDict,orient="records")

df_data3 = pd.read_csv(r"Z:\Data Quality checking\FY 2021-22\Allocation Charges\raw_data-27_01_2B.csv")
file_data3 = df_data3.to_dict(into= OrderedDict,orient="records")

df_ratesheet = pd.read_excel(r"Z:\Aarti\ALLOCATION _CHARGES\Ratesheets\final_ratesheet_alloc-updated.xlsx")
file_ratesheet = df_ratesheet.to_dict(into= OrderedDict,orient="records")

df_1H = pd.read_excel(r"Z:\Aarti\ALLOCATION _CHARGES\Ratesheets\updated_1H_ratesheet_V2.xlsx")
file_1H = df_1H.to_dict(into= OrderedDict,orient="records")

df_oldprod = pd.read_excel(r"Z:\Aarti\ALLOCATION _CHARGES\Ratesheets\Copy of old_ratesheet_alloc.xlsx")
file_oldprod = df_oldprod.to_dict(into= OrderedDict,orient="records")

pol_id = []
date_of_commencement = []
EMPTY = ['']
list_of_fund = []
fund_list = []
proposal_no_list2,proposal_no_list1 = ([],[])
doc_list2,doc_list1 = ([],[])
premium_list2,premium_list1 = ([],[])
charges2,charges1 = ([],[])
CIA_EFF_DT2,CIA_EFF_DT1 = ([],[])
risk_premium_final2,risk_premium_final1 = ([],[])
date_monthly2,date_monthly1 = ([],[])
reason_code2,reason_code1 = ([],[])
DIFFERENCE,DIFFERENCE1,DIFFERENCE2 = ([],[],[])
CAL_UNITS,CAL_UNITS1,CAL_UNITS2 = ([],[],[])
proposal_no_list_r = []
monthlydate_list_r = []
staff_discount_flag_list = []
ap_list_r = []
ini_amt_list_r = []
cd_type_rs_r = []
final_risk_value_r = []
final_sar_value_r = []
final_mortal_fact_r = []
CIA_EFF_DT_r = []
CAL_CHARGES_r = []
fund_units,fund_units1,fund_units2 = ([],[],[])
value_nav,value_nav1,value_nav2 = ([],[],[])
fnd_name,fnd_name1,fnd_name2 = ([],[],[])
fund_percentage,fund_percentage1,fund_percentage2 = ([],[],[])
fund_v0,fund_v1,fund_v2 = ([],[],[])
result_list, result_list1, result_list2, result_list3 = ([],[],[],[])
sum_fund_mor = []
sum_fund_pel = []
c_p = []
CAL_CHARGES_tot = []
plan_list = []
pol_list = []
date_monthly = []
final_value = []
fund_details_list = []
policy_no_list = []
policy_no_list1 = []
proposal_no_list = []
age_list = []
sar_final_value = []
risk_premium_final = []
cia_cli_trxn_amt = []
ini_no_list = []
month_d = []
policy_no = []
sum_assured_list = []
premium_list = []
doc_list = []
ini_charge = []
policy_year_list = []
mode_list = []
rate_list = []
units_diff_list = []
reason_code = []
CIA_TYP_CD = []
CIA_EFF_DT = []
GST_REGION = []
ISSUE_LOCATION = []
POL_BILL_MODE_CD = []
POL_BILL_TYP_CD = []
CHARGES = []
result_final = ""
sum_at_risk = 0.0
final_risk_charges = 0.0
calculated_fv = []
calculated_fnd_perct = []
sys_perct = []

yest_date = datetime.datetime.now().date()-timedelta(days=1)

counting = 0
total_count = 0
for data1 in file_data1:
    pol_list.append(data1['POLICY_NO'])
    date_eff = datetime.datetime.strptime(str(data1['DATE_EFFECTIVE']),"%d-%m-%y").date()
    if data1['CIA_REASN_CD'] == 'MOR':
        counting += 1
    total_count += 1
    if data1['FUND_NAME'].strip() not in fund_list:
        fund_list.append(data1['FUND_NAME'].strip())
policy_list = list(set(pol_list))
count_fund = len(fund_list)
counting = int(counting/count_fund)

for i in file_data3:
        charges = 0.0
        gst = 0.0
        a = 0.025
        total_amt = 0.0
        flag = 0
        f_units = 0.0
        mode_val = 0.0
        rate_val = 0
        Mode = ''
        if i['POL_BILL_TYP_CD']=='P':
            Mode = 'Single'
        else:
            Mode = 'Regular'
                                    
        Pol = i['POL_ID']
        prop = str(Pol)[:2]
        policy_term = i['POLICY_TERM']
        
        uin = i['UIN']
        try:
            effc_date =  datetime.datetime.strptime(str(i['CIA_EFF_DATE']),"%d-%m-%y").date()
        except ValueError:
            effc_date =  datetime.datetime.strptime(str(i['CIA_EFF_DATE']),"%d-%m-%y").date()
            
        try:
            doc =  datetime.datetime.strptime(str(i['DOC']),"%d-%m-%y").date()
        except ValueError:
            doc =  datetime.datetime.strptime(str(i['DOC']),"%d-%m-%y").date()


        policy_yr = effc_date - doc
        poli_yr = relativedelta(effc_date,doc).years
        policy_year = relativedelta(effc_date,doc).years
        pol_yr = policy_yr.days
        pol_year_effc_date = effc_date.year
        pol_year_doc = doc.year
        pol_month_effc_date = effc_date.month
        pol_month_doc = doc.month
        pol_Days_effc_date = effc_date.day
        pol_Days_doc = doc.day
        ceil_pol = pol_yr/365.25
        policy_year_round = math.ceil(ceil_pol)
    
        if policy_year>=0 and pol_month_doc==pol_month_effc_date and pol_Days_effc_date==pol_Days_doc:
            policy_year += 1
        else:
            policy_year=math.ceil(ceil_pol)

        poli_yr = (poli_yr) + 1

            
        basic_premium = i['PREMIUM']

        Policy = str(Pol)[0:2].strip()

        for j in file_ratesheet:
            for key,value in j.items():
            
                if Policy == str(j['PROPOSAL_NO']) and poli_yr == j['Policy _Year'] and Mode==j['MODE'] and i['DISTRIBUTION_CHANNEL']=='O' and Policy==('2J' or '2H'):
                    if i['STAFF_DISCOUNT_FLAG'] == key:
                        additional_rate = abs(value - a)
                        charges = i['PREMIUM'] * additional_rate
                        pol_yrr = (j['Policy _Year'])
                        mode_val = (j['MODE'])
                        rate_val = (value)
                        if i['POL_ISS_LOC_CD'] == 15 and i['GST_REGION']==15:
                            total_amt = charges * 1.19
                        else:
                            total_amt = charges * 1.18
                else:
                    if Policy == str(j['PROPOSAL_NO']) and poli_yr == j['Policy _Year'] and Mode==j['MODE']:
                        if i['STAFF_DISCOUNT_FLAG'] == key:
                            rate_val = (value)
                            pol_yrr = (j['Policy _Year'])
                            mode_val = (j['MODE'])
                            charges = i['PREMIUM'] * value
                            if i['POL_ISS_LOC_CD'] == 15 and i['GST_REGION']==15:
                                total_amt = charges * 1.19
                            else:
                                total_amt = charges * 1.18
                        
        for m in file_1H:
                for key,value in m.items():
                    if Policy == str(m['PROPOSAL_NO']) and poli_yr == m['Policy _Year'] and Mode==m['MODE'] and uin == m['UIN']:
                        if i['STAFF_DISCOUNT_FLAG']==key and  policy_term==m['Policy_Term']:
                            charges = i['PREMIUM'] *  value
                            if i['POL_ISS_LOC_CD'] == 15 and i['GST_REGION']==15:
                                total_amt = charges * 1.19
                            else:
                                total_amt = charges * 1.18
        for k in file_oldprod:
            for key,value in k.items():
                Prop_no = int(k['PROPOSAL_NO'])
                Pol_yr = int(k['POLICY_YEAR'])
                if  Policy == str(Prop_no)  and str(poli_yr) == str(Pol_yr):
                    pol_yrr = int(k['POLICY_YEAR'])
                    mode_val = ''
                    rate_val = k['RATE']
                    charges = k['RATE'] * basic_premium
                    if i['POL_ISS_LOC_CD'] == 15 and i['GST_REGION']==15:
                        total_amt = charges * 1.19
                    else:
                        total_amt = charges * 1.18

        policy_year_list.append(poli_yr)
        mode_list.append(mode_val)
        rate_list.append(rate_val)

        alloc_with_gst = i['PREMIUM'] - total_amt
        alloc_without_gst = i['ALLOCATION_WITHOUT_GST'] - charges
        alloc_with_gst = i['PREMIUM'] - total_amt
        allocation_with_gst = i['INVESTABLE_AMT_POST_ALLOCATION'] - alloc_with_gst
        sys_units = i['NO_OF_UNITS_TRANSACTED']
        sys_perct.append(i['CDI_ALLOC_PCT'])

        cal_chrg = round(total_amt,2)-abs(i['CHARGES'])
        pending_amt = i['PREMIUM'] - i['INITIAL_AMOUNT']

        nom = str(i['FUND_NAME']).strip()
        fund_name2 = str(i['FUND_NAME'])
        try:
            dom_date2 = datetime.datetime.strptime(str(i['CIA_EFF_DATE']),"%d-%m-%y").date()
        except ValueError:
            dom_date2 = datetime.datetime.strptime(str(i['CIA_EFF_DATE']),"%d-%m-%y").date()

        fv_cal = round(float(i['SYS_NAV']) * float(i['NO_OF_UNITS_TRANSACTED']),4)
        fv_pct = round((fv_cal / float(i['INITIAL_AMOUNT']))*100,2)
        calculated_fv.append(fv_cal)
        calculated_fnd_perct.append(fv_pct)
        
        amt_w_o_gst = round(total_amt,2) - round(abs(i['CHARGES']-total_amt),4)
        chrg_diff = i['CHARGES'] - amt_w_o_gst
        #units_diff = sys_units - f_units
        units_diff = fv_pct - i['CDI_ALLOC_PCT']
        units_diff_list.append(units_diff)

        if chrg_diff > -1 and chrg_diff < 1:
            result_list2.append('PASS')
        else:
            result_list2.append('FAIL')
            
        if units_diff > -1 and units_diff < 1:
            result_list3.append('PASS')
        else:
            result_list3.append('FAIL')

        proposal_no_list2.append(i['POL_ID'])
        doc_list2.append(i['DOC'])
        premium_list2.append(i['PREMIUM'])
        charges2.append(i['CHARGES'])
        staff_discount_flag_list.append(i['STAFF_DISCOUNT_FLAG'])
        CIA_EFF_DT2.append(i['CIA_EFF_DATE'])
        risk_premium_final2.append(amt_w_o_gst)
        date_monthly2.append(i['MONTHIWERSARY'])
        reason_code2.append(i['CIA_REASN_CD'])
        DIFFERENCE2.append(chrg_diff)

start=0
start_alloc=0
end_alloc=count_fund
end=count_fund*2
for z in range(counting):
    fund_percentage.extend(fund_percentage2[start_alloc:end_alloc]+fund_percentage1[start:end])
    CAL_UNITS.extend(CAL_UNITS2[start_alloc:end_alloc]+CAL_UNITS1[start:end])
    proposal_no_list.extend(proposal_no_list2[start_alloc:end_alloc]+proposal_no_list1[start:end])
    doc_list.extend(doc_list2[start_alloc:end_alloc]+doc_list1[start:end])
    premium_list.extend(premium_list2[start_alloc:end_alloc]+premium_list1[start:end])
    CHARGES.extend(charges2[start_alloc:end_alloc]+charges1[start:end])
    CIA_EFF_DT.extend(CIA_EFF_DT2[start_alloc:end_alloc]+CIA_EFF_DT1[start:end])
    risk_premium_final.extend(risk_premium_final2[start_alloc:end_alloc]+risk_premium_final1[start:end])
    date_monthly.extend(date_monthly2[start_alloc:end_alloc]+date_monthly1[start:end])
    reason_code.extend(reason_code2[start_alloc:end_alloc]+reason_code1[start:end])
    DIFFERENCE.extend(DIFFERENCE2[start_alloc:end_alloc]+DIFFERENCE1[start:end])
    fnd_name.extend(fnd_name2[start_alloc:end_alloc]+fnd_name1[start:end])
    value_nav.extend(value_nav2[start_alloc:end_alloc]+value_nav1[start:end])
    fund_units.extend(fund_units2[start_alloc:end_alloc]+fund_units1[start:end])
    fund_v0.extend(fund_v2[start_alloc:end_alloc]+fund_v1[start:end])
    result_list.extend(result_list2[start_alloc:end_alloc]+result_list1[start:end])
fund_percentage
proposal_no_list
doc_list
premium_list
CHARGES
CIA_EFF_DT
risk_premium_final
date_monthly


start=end
start_alloc=end_alloc
end+=count_fund*2
end_alloc+=count_fund



yest_date = datetime.datetime.now().date()-timedelta(days=1)


c0 = pd.Series(policy_no_list, name = "POLICY_NO")
c1 = pd.Series(doc_list2, name = "DOC")
c3 = pd.Series(age_list, name = "AGE")
c4 = pd.Series(sum_assured_list, name = "SUM_ASSURED")
c6 = pd.Series(premium_list2, name = "PREMIUM")
c7 = pd.Series(proposal_no_list2, name = "PROPOSAL_NO")
c8 = pd.Series(charges2, name = "SYS_CHARGES")
#c10 = pd.Series(CIA_EFF_DT2, name = "MONTHIWERSARY") # date_monthly2
#c10 = pd.Series(cal_date_ls, name = "CALCULATED_DATE")
c32 = pd.Series(staff_discount_flag_list, name = "STAFF_DISCOUNT_FLAG")
c20 = pd.Series(fnd_name2, name = "FUND_NAME")
c21 = pd.Series(calculated_fnd_perct, name = "CAL_PERCENT") #fund_percentage2
c18 = pd.Series(fund_units2, name = "FUND_UNITS") #calculated_fv
c19 = pd.Series(value_nav2, name = "FUND_NAV")
c22 = pd.Series(calculated_fv, name = "CALCULATED_FV") #fund_v2
c12 = pd.Series(fund_details_list, name = "FUND_DETAILS")
c13 = pd.Series(sar_final_value, name = "SAR")
c29 = pd.Series(policy_year_list, name = "POLICY_YEAR")
c30 = pd.Series(mode_list, name = "MODE")
c31 = pd.Series(rate_list, name = "RATE")
c14 = pd.Series(risk_premium_final2 , name = "CAL_CHARGES")
c17 = pd.Series(result_list2 , name = "CHARGES_RESULT")
c23 = pd.Series(reason_code2 , name = "REASON_CODE")
c24 = pd.Series(CIA_TYP_CD , name = "CIA_TYP_CD")
c25 = pd.Series(CIA_EFF_DT2 , name = "CIA_EFF_DT")
c26 = pd.Series(DIFFERENCE2 , name = "CHARGES_DIFFERENCE")
c27 = pd.Series(sys_perct , name = "SYS_PERCT") #CAL_UNITS2
c28 = pd.Series(result_list3 , name = "PERCT_RESULT") 

data_concat = [c7,c1,c25,c29,c30,c31,c32,c23,c6,c20,c8,c14,c26,c17,c22,c27,c21,c28]
df = pd.concat(data_concat, axis=1)
print('End time:',datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
print("---Total Execution time: %s seconds ---" % (time.time() - start_time))
print("-----------------------------------------------------")
print('Succes , choose path to save')

path_to_save = filedialog.asksaveasfilename(filetypes = (("csv files", "*.csv"),("all files","*.*")))
df.to_csv(path_to_save, index=False)

print("csv file is generated at path ")
print(path_to_save)
