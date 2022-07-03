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
filename =  filedialog.askopenfilename(initialdir = "/",title = "Select annuity file",
#filetypes = (("csv","*.csv"),("all files","*.*")))
filetypes = (("xlsx","*.xlsx"),("all files","*.*")))
print(filename)
sheet = pd.read_excel(filename) 
df1 = pd.DataFrame(sheet)
dict1 = df1.to_dict(orient='records')

#print('dict4 >>>',dict1)
print(len(dict1))



if __name__ == '__main__':

	plan_id = []
	plan_irda_uin_id = []
	
	primary_annuitant_age = []
	secondary_annuitant_age = []
	annuity_option = []
	annuity_rate = []
	
	#PLAN_IRDA_UIN_ID	
	version = []
	logic_flag = []
	product_name=[]
     

i=0
while i < len(dict1):
	l = dict1[i]
	
	#if i < 3:
	#print('dict1[i] >>', l)		
	#print('i >>>',i)
	
	for key,value in l.items():
		annuity_option_ = 2.3
		#print('key >>',key)
		#print('value >>',value)
		plan_id.append('AN0072')
		plan_irda_uin_id.append('111N083V04')
		primary_annuitant_age.append(i)
		secondary_annuitant_age.append(key)
		annuity_option.append(annuity_option_)
		annuity_rate.append(value)
		version.append('V4')
		logic_flag.append('2.3')
		product_name.append('SBIL - Life and Last survivor-100% Income - Series 4')

	i=i+1
	
o1 = pd.Series(plan_id,dtype=str,name ='PLAN_ID')
o2 = pd.Series(plan_irda_uin_id,dtype=str,name ='PLAN_IRDA_UIN_ID')
o3 = pd.Series(primary_annuitant_age,dtype=str,name ='PRIMARY_ANNUITANT_AGE')
o4 = pd.Series(secondary_annuitant_age,dtype=str,name='SECONDARY_ANNUITANT_AGE')
o5 = pd.Series(annuity_option,dtype=str,name='ANNUITY_OPTION')
o6 = pd.Series(annuity_rate,dtype=str,name='ANNUITY_RATE')
o7 = pd.Series(version,dtype=str,name='VERSION')
o8 = pd.Series(logic_flag,dtype=str,name='LOGIC_FLAG')
o9= pd.Series(product_name,dtype=str,name='PRODUCT_NAME')

data_concat = [o1,o2,o3,o4,o5,o6,o7,o8,o9]

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
