#!/usr/bin/env python
# coding: utf-8

# In[7]:


import urllib.request
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
url = 'https://api.hkma.gov.hk/public/market-data-and-statistics/monthly-statistical-bulletin/er-ir/er-eeri-endperiod?offset=0'

with urllib.request.urlopen (url) as req:
    print (req.read())


# In[64]:


from tkinter import *
from xlrd import open_workbook
import time
import urllib.request

url = "https://www.hkma.gov.hk/media/eng/doc/market-data-and-statistics/monthly-statistical-bulletin/C04.xls"
urllib.request.urlretrieve(url, "C04.xls")
#香港金融管理局 Hong Kong Monetary Authority

def hk_exchange_rate():
    pass_time = df.get()
    wb = open_workbook('C04.xls')
    sheet_names = wb.sheet_names()
    s = wb.sheet_by_name(sheet_names[0])
    #print 'Sheet:',s.name
    values = []
    for row in range(s.nrows):
        if s.cell(row,1).value != '':
            tm = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime((s.cell(row,1).value-25569) * 60 * 60 * 24))
            value = (tm, s.cell(row,3).value)
            if(tm[0:7] == pass_time):
                text.insert(END, value)
        

def main():
    global df,text
    
    background = Tk() 
    
    background.title('港幣歷年當月匯率：US$/HK$每美元兌港元') 
    
    background.geometry('550x350') 
    
    Label(background,text='請輸入例如：「2004-08」:',font=("微軟雅黑",20),fg='blue').grid()
    
    df=Entry(background,font=("微軟雅黑",15))
    
    df.grid(row=0,column=1) 
    
    text=Listbox(background,font=('微軟雅黑',15),width=45,height=10)
    
    text.grid(row=1,columnspan=2) 
    
    Button(background,text='點我下載',font=("微軟雅黑",15),command=hk_exchange_rate).grid(row=2,column=0,sticky=W)
    
    Button(background,text='離開',font=("微軟雅黑",15),command=background.destroy).grid(row=2,column=1,sticky=E)
    
    mainloop() 


# In[65]:


main() 


# In[48]:


import pandas as pd
import numpy as np
file_loc = "C04.xls"
df = pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols = "B,D")
data={'date':[pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols = "B")], 'exchange_rate':pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols = "D")}

data1=pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols = "B")
data2=pd.read_excel(file_loc, index_col=None, na_values=['NA'], usecols = "D")

df1 = pd.DataFrame(data, columns=['date','exchange_rate'])

x_date=data1[2:211]
y_exchange_rate=data2[2:211]


# In[62]:


import matplotlib.pyplot as plt
import pandas as pd
x_values=[1,2,3]
y_values=y_exchange_rate
plt.plot(x_values, y_values, color='b')
plt.xlabel('date') # 設定x軸標題
plt.xticks(x_values, rotation='vertical') # 設定x軸label以及垂直顯示
plt.title('港幣歷年當月匯率：US$/HK$每美元兌港元') # 設定圖表標題
plt.show()


# In[59]:


dates = ["01/15/2004", "01/15/2012", "01/15/2021"]

x_values = [dates.datetime.strptime(d,"%m/%d/%Y").date() for d in dates]
y_values = y_exchange_rate

ax = plt.gca()

formatter = mdates.DateFormatter("%Y-%m-%d")
ax.xaxis.set_major_formatter(formatter)

locator = mdates.DayLocator()
ax.xaxis.set_major_locator(locator)

plt.plot(x_values, y_values)


# In[51]:


x_date


# In[63]:


y_exchange_rate


# In[ ]:


import matplotlib.pyplot as plt
import pandas as pd

plt.plot(x_date, y_exchange_rate, color='b')
plt.xlabel('date') # 設定x軸標題
plt.xticks(x_date, rotation='vertical') # 設定x軸label以及垂直顯示
plt.title('港幣歷年當月匯率：US$/HK$每美元兌港元') # 設定圖表標題
plt.show()

