#web parser
import bs4 as bs
#save/load python objects, including ml models or pd dataframes
import pickle
#web info requestor
import requests
import pandas as pd
#retrieve yahoo finance data
import pandas_datareader.data as web
import numpy as np
import datetime as dt
import os
import time
import matplotlib.pyplot as plt
import warnings
import re



def get_data(url,tickers):
    df=pd.DataFrame()
    for ticker in tickers['ticker']:
        resp=requests.get(url.format(ticker))
        sauce=bs.BeautifulSoup(resp.text,'xml')
        dict={}
        dict['ticker']=ticker
        ind=list(tickers['ticker']).index(ticker)
        dict['code']=tickers.loc[ind,'code']
        dict['name']=tickers.loc[ind,'name']
        print('Start pulling data from MSN Money for: ','\n',tickers.loc[ind,'name'])
        # notice that each statistic is wrapped up in one 'ul';under each 'ul', the statistic's name and value are then wrapped up in separate 'p' 
        for ls in sauce.select('ul'):
            para=ls.select('p')    
            if len(para)!=0:
                # when there are 3 'p' in a 'ul',it's because there is an additional note to the name of the statistic
                if len(para)==3:
                    dict[para[0].text+para[1].text]=para[2].text
                # some 'ul' have 4 'p' in it because the statistics also display last period's value
                elif len(para)==4:
                    dict[para[0].text]=para[2].text
                    dict[para[0].text+para[1].text]=para[3].text
                else:
                    dict[para[0].text]=para[1].text
        df=df.append(dict,ignore_index=True)
        pct=(ind+1)/len(tickers)*100
        print(f'{pct:.2f}% completed','\n','###########################################')
        time.sleep(20)
    df['Date']=dt.date.today()
    return df
 

url_to_read='https://www.msn.com/en-us/money/stockdetails/analysis/fi-{}'
df_tickers=pd.read_excel('ticker_list_nz.xlsx')


st=dt.datetime.now()
df_data=get_data(url_to_read,df_tickers)
et=dt.datetime.now()
print('Time spent: ',et-st)
time.sleep(10)

df_data.to_excel('stock_data_nz/stock_stats_nz_'+str(dt.date.today())+'.xlsx',index=False)