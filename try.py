#import raw
import plotly.express as px
import pandas as pd
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

#import stock data
import yfinance as yf

#make sample set of SET50 to exclude the range
SET50 = {'ADVANC','AWC','BANPU','BBL','BDMS','BEM','BGRIM','BH','BLA','BTS','CBG','CPALL',
         'CPF','CPN','CRC','DTAC','EA','EGCO','GLOBAL','GPSC','GULF','HMPRO','INTUCH','IRPC','IVL',
         'JMART','JMT','KBANK','KCE','KTB','KTC','LH','MINT','MTC','OR','OSP','PTT','PTTEP','PTTGC',
         'SAWAD','SCB','SCC','SCGP','TIDLOR','TISCO','TOP','TRUE','TTB','TU'}

for eachstock in SET50:
    stock_name = eachstock + '.' + 'BK'
    #because in yahoo_finance, it has to be with .bk because it's not in US
class STOCK():
    def __init__(self,stock):
        self.stock = stock
        
    def getsettable(self):
        '''making table that tell about the price of each stock including open close adjusted close etc'''
        if self.stock in SET50:
            datastock = yf.download(tickers= self.stock, period='30d', interval='1d')
            # creating excel writer object
            tableset = pd.ExcelWriter('set50_table.xlsx')
            datastock.to_excel(tableset)
            tableset.save()
            datastock.to_excel("set50_table.xlsx", sheet_name= self.stock + '.BK')
            os.system("start EXCEL.EXE set50_table.xlsx")
            return datastock
        


    def getgraph(self):
        '''making graph including with time range by using plotly'''
        if self.stock in SET50:
            datastock2 = yf.download(tickers= self.stock)
            datastock2["Datetime"] = datastock2.index
            datastock2 = datastock2[["Datetime", "Open", "High", "Low", 
                    "Close", "Adj Close", "Volume"]]
            graph = px.line(datastock2, x='Datetime', y='Adj Close',  title = self.stock + ' ' + 'price graph with the time period selectors')

            graph.update_xaxes(
                rangeselector=dict(
                    buttons=list([
                        dict(count=5, label="5d", step="day", stepmode="backward"),
                        dict(count=1, label="1m", step="month", stepmode="backward"),
                        dict(count=3, label="3m", step="month", stepmode="backward"),
                        dict(count=6, label="6m", step="month", stepmode="backward"),
                        dict(count=1, label="1y", step="year", stepmode="backward"),
                        dict(step="all")
                    ])
                )
            )
            graph.show()


counter = 0
while counter < 2:
    stock_name = input('Enter a stock name: ')
    if stock_name not in SET50:
        print('Sorry, this program is for Thai stocks only. Please try again.')
        counter = 0
    else:
        counter += 1
        input_for_graph = input()
        if input_for_graph == 'open graph':
            STOCK(stock_name)
        counter += 1