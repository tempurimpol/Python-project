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
SET50 = {'BGRIM.BK','OSP.BK','CBG.BK','KBANK.BK','ADVANC.BK','TU.BK','CPN.BK','KTC.BK'
        ,'SCGP.BK','EGCO.BK','AOT.BK','CPF.BK','SCB.BK','KCE.BK','EA.BK','LH.BK','PTTGC.BK'
        ,'CRC.BK','JMART.BK','IVL.BK','IRPC.BK','BBL.BK','CPALL.BK','MINT.BK','GULF.BK','TIDLOR.BK'
        ,'PTTEP.BK','BH.BK','BLA.BK','DTAC.BK','OR.BK'}
    #because in yahoo_finance, it has to be with .bk because it's not in US

#define functions for making statistic table and graph        
def getsettable(stock):
    '''making table that tell about the price of each stock including open close adjusted close etc'''
    if stock in SET50:
        datastock = yf.download(tickers= stock, period='30d', interval='1d')
        # creating excel writer object
        tableset = pd.ExcelWriter('set50_table.xlsx')
        datastock.to_excel(tableset)
        tableset.save()
        datastock.to_excel("set50_table.xlsx", sheet_name= stock)
        return datastock
        


def getgraph(stock):
    '''making graph including with time range by using plotly'''
    if stock in SET50:
        datastock2 = yf.download(tickers= stock)
        datastock2["Datetime"] = datastock2.index
        datastock2 = datastock2[["Datetime", "Open", "High", "Low", 
                    "Close", "Adj Close", "Volume"]]
        graph = px.line(datastock2, x='Datetime', y='Volume',  title = stock + ' ' + 'price graph with the time period selectors')

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
        if input_for_graph == 'make data':
            getsettable(stock_name)
            getgraph(stock_name)
        counter += 1

#for open table via excel
os.system("start EXCEL.EXE set50_table.xlsx")

# Create an instance of tkinter frame
win = Tk()

# Set the size of the tkinter window
win.geometry("3000x350")

# Create a Frame
frame = Frame(win)
frame.pack(pady=20)

# Create an object of Style widget
style = ttk.Style()
style.theme_use('clam')

# Create a Treeview widget
tree = ttk.Treeview(frame)

# Define a function for opening the file
def open_file():
   filename = filedialog.askopenfilename(title="Open a File", filetype=(("xlxs files", ".*xlsx"),))

   if filename:
      try:
         filename = r"{}".format(filename)
         dataframe = pd.read_excel(filename)
      except ValueError:
         label.config(text="File could not be opened")
      except FileNotFoundError:
         label.config(text="File Not Found")


   # Add new data in Treeview widget
   tree["column"] = list(dataframe.columns)
   tree["show"] = "headings"

   # For Headings iterate over the columns
   for col in tree["column"]:
      tree.heading(col, text=col)

   # Put Data in Rows
   dataframe_rows = dataframe.to_numpy().tolist()
   for row in dataframe_rows:
        tree.insert("", "end", values=row)

   tree.pack()


# Add a Menu
m = Menu(win)
win.config(menu=m)

# Add Menu Dropdown
file_menu = Menu(m, tearoff=False)
m.add_cascade(label="Menu", menu=file_menu)
file_menu.add_command(label="Open data to link", command=open_file)

# Add a Label widget to display the file content
label = Label(win, text='')
label.pack(pady=20)

win.mainloop()
