import quandl
import pandas as pd
import numpy as np
import xlsxwriter 
import xlrd
import urllib2 
import re
import time
import string
from quandl.errors.quandl_error import (
    QuandlError, LimitExceededError, InternalServerError,
    AuthenticationError, ForbiddenError, InvalidRequestError,
    NotFoundError, ServiceUnavailableError)

quandl.ApiConfig.api_key = "oXecxhUApx-7fSV1DqiR"


''' #read from quandl dataset and write into Excel

df = quandl.get("NSE/UJJIVAN", returns="pandas")
writer = pd.ExcelWriter('UJJIVAN.xlsx', engine="xlsxwriter")
df.to_excel(writer, sheet_name='Sheet1')
writer.save()
print "done"
'''

'''
xl_file = pd.read_excel("/Volumes/Elements/Trading/Automation/NSEStocks/NSEStocksList.xlsx", sheetname=None)

#for (k,v) in xl_file.items():
#	print v

print xl_file.items()[0]
'''

#read from quandl for all stocks 

def fetchId(j):

	book = xlrd.open_workbook("/Volumes/Elements/Trading/Automation/NSEStocks/NSEStocks"+str(j)+".xlsx")

	first_sheet = book.sheet_by_index(0)
	NSEStocksList = first_sheet.col(0)
	finalSheetFrame = pd.DataFrame(columns=['StockName', 'CMP', '50D/200D Crossover', '21D/10D Turtle', '21D High', '10D Low', '55D/21D Turtle', '55D High', '21D Low'])
	for i in range (1, len(NSEStocksList)):
		print NSEStocksList[i].value
		try:
			df = quandl.get("NSE/"+NSEStocksList[i].value, returns="pandas")
			stockName = NSEStocksList[i].value
			#Moving average crossover strategy
			#print "50D/200D crossover"
			df['50-SMA'] = df['Close'].rolling(window=50).mean()
			df['200-SMA'] = df['Close'].rolling(window=200).mean()
			df['5-EMA'] = df['Close'].ewm(span=5, min_periods=6, adjust=False).mean()
			
			df['Buy/Sell'] = np.where((df['50-SMA'] > df['200-SMA']), 'BUY', 'SELL')
			size = df['Buy/Sell'].size
			maCrossOver = ""
			if(df['Buy/Sell'][size-1] != df['Buy/Sell'][size-2]):
				maCrossOver = " Fresh signal " + df['Buy/Sell'][size-1]
			else:
				maCrossOver = " Maintain " + df['Buy/Sell'][size-1]

			#Turtle Trading strategy
			df['20DHigh'] = df['High'].rolling(window=20, min_periods=20).max()
			df['55DHigh'] = df['High'].rolling(window=55, min_periods=55).max()
			df['10DLow'] = df['Low'].rolling(window=10, min_periods=10).min()
			df['21DLow'] = df['Low'].rolling(window=21, min_periods=21).min()

			df['20DBuy'] = np.where((df['High'] >= df['20DHigh']), 'BUY', '0')
			df['20DSell'] = np.where((df['Low'] <= df['10DLow']), 'SELL', '0')
			df['55DBuy'] = np.where((df['High'] >= df['55DHigh']), 'BUY', '0')
			df['55DSell'] = np.where((df['Low'] <= df['21DLow']), 'SELL', '0')

			twentyDayBuySize = df['20DBuy'].size
			twentyDaySellSize = df['20DSell'].size
			fiftyFiveDayBuySize = df['55DBuy'].size
			fiftyFiveDaySellSize = df['55DSell'].size

			turtleStrategy1 = ""
			last20DHigh = ""
			last10DLow = ""
			last55DHigh = ""
			last21DLow = ""
			closePrice = df['Close'][twentyDaySellSize-1]
			if(df['20DBuy'][twentyDayBuySize-1] == "Buy"):
				turtleStrategy1 = "Start Buying - Crossed 21 day high"
			elif(df['20DSell'][twentyDaySellSize-1] == "Sell"):
				turtleStrategy1 = "Sell it all - Crossed 10 day low"
			else:
				lastHigh = df['High'][twentyDaySellSize-1]
				lastLow = df['Low'][twentyDaySellSize-1]
				last20DHigh = df['20DHigh'][twentyDaySellSize-1]
				last10DLow = df['10DLow'][twentyDaySellSize-1]
				if(last20DHigh-lastHigh <= lastLow-last10DLow):
					turtleStrategy1 = "Close to 20D High"
				else:
					turtleStrategy1 = "Close to 10D Low"
			turtleStrategy2 = ""
			if(df['55DBuy'][fiftyFiveDayBuySize-1] == "Buy"):
				turtleStrategy2 = "Start Buying - Crossed 55 day high yesterday"
			elif(df['55DSell'][fiftyFiveDaySellSize-1] == "Sell"):
				turtleStrategy1 = "Sell it all - crossed down 21 day low"
			else:
				lastHigh = df['High'][fiftyFiveDaySellSize-1]
				lastLow = df['Low'][fiftyFiveDaySellSize-1]
				last55DHigh = df['55DHigh'][fiftyFiveDaySellSize-1]
				last21DLow = df['21DLow'][fiftyFiveDaySellSize-1]
				if(last55DHigh - lastHigh <= lastLow - last21DLow):
					turtleStrategy2 = "Close to 55D High"
				else:
					turtleStrategy2 = "Close to 21D Low"

			finalSheetFrame.loc[i] = [stockName, closePrice, maCrossOver, turtleStrategy1, last20DHigh, last10DLow, turtleStrategy2, last55DHigh, last21DLow]
			#for value in df['200-SMA']:
			#	print value
			#print df
			writer = pd.ExcelWriter(NSEStocksList[i].value+".xlsx", engine="xlsxwriter")
			df.to_excel(writer, sheet_name='Sheet1')
			writer.save()

		except NotFoundError:
			print NSEStocksList[i].value + "Not Found"
		except QuandlError:
			print "Quandl Error"
	writer = pd.ExcelWriter("CallsForTheDay"+str(j)+".xlsx", engine="xlsxwriter")
	finalSheetFrame.to_excel(writer, sheet_name="Sheet1")
	writer.save()

fetchId(1)
print "First part over"
fetchId(2)
print "Second part over"
fetchId(3)
print "Third part over"
fetchId(4)
print "Fourth part over"
fetchId(5)
print "Fifth part over"
fetchId(6)
print "Complete"