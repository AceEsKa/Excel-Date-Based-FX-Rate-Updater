# Excel-Date-Based-FX-Rate-Updater
Specificaly app made for small a small auditor firm
ECB foreign exchange reference rates finder and appender for excel
Takes the excel file name
Takes column with dates assuming those a transaction days, days on which transaction happend
takes the starting row 
then downloads xml file form nbs https://nbs.sk/export/sk/exchange-rate/YYYY-MM-DD/xml
finds the exchange coressponding to the currency we want after that all the echange rates for all of our transactions a are written back to the excel spreadsheet in the column we want
