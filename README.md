# VBA-challenge

Jeff Pinegar
Jeffpinegar1@gamil.com
+1-717-982-0516

# Readme file for the Wallstreet Challenge

## Submission Note
The maximum file size for my GitHub is 50MB
The Multi_year_stock_data.xlsb file is 56MB so I could not load it.
Therefore I split the output sheet out and named it: Wallstreet Results Output.xlsx 

_____________________________________________
### Assumptions:
* The source workbook only contains worksheets with ticker data
* Each sheet starts in cell A1
* Each worksheet has these columns, in this order <ticker> <date> <open> <high> <low> <close> <vol>
* The data starts in cell A2
* I place all the results in a sperate worksheet named outputs.  The first column of the output table will include the name of the source worksheet.   To see an individual worksheet result the table can be filtered using the filters that are presented.

_____________________________________________
### NOT Assume in my project:
* The number of worksheets of data in the source workbook, the only limit is n-1 where is Excel's limit.
* The worksheets are not in order by ticker or date. I sort each sheet by ticker and date before I process.
* Each worksheet has an unknown number of rows.  I process down the data till I reach a blank.  There is no risk of a mid table blanks since I sort the table to start.

_____________________________________________
### How to run this project
1.  Open the worksheets containing the macros
2.  Open the file containing the Data
3.  Run the macro "Main_Code"
4.  All the results will be on the first worksheet "Output"

### Execution times
*   Alphabetical_testing.xlsx -- 20 seconds
*   Multiple_year_stock_data.xlsx -- 15 minutes

### output Images
*  Greatest Changes per Year.jpg
*  2018 Results.jpg -- result for 2018
*  2019 Results.jpg -- result for 2019
*  2020 Results.jpg -- result for 2020

____________________________________________
# Macros contained in the project
### Main_Code.vbs
* This is where the project starts, run main_code()
* calls to outputsheet()
* calls TickerTotal() once for each sheet 
* after all the sheets have been processed calls GrtSummary() to identify the tickers with the greats % gain, % loss, and volume
* Calls OutputHeadings() to format the output sheet

### TickerTotal.vbs 
* Once a new ticker is identified, calculate the yearly change, %change, and total volume
* When a ticker is totaled call writeTicker() to record results in the output table

### WriteTickerOutput.vbs
* Once a Ticker is totaled (TickerTotal.vbs) This macro writes the results in the output table.

### Outputsheet.vbs
* This file prepares a worksheet for recording the results.

