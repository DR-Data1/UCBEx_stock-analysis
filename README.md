# UCBEx_stock-analysis
Module 2: Analysis using VBA 
# Challenge

The file (./green_stocks_DR-Challenge.xlsm) contains a macro that analyzes all stock data returning total volume and return for the year.

The macro (AllStocksAnalysis) will create title, headers and summary data for each stock ticker. 
First, it request the user to enter a YYYY year.
Second, it creates an array of ticker symbols
Third, it creates empty arrays to hold the main variables of interest (i.e., total volume, starting price, and ending price)
Next, the program loops through each row of data.  The program then checks to see what ticker value (via index) matches the value in first cell of the record  
If the values match, then three conditional statements are evaluated corresponding to each variable.  If conditions are satisfied the value of the variables are modified

Next loop, writes the in appropriate worksheet.
THen data is formatted.
