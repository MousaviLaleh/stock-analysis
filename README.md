# stock-analysis

## Overview of Project
Steve wnats to research about Green Energy stocks and finds out which company has the best performance, to invest in, specifically DAQO New Energy Corporation, a company that makes silicon wafers for solar panels.<br/>
I will be using VBA ( Visual Basic for Appllication ) throughout this project, to automate tasks and reus codes for any stock. VBA reduces the chance of accidents and errors and increases the code running time.<br/>

### Purpose
Steve wants to to expand the dataset to include the entire stock market. I have prepared an Excel file containing the stock data over the last few years. The purpose is improving the logic of the VBA code in order to make it more efficient to works well for thousands of stocks. At the click of a button, Steve can analyze an entire dataset.<br/>


## Code Review
To have an automated code, we need to use an InputBox to get the choice of the year <br/>
![inputBox.png](/resources/inputBox.png)

1. Source Sheets:  are the 2017, and 2018  sheets; which contain the data of 11 different stock company for the years of 2017 and 2018
    we use this code to active the output data sheet <br/>
3. output Sheet: All Stocks Analysis
    to active the output data sheet, we write this code
    Worksheets("All Stocks Analysis").Activate br/>
    
    
5. Arrays that save the result for every ticker: 
   - tickers : an array to save the name of tickers
   - tickerVolumes : an array to save the total revenue volume of each ticker
   - tickerStartingPrices : an array to save the first price of each ticker
   - tickerEndingPrices : an array to save the last price of each ticker
6. startTime and  endTime : two Single variables to save the code running time
7. first For loop ( tickerIndex = 0 To 11 )  : to loop over all the rows in the spreadsheet
8. second For loop ( For i = 2 To RowCount ) : to loop over each ticker data, and calculate the total volume of, startinPrice and endingPrice of each ticker
9. third For loop ( For i = 0 To 11 ) : to loop through four arrays to output the Ticker name, Total Daily Volume, and Return
10. formatting section 


## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## Summary
address the following questions:
  What are the advantages or disadvantages of refactoring code?
  How do these pros and cons apply to refactoring the original VBA script?
