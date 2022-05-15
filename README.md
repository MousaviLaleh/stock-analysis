# stock-analysis

## Overview of Project
Steve wnats to research about Green Energy stocks and finds out which company has the best performance, to invest in, specifically DAQO New Energy Corporation, a company that makes silicon wafers for solar panels.<br/>
I will be using VBA ( Visual Basic for Appllication ) throughout this project, to automate tasks and reus codes for any stock. VBA reduces the chance of accidents and errors and increases the code running time.<br/>

### Purpose
Steve wants to to expand the dataset to include the entire stock market. I have prepared an Excel file containing the stock data over the last few years. The purpose is improving the logic of the VBA code in order to make it more efficient to works well for thousands of stocks. At the click of a button, Steve can analyze an entire dataset.<br/>


## Code Review
The code should do the following:
- Set the startTime and endTime variables to measure and show the code running time.<br/>
    ![times.png](/resources/times.png)<br/>
    note that the startTime must be call right after inputBox to calculate the time from this point of progress. <br/>

- To have an automated progress, we need to use an InputBox to get the choice of the year from user. <br/>
![inputBox.png](/resources/inputBox.png)<br/>
![msgBox.png](/resources/msgBox.png)<br/>

- Format output Sheet: is the sheet that we show the results in it. In this project "All Stocks Analysis" is the output sheet.<br/>
    to format the output data sheet: <br/>
    Worksheets("All Stocks Analysis").Activate <br/>
       
- Initialize arrays that save the result for every ticker: 
   - tickers : an array to save the name of tickers, which we hardcopy the names for this array.
   - tickerVolumes : an array to save the total revenue volume of each ticker.
   - tickerStartingPrices : an array to save the first price of each ticker.
   - tickerEndingPrices : an array to save the last price of each ticker.<br/>
    to initialize arrays : <br/>
    ![arrays.png](/resources/arrays.png) <br/>

- active the data sheet: 
    
    Worksheets(yearValue).Activate <br/>
    
- Find the number of rows to loop over in the data sheet
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

- bjhbjf

- first For loop ( tickerIndex = 0 To 11 )  : to loop over all the rows in the spreadsheet

- second For loop ( For i = 2 To RowCount ) : to loop over each ticker data, and calculate the total volume of, startinPrice and endingPrice of each ticker

- third For loop ( For i = 0 To 11 ) : to loop through four arrays to output the Ticker name, Total Daily Volume, and Return
11. formatting section 
12. ClearWorksheet() subroutine is for clear the entire worksheet from any formatting 


## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## Summary
address the following questions:
  What are the advantages or disadvantages of refactoring code?
  How do these pros and cons apply to refactoring the original VBA script?
