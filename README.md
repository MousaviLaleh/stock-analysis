# stock-analysis
##### :card_file_box: [Download the File](green_stocks_01.xlsm)

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

- Create two For loop to work on data
    - first For loop : to loop through the tickers and initialize the tickerVolumes to zero.
    - inner For loop + conditional : to loop over all the rows in the spreadsheet, and calculate the total volume, startinPrice and endingPrice for each ticker <br/>
  ![forloops_01.png](/resources/forloops_01.png) <br/>
multiple conditions check if the current row is the first row with the selected ticker. If it is, then assign the current price to tickerStartingPrices variable. Similarly, for the tickerEndingPrices variable.
    

- third For loop: to loop through four output arrays to pull the results from, and show the output  in "All Stocks Analysis" sheet.<br/>
  ![forloops_02.png](/resources/forloops_02.png) <br/>

- formatting section: by adding some font style, borders, number formatsand also highlishts, we make it easier for Steve to read the data.<br/>
![formatting.png](/resources/formatting.png) <br/>


## Results
By running the code, first window is asking the year, then we have the result in the "All Stocks Analysis" sheet with the highlight of stocks' rise and falls, and also code runtime which indicates the performance of the refactored code. <br/>
![msgBox.png](/resources/msgBox.png) <br/>
![result_2017.png](/resources/result_2017.png) --- 
![result_2018.png](/resources/result_2017.png) <br/>
![runTime_2018.png](/resources/runTime_2018.png) <br/>
Most tickers have significant drop in their stocks in 2018, which narrows the Steve's investing down to two tickers, RUN and ENPH. The RUN ticker has a skyrocket rise in the stock in compare to 2017.<br/>

## Summary
Refactoring is intended to improve the design and structure of the code, while preserving its functionality. It makes the code easier to understand but it takes time. By using arrays and indexes, user can 

address the following questions:
  What are the advantages or disadvantages of refactoring code?
  How do these pros and cons apply to refactoring the original VBA script?
