# :seedling: Green Stock Analysis


## Overview of Project
Steve, a friend, is passionate about Green Energy stocks and would like to find out which company has the best performance, to invest in, specifically DAQO New Energy Corporation, a company which makes silicon wafers for solar panels.<br/>
We will be using VBA ( Visual Basic for Appllication ) throughout this project, to automate tasks and reuse codes for any stock. VBA reduces the chance of accidents and errors and increases the code running time.<br/>

### Purpose
In the first step we found out that the DAQO Corp had a drop over 63% in 2018. Now, Steve wants to expand his research to include the entire stock market, to find some better stocks. We have an Excel file containing the stock data over the last few years. The purpose is improving the logic of the VBA code in order to make it more efficient to works well for thousands of stocks. At the click of a button, Steve can analyze an entire dataset.<br/>


## Code Review
##### :card_file_box: [refactored file - VBA_Challenge.xlsm](VBA_Challenge.xlsm)
##### :card_file_box: [original file - Challenge.xlsm](Challenge.xlsm)
The refactored code follows the steps:
- Set the startTime and endTime variables to measure and show the code running time.<br/>
    ![times.png](/Resources/times.png)<br/>
    

- InputBox:  to select the year by user. <br/>
    ![inputbox.png](/Resources/inputbox.png)<br/>

- Create a tickerIndex variable to loop over output arrays, and set it to zero. <br/>
  ![tickerIndex.png](/Resources/tickerIndex.png) <br/> 

- Format output Sheet: output sheet, is the sheet that we show the code results in that. In this project "All Stocks Analysis" is the output sheet.<br/>
    to format the output data sheet:  <br/>
                                        Worksheets("All Stocks Analysis").Activate <br/>
       
- Initialize arrays that save the result for every ticker: 
   - tickers : an array to save the tickers' name, which we hardcopy the names for this array.
   - tickerVolumes : an array to save the total stock volume of each ticker.
   - tickerStartingPrices : an array to save the first price of each ticker.
   - tickerEndingPrices : an array to save the last price of each ticker.<br/>
    to initialize arrays : <br/>
    ![arrays.png](/Resources/arrays.png) <br/>

- Active the data sheet: 
    
    Worksheets(yearValue).Activate <br/>
    
- Find the number of rows to loop over in the data sheet
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

- Create For loops to work on data:
    - first For loop : to loop through the tickers and initialize the tickerVolumes to zero.
    - inner For loop + conditional : to loop over all the rows in the spreadsheet, and calculate the total volume, startinPrice and endingPrice for each ticker <br/>
    ![forloops_01.png](/Resources/forloops_01.png) <br/>
     Multiple conditions check if the current row is the first row with the selected ticker. If it is, then assign the current price to tickerStartingPrices variable. Similarly, for the tickerEndingPrices variable.
    

- third For loop: to loop through four output arrays to pull the results from, and show the outputs in "All Stocks Analysis" sheet.<br/>
  ![forloops_02.png](/Resources/forloops_02.png) <br/>

- formatting section: by adding some font style, borders, number formats and highlishts, we make it easier for Steve to read the data.<br/>
  ![formatting.png](/Resources/formatting.png) <br/>


## Results
By running the code, the first pop-up window is asking the year. After entering the year, we have the result in the "All Stocks Analysis" sheet with the highlight of stocks' rise and falls, and also code runtime which indicates the performance of the refactored code. <br/><br/>
![msgbox.png](/Resources/msgbox.png)<br/>
Code runtime for original code : <br/>
![runtime2017.png](/Resources/runtime2017.png) - - - - - - 
![runtime2018.png](/Resources/runtime2018.png)<br/>

Code runtime for refactored code : <br/>
![runtime_2017.png](/Resources/runtime_2017.png) - - - - - - 
![runtime_2018.png](/Resources/runtime_2018.png) <br/><br/>
Results: <br/>
![result_2017.png](/Resources/result_2017.png) - - - - - 
![result_2018.png](/Resources/result_2018.png) <br/><br/>
:small_orange_diamond: Most tickers have significant drop in their stocks in 2018, which narrows the Steve's investing down to two tickers, RUN and ENPH. The RUN ticker has a skyrocket stock rise in 2018 in compare to 2017.<br/>

## Summary
Refactoring is intended to improve the design and structure of the code, while preserving its functionality. It makes the code easier to understand but it takes time. A huge risk with refactoring is that the errors may destroy an already working code. It is highly recommended to save the original code, the way we can always go back a step without needing to start completely over. <br/><br/>

