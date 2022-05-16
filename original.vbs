
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

    '1) Format the output sheet on All Stocks Analysis Worksheet

    'Activate "All Stocks Analysis" worksheet
    Worksheets("All Stocks Analysis").Activate

    'Title Analysis
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a Header Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2)Initialize an array of all tickers.
    
    'Declare an array with 12 string elements
    Dim tickers(12) As String
    
        'Assign tickers to an element in the array
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
    '3) Prepare for the analysis of all tickers.

    '3a) Initialize variables for the starting price and ending price.
    
        'Creating a Variable for Starting & Ending Price
        Dim startingPrice As Double
        Dim endingPrice As Double
    
    '3b) Activate the data worksheet.
        
        Worksheets(yearValue).Activate
        
    '3c) Find the number of rows to loop over.
        
        rowStart = 2
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    '4) Loop through the tickers.
    
    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0

        '5) Loop through the rows in the data.

        'Activate Data Worksheet
        Worksheets(yearValue).Activate
        
        For j = rowStart To RowCount
        

    
        '5a) Find the total volume for the current ticker.
    
            'Identify ticker
            If Cells(j, 1).Value = ticker Then
                
                'increase ticker totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
        '5b) Find the starting price for the current ticker.
    
            'Identify first row of ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
                
            End If
            
        '5c) find the ending price for the current ticker.
    
            'Identify last row of ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
        
    
    
    '6) Output the data for the current ticker.

        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker header
        Cells(i + 4, 1).Value = ticker
    
        'Sum of Volume
        Cells(i + 4, 2).Value = totalVolume
    
        'Return Value
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub