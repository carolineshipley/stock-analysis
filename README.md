# Green Stock Analysis With VBA

## Overview of Project
Refactor a code created to analyse green stocks daily volumes and annual return using Visual Basic for Application.

### Purpose
The purpose of this project is to make an efficient way to analyse the daily volumes and annual returns of multiple green stocks in 2017 and 2018. Creating a visual way to quickly identify which stocks might be a good investment or not based on annual returns.

## Results
To be able to make the code more efficient a technique called refactoring was applied, which consists in restructuring an existing code to make it more efficient and maintanable.

### Original Code
The original code ran multiple times to assign values one by one to the output table, which ended up taking on average 1.7 seconds to run the code for each year.

This may not seem like a long time however taking into consideration the time taken was to analyse only 12 stocks, this number could be significantly higher if we increased the number of stocks to be analysed. 

With that said is very important to refacture the code to be able to work more efficiently with the current amount of data or potentially with a higher number of data.   

   
    Dim tickers(12) As String
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

    Dim startingPrice As Single
    Dim endingPrice As Single

    Worksheets(yearValue).Activate
    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop over all the rows
      For i = 0 To 11
    
          ticker = tickers(i)
          totalVolume = 0
   

    '5) loop through rows in the data
         Worksheets(yearValue).Activate
        
         For j = 2 To RowCount
    
       '5a) Get total volume for current ticker
              If Cells(j, 1).Value = ticker Then
             totalVolume = totalVolume + Cells(j, 8).Value
              End If

       '5b) get starting price for current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        End If

       '5c) get ending price for current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        End If

        Next j
    
     '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
     
### Refactured Code
To make the code more efficient I created a ticker index and 4 more arrays to improve the logic: ticker, tickerVolumes, tickerStartingPrices and tickerEndingPrices.

Now the code run more efficiently gathering all the information and display on the output table only once. 

This reduced the average run time from 1.7 seconds to 0.3 seconds on average.

    Dim tickers(12) As String
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        ticker = tickers(i)
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
         End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
## Summary
### Advantages and Disadvantages of Refactoring Code
The main advantages of refacturing a code is that the code runs in a more organized way, it's easier to read, easier to debug and runs more efficiently. 

But there's also some disavantages to this suck as the possibility of introducing bugs to the code and in some cases may cost more to refacture than to rewite the code.
### Advantages and Disadvantages of Original and Refactored VBA script
The main advantage of refacturing the code was decreasing considerable the time to run it. As mentioned above the run time went from 1.7 seconds to 0.3 seconds on average. 

Below is an example of the run time for 2017 and 2018 after refacturing.

The disavantage encounter was getting caught up on little bugs which is very time consuming to debug when trying to change an existing code.
#### Stock Analysis for 2018 on Original Code
![VBA 2017 Screenshot](https://github.com/caseychen3605/stock-analysis/blob/master/Resources/VBA_Challenge_2017.PNG)
#### Stock Analysis for 2018 on Refactured Code
![VBA 2018 Screenshot](https://github.com/caseychen3605/stock-analysis/blob/master/Resources/VBA_Challenge_2018.PNG)
