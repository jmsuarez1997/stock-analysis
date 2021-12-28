# stock-analysis
A stock analysis using VBA to uncover perfomance insight

## Overview of the Project:
This project uses VBA to automate a stock analysis. The analysis uses stock data from green energy companies and creates performance insight for 2017 and 2018. With a few lines of VBA code, the user can click a button to see the percentage of return and Total daily volume traded. This project also walks through the refactoring process of making our VBA code run as efficiently as possible.

## Results:
Below is an outline of the first version of the VBA code:
1) Created an input box for the year value. 
2) Formatted the table for the *All Stocks Analysis Worksheet*

`Worksheets("All Stocks Analysis").Activate
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        'adding a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"`
        
3) Initialized an array of all tickers

`Dim tickers(11) As String
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
        tickers(11) = "VSLR"`

4) Prepared for the analysis of tickers
a) Initialize variables for the starting price and ending price.
b) Activate the data worksheet that uses value from the input box

`Dim startingPrice As Double
 Dim endingPrice As Double
 Sheets(yearValue).Activate`


6) Found the number of rows to loop over – *https://docs.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro*

`RowCount = Cells(Rows.Count, "A").End(xlUp).Row`

7) Looped through the tickers

`For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0`

8) Looped through rows in the data 

`Sheets(yearValue).Activate
 For j = 2 To RowCount`

9) Get total volume for current ticker

`If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value`

10) Get starting price for current ticker 

`If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        startingPrice = Cells(j, 6).Value`

11) Get ending price for current ticker
12) Output the data for the current ticker in the table created

`Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1`

13) Format the data

Below are screenshots of the time elapsed to run the first version of the VBA code for 2017 and 2018.

![Outcomes_vs_Goals](https://raw.githubusercontent.com/jmsuarez1997/stock-analysis/main/Resources/2017_Unfactoredcodetime.png)

![Outcomes_vs_Goals](https://raw.githubusercontent.com/jmsuarez1997/stock-analysis/main/Resources/2018_Unfactoredcodetime.png)

The refactored code changes:
1) Created an input box for the year value
2) Formatted the table for the *All Stocks Analysis Worksheet*
3) Initialized an array of all tickers
4) Prepared for the analysis of tickers
a) Activated the correct worksheet with the stock data
b) Found the number of rows to loop over 
5) *Where updated code starts to change:* 
a)Created a ticket index
    
`Dim tickerIndex As Single
    tickerIndex = 0`

b)Created three output arrays

`Dim tickerVolumes(12) As Long
 Dim tickerStartingPrices(12) As Single
 Dim tickerEndingPrices(12) As Single`

6) Created a loop to initialize the thicker volumes to zero

`For i = 0 To 11
    tickerVolumes(i) = 0
 Next i `

7) Looped over all the rows in the spreadsheet
a) Increase volume for current ticker

`tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`

b) Check if the current row is the first row with the selected ticket index

`If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 3).Value`

c) Check if the current row is the last row with the selected ticket index
d) Increase the ticket index
    
`tickerIndex = tickerIndex + 1`

8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

`For i = 0 To 11
  tickerIndex = i
     Worksheets("All Stocks Analysis").Activate
     Cells(4 + i, 1).Value = tickers(tickerIndex)
     Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
     Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1`

9) Format data for font style, number format, and change interior color red or green for negative and positive return values. 

Below are screen shots of the time elapsed to run the refactored VBA code 2017 and 2018. 

![Outcomes_vs_Goals](https://raw.githubusercontent.com/jmsuarez1997/stock-analysis/main/Resources/VBA_Challenge_2017.png)

![Outcomes_vs_Goals](https://raw.githubusercontent.com/jmsuarez1997/stock-analysis/main/Resources/VBA_Challenge_2018.png)

The refactored code improved the efficiency of the code significantly. For 2017 the first version of the code ran in .75 seconds and the refactored version ran in .1289063 seconds, which is an 82.81% improvement in time. For 2018 the first version of the code ran in .734375 seconds and the refactored version ran in .1289063 seconds which is an 82.45% improvement in time.

## Summary

Some advantages to refactoring code are it will lead to more efficient code, saves time for the user, and saves energy resources. The potential disadvantages to refactoring code are the time and resources it will take to update the new version of the code. In this example, the change in the refactored VBA code was very small and hardly noticeable to the user, but the % of improvement in the seconds elapsed was above 80%. On an individual level, this might not make a huge difference, but if reports like this were being generated automatically this would make a huge impact.
