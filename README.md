# Green Stock Analysis

## Overview of Project

  Steve requested for me to do an analysis on green stock for the sole purpose of his parents, to see whether or not specific stocks are worth invensting in. I thought it was the most effecient to use Visual Basic Application (VBA) to complete this task, and ultimately find the annual return on this specific stock. I was able to analyze multiple green stocks, and ultimately give Steve and his parents a very interesting analysis which will hopefully give them an idea of what stocks to invest in, and which stocks to let go of. 
  
  To optimize our visualization of these stocks, we first needed to run the analysis of the 12 different stocks initially, and then work our way to refactor the code, to not only make the analysis more efficient, but to also make the run time more efficient. 
 
 ## Results of the Analysis
 
 ### Refactoring the Code
 
  The first thing I did to make the code more efficient was to reverse my for nesting loops by creating 3 different arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) on top of the ticker array. The tickers array was simply to determine the ticker symbol of a specific stock. The three arrays were matched using a tickerIndex variable.
  
Please see below code:

#### Refactored Code 

    '3) Initialize array for all tickers
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
    
    '4a) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '5) Loop over number of rows
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '6) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '7) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim TickerEndingPrices(12) As Single
    
    '8) Initialize ticker volumes to zero
        
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    
    '9) loop over rows
    
    For i = 2 To RowCount
    
        '10) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
        
        '11) Make sure the current row is the first row with this specific selected tickerIndex.
        If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
            
            
        End If
        
        '12) see if the current row is the last row with the selected ticker
        If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
            TickerEndingPrices(tickerIndex) = Cells(i, 7).Value
            

            '13) Now we increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '14) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = TickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i


#### Original Code
2) Initialize array of all tickers

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

'3a) Initialize variables for starting price and ending price

    Dim startingPrice As Double
    Dim endingPrice As Double

'3b) Activate data worksheet

    Worksheets(yearValue).Activate

'3c) Get the number of rows to loop over

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through tickers

    For i = 0 To 11
    ticker = tickers(i)
    TotalVolume = 0
    Worksheets(yearValue).Activate

'5) loop through rows in the data
        
For j = 2 To RowCount

    '5a) Get total volume for current ticker

    If Cells(j, 2).Value = ticker Then

        'increase totalVolume by the value in the current row
        TotalVolume = TotalVolume + Cells(j, 9).Value

End If

        '5b) get starting price for current ticker

    If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
        'set starting price
        startingPrice = Cells(j, 7).Value

    End If

        '5c) get ending price for current ticker
        
        If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
        'set ending price
        endingPrice = Cells(j, 7).Value

    End If

    Next j
'6) Output data for current ticker

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

 Next i
 
 #### Analysis
 
  As you can clearly see, the tickerindex variable allowed for me to assign each ticker symbol to tickerVolumes, tickerStartingPrices, and tickerEndingPrices respectively. 
  
### Run Time 

#### Refactored Code 

![Image 11-22-20 at 9 32 PM](https://user-images.githubusercontent.com/74481469/99934249-585bbb00-2d12-11eb-991c-3050488d72e7.jpeg)

![Image 11-22-20 at 9 32 PM (1)](https://user-images.githubusercontent.com/74481469/99934533-1aab6200-2d13-11eb-9d39-c7fa17c36d3f.jpeg)

Clearly, the refactored code is running much faster than the orginal code! 



