# Stock Analysis with VBA

## Overview of Project

### Instead of using formulas in a excel workbook, it is better to use code in Microsoft Excel's VBA Script Editor. This allows one to write code and automate a process through code and sometimes a simple click of a button.

## Results

- The code is intended to ask the user what year they would like analysis on. Then it creates an array of data that includes the stock information for each stock.
- The stocks in 2017 compared to 2018 seemed to have a higher return. The only stock that had a negative return in 2017 was TERP. All other stocks had quite large returns, only a couple positive returns were single digits.
- In 2018, on the other hand, only two stocks had a positive return and were both in the 80% range. All other stocks had a negative return with less daily volume as well. There were a few that had higher daily volume, but I'm not sure if the change in volume is attributable or correlates to the return rate of the stock.

[2018 Stocks](https://github.com/mbugyis/stocks-analysis/blob/main/Resources/2018_Stocks.png)

[2017 Stocks](https://github.com/mbugyis/stocks-analysis/blob/main/Resources/2017_Stocks.png)


- After formatting, initializing, and setting varibles...My code looked like this:

'''
    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
            
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
    'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
    'If  Then
    
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
    '3d Increase the tickerIndex.
    
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
                
            tickerIndex = tickerIndex + 1
        
        End If
    
    
    'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
        
    Next i
'''

## Summary

### Advantages and/or disadvantages of refactoring code

- An advantage of refactoring code is that it applies small changes without really altering what the code is intended to do. A lot of times people write code and end up making it longer than it really should be. Or they write different scripts that can't be executed at once, and they must be brought together. For example, a formatting script and a data analysis script. By combining the two, it makes the process easier and still acheives the same goal.
- A disadvantage may be that a step may be done incorrectly. For example, if a variable is changed, or one aspect of the original code must comply with a new aspect of the refactored, and one must go in and change every instance of were a variable is, or an index, etc. This can cause several syntax errors and may complicate the process more. This is particular when combining two subscripts like mentioned above. Especially when two different people are working on it, they usually have their own variable names or ideas to write code, and they may not align.


### Relationship between Refacorted and Original VBA

- The original code was nice and was able to do the same thing, if not all similar aspects, of the refactored code. In the refactored code, I was able to loop through the output data at the beginning at set it to zero for each stock/index. Setting the index at the beginning was better in my opinion, so one didn't have to loop through the ticker each time. But in the end, the new refactored code went to the next stock index in the embedded for loop, not in the first for loop like the original code. I found this difficult at first, but later understood it.

- I also liked how instead of using the stock ticker when getting a price or volume, the stocks index was used. This made storing the output values such as ticker, start price, ending price, and volume much more easier to output in the end. Because when outputting, one could just run through another for loop which would go through each stock index.

