# stock-analysis

## Overview of Project
This project looks at stock data for 2017/2018 and calculate for each stock (1) the yearly volume or how often a stock is traded ; and (2)the yearly return or % of increase/decrease in price from the beginning of the year to the end of the year.  

### Purpose
The purpose of this project is to provide an end user with a summary analysis containing yearly volume and yearly return for each stock from thousands of lines of stock market dataset.

## Results
In ***VBA_Challenge.xlsm workbook*** (https://github.com/legeren/stock-analysis/blob/196eb169fc607fbd54534161a6b012c909e1ba4c/VBA_Challenge.xlsm), the end user can run the macro and see stock performance for 2017/2018.  At the end of the code, the end user will also find out how long the macro ran for.  

One key difference between the original and refactored codes is the creation of the ticker index for the refactored code.  This ticker index is what later defines how each volume and starting/ending price is calculated instead of having nested loops.

### Original code snippet:
```
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0

'5. Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
       
'5a. Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
            
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If

'5b. Find the starting price for the current ticker.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        End If

'5c. Find the ending price for the current ticker.

        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        
        Next j
```

### Refactored code snippet
```
    For i = 0 To 11

     tickerVolumes(i) = 0
     tickerStartingPrices(i) = 0
     tickerEndingPrices(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
        Worksheets(yearValue).Activate
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.

            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If

        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

 

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        'End If
            End If
        
        Next i
```

### Runtime Comparison
Comparing the final runtime for the codes between the original and refactored code, end user can find that the refactored code ran quicker.  
- *2017 original code runtime* https://github.com/legeren/stock-analysis/blob/196eb169fc607fbd54534161a6b012c909e1ba4c/Resources/Allstockanalysis2017.png vs *2017 refactored code runtime* https://github.com/legeren/stock-analysis/blob/196eb169fc607fbd54534161a6b012c909e1ba4c/Resources/VBA_Challenge_2017.png
- *2018 original code runtime* https://github.com/legeren/stock-analysis/blob/196eb169fc607fbd54534161a6b012c909e1ba4c/Resources/Allstockanalysis2018.pngvs *2018 refactored code runtime* https://github.com/legeren/stock-analysis/blob/196eb169fc607fbd54534161a6b012c909e1ba4c/Resources/VBA_Challenge_2018.png

## Summary

### Refactoring Code
- Pro: Generally, refactoring code improves the efficiency of the code for run time.  It also allows the code to be more flexible in case of changing dataset.
- Con: A disadvantage of refactoring code could include being stuck with the errors (if there are any) from the original code and having to comb through it before understanding next steps.

### Refactored VBA script
- Pro: The original VBA script was limited only to the existing dataset.  The refactored code will be more flexible in its ability to handle any additional data.
- Con: Seeing that the results are the same since the dataset did not change, refactoring this VBA script took hours of extra coding ending with the same results as the original code.
