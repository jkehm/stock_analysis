# Stock Analysis Module 2 Challenge

## Overview

### The Data
The data that we will be working with was provided. It consists of two worksheets titled "2017" and "2018". These sheets contain the stock information with a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. Our goal is to utilize macros in order to output a table with the 12 different ticker values, as well as the Total Daily Volume, and Return on the stock for that year. This will allow Steve's parents to make a more informed decision when looking at these stock options.  

### The Purpose
We have already helped our friend Steve create an Excel workbook with Macros utilizing VBA (Visual Basic Analysis) programming. The Macro in class works pretty well for a few stocks, however it would struggle with a larger dataset. In this assignment we will refactor the code we had already written. Ideally, the macro will run much more efficently and take significantly less time to run.

## Results

#### Below is a direct comparison of the Results for all of the stocks in the years 2017 and 2018. Overall, it is rather clear that 2017 was a much better year overall for the Return of these stocks compared to 2018. If Steve's parents were to invest in any stocks, I would recommend ENPH or RUN as the primary two. These two saw positive returns in both years. Otherwise, VSLR and SEDG could both be valid options as they grew and then had a much smaller decline than some of the other stocks. I would be careful investing in TERP as it was the only Ticker that had negative growth in both years.
![All_Stocks_2017_Result](https://github.com/jkehm/stock_analysis/blob/main/Resources/All_Stocks_2017_Result.png)
![All_Stocks_2018_Result](https://github.com/jkehm/stock_analysis/blob/main/Resources/All_Stocks_2018_Result.png)

#### Below are the Results for the run-time of the script. This is about 10 times more efficent than the original script that was written.
![VBA 2017 Screenshot](https://github.com/jkehm/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_2018_Screenshot](https://github.com/jkehm/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)


#### The full AllStocksAnalysisRefactored Script Can be Found Below
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
     For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
        
    For i = 2 To RowCount
               
        
                    
        '3a) Increase volume for current ticker
        
           
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
      
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            'Stores Starting Price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            'Stores Ending Price
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
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```

## Summary
####   1) What are the advantages or disadvantages of refactoring code?
##### Once a script or macro is written and works succesfully, some people may think their work is done and it is time to move on to the next project. However, there are some reasons to go back and refactor your code to work better. Of course, there are some reasons and situations where it may **not** be worth taking the extra time to do this. 
##### The most obvious advantage to going back to refactor code could be to decrease the run time of the script. This may not seem like a big deal on a fairly small data set like we used here. However, as datasets get larger you could be saving several minutes in run time if the script is written more efficently. Another benefit is to go back and add comments, or add more detail to existing comments. This way, if the script needs to be changed you still understand what everything is doing, even if it is years later. Or if someone else were to pick up the project they can make sense of the code.
##### The biggest disadvantage is that the process of refactoring can be very time consuming. And in a company setting, it may not be worth it to invest more time, energy, and money into a project if it is already satisfactory. Theoretically refactoring a code that already works may lead to bugs in the code as well, if precautions are not taken. 

####   2) How do these pros and cons apply to refactoring the original VBA Script?
##### In this example the main benefit that was clearly noticed was the efficency of the new script. Where the original Script took right about 1 second to run, the refactored version takes about 0.1 seconds to run. The comments on this script were also more step-by-step and easier to follow for someone else to understand what each piece of code is doing.
##### The first disadvantage I mentioned above would not apply to this situation at all. Since we are still learning, refactoring this code was a very useful exercise and we are not costing a company anything by continuing to work on this script. Causing bugs or having run issues is a valid disadvantage. But again, we are learning here, so the process of debugging is extremely important to get thorough and learn from. 
