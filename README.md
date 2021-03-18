# An Analysis of Stocks (2017 - 2018)

## Overview of Project
Steve was pleased with the workbook that I prepared for him. He has requested that I prepare analysis for the entire stock market over the last few years.
I have refactored code previously used and have a new VBA script that will analyze the entire data set and run more efficiently, as well.

### Purpose
In the analysis, I will be reviewing twelve different stocks and the following two outputs:

- Total Daily Volume
- Return

The two years that will be included in the analysis are: 

- 2017
- 2018

Using my refactored script, I will be able to provide Steve with a breakdown of which stocks provided the best return from their starting price and ending price in each year.

 
## Results
### Analysis of 2017 Stocks
Looking at the 2017 results we can observe the following:

-	*Largest Return* - DQ: 199.4%
-	*Smallest Return* - TERP: -7.2%
-	*Largest Total Daily Volume* - SPWR: 782,187,00

![image](https://github.com/jb-ut/stock-analysis/blob/main/VBA_Challenge_2017-AllStocksAnalysis.PNG)

### Analysis of 2018 Stocks
Looking at the 2018 results we can observe the following:

-	*Largest Return* - ENPH: 81.9%
-	*Smallest Return* - DQ: -62.6%
-	*Largest Total Daily Volume* - 607,473,500

![image](https://github.com/jb-ut/stock-analysis/blob/main/VBA_Challenge_2018-AllStocksAnalysis.PNG)

Overall, 2017 was a great year for the stock market while 2018 was a bad year for most of the high performing stocks from 2017. Two stocks performed well in both years: ENPH and RUN.
Based on the results of these two years, it is recommended that both stocks should be considered for investment. 
## Summary

### What are the advantages or disadvantages of refactoring code? 
One benefit of refactoring code is that is being able to reduce the runtime through changes in organization and type of script used. In this example, I was able to reduce the runtime while also expanding the breadth of my analysis by combining a for Loop with the tickerIndex.
Here is a screenshot of the code used for this analysis.
A disadvantage of refactoring code is that it is possible to introduce new bugs into working code. Also, depending on the type of analysis there may be limited upside in updating code from an efficiency standpoint. 

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
    
        Dim tickerIndex As Integer
        tickerIndex = 0

    '1b) Create three output arrays
        
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
        For i = 0 To 11
        
            tickerVolumes(i) = 0
            
        Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
        
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
             tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     
        End If

            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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

### How do these pros and cons apply to refactoring the original VBA script?
In the case of this stock analysis, there was an opportunity to improve the scale and efficiency of the stock analysis VBA script. Based on the request and the sample code provided, there was a clear path forward to refactor the code and analyze more data and improve efficiency.
The prior scripts took several seconds to run and the refactored scripts both take under one second to run providing quicker results for the Steve. 

Below are the execution times of the new script.

![image](https://github.com/jb-ut/stock-analysis/blob/main/VBA_Challenge_2017.PNG) ![image](https://github.com/jb-ut/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

