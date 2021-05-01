# stock-analysis
Repository for stock analysis for VBA module


# Module 2 Challenge: Visual Basic for Applications


## Overview of the Project

The purpose of this initial analysis was to analyze 12 stocks for the yearly return over 2017 and 2018. Following that analysis, our code was refactored to be more efficient so as to be more usable for lager data sets, i.e., the entire stock market per the request of Steve. 


## Results

Analysis was performed on the green energy company DAQO New Energy Corp. ($DQ). Yearly returns were calculated by taking the percent difference between the closing prices of DAQO at the beginning and end of 2018. Results showed that DAQO New Energy Corp had a yearly return of -62.60 percent (%). See image below. 

DQ_Return.png![image](https://user-images.githubusercontent.com/82423123/116794230-0ec98600-aa91-11eb-8be0-a6d7c953daaa.png)


Following that analysis, Steve was interested in performing the same analysis on a larger dataset and over two years (2017 and 2018). A more comprehensive script was prepared to analyze the yearly return and total trading volume of 12 companies. the code is available in the yearValueAnalysis.bas file, or as the yearValueAnalysis module in the VBA_Challenge.xlsm file. Additionally, the code is provided below. 

```Sub AllStocksAnalysis()

    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks 2018"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
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
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    Worksheets("2018").Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Worksheets("2018").Activate
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
            End If

        Next j
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   
   'Next i'
   
End Sub```
```

Results of the 12 stocks in 2017 and 2018 are presented below. 

![All_2017](https://github.com/curtissmith291/stock-analysis/blob/main/Resources/All_2017.png)

![All_2018](https://github.com/curtissmith291/stock-analysis/blob/main/Resources/All_2018.png)

After seeing the positive applications of the aforementioned code, Steve wanted the code to be refactored so as to be more viable for a larger dataset. The runtimes for the aforementioned code are presented below. 

![VBA_CHallenge_2017_before](https://github.com/curtissmith291/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Before.png)

![VBA_Challenge_2018_Before](https://github.com/curtissmith291/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Before.png)

The runtimes were relatively high, approximately 0.85 seconds for both 2017 and 2018. As Steve wanted to perform the analysis on a much larger dataset, all stocks, a more efficient script would need to be prepared. The previous code was refactored so that the for loops only ran through each row of data once, rather than twice. The refactored code is available in the AllStocksAnalysisRefactored.bas file, in the VBA_Challenege.xlsm file as module AllStocksAnalysisRefactored, and below. 

```Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single'

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
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = LBound(tickerVolumes) To UBound(tickerVolumes)
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
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            

            '3d Increase the tickerIndex if the next row's ticker doesn't match the previous row's ticker
            
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
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

The updated runtimes are presented below. 

![VBA_Challenge_2017_After](https://github.com/curtissmith291/stock-analysis/blob/main/Resources/VBA_Challenge_2017_After.png)

![VBA_Challenge_2018_After](https://github.com/curtissmith291/stock-analysis/blob/main/Resources/VBA_Challenge_2018_After.png)

As shown in the above images, the runtime was significantly improved to approximately 0.22 seconds, approximately 0.63 seconds faster than the non-refactored code. 


## Summary

Advantages/Disadvantages of Refactoring Code

Depending on the required analysis performed and the timeframe for deliverables, the effort spent refactoring code could worthwhile. If the code is to be used many times over, on large datasets, or for multiple types of applications, refactoring could save time in the long term. Additionally, the time spent refactoring could help optimize the code for more broad applications rather than a narrow focus. However, some analyses just need to be run a few times, or just once, or the time frame for the deliverable is short enough the spending additional hours refactoring would delay the project. 

Advantages/Disadvantages of the Original and Refactored VBA Script

The first disadvantage for both scripts (even more so for the intended use of the refactored script, i.e., analysis of all stocks) is that the ticker array needs to be manually populated; a script to automatically populate the ticker array is included below:

```Sub arrTest()
    
    Worksheets("2018").Activate
    
    Dim tickers() As String, i As Integer, y As Integer
    y = 0 'variable for number of indices in tickers()
    rowStart = 2
    ReDim tickers(0)
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row 'determines the number of rows
    
    For i = rowStart To rowEnd
        'Below adds the first item to the array; the array always needs to be 1 value lager than the number of items to add another item
        If y = 0 Then
            tickers(y) = Cells(i, 1).Value 'Adds the first ticker to the only empty space inside the array
            y = y + 1 'ups the count of spaces in the array by 1 so another ticker can be added
            ReDim Preserve tickers(y) 'redimensioning array to include 1 more index
        Else
            For j = LBound(tickers) To y - 1 'loops through the tickers() array to check value in cell is already in the array
                'MsgBox ("Cell value = " & "j is " & Cells(i, 1).Value & j & ";  ticker(j) is " & tickers(j)) 'displays the count of j and the corresponding ticker assigned to that index in the array
                If Cells(i, 1).Value = tickers(j) Then 'checks if worksheet ticker string is in array
                    Exit For 'if ticker in the worksheet  value equals ticker(j), then it is in the array already; for loop ends, iterates to next "i"
                ElseIf Cells(i, 1).Value <> tickers(j) Then 'condition if worksheet ticker value does not equal the ticker in the array at the j index position
                    If tickers(j) = tickers(y - 1) Then 'checks if j is at second to last position in array (last index is empty so another value can be added)
                        tickers(y) = Cells(i, 1).Value 'adds worksheet ticker value to array
                        'MsgBox (tickers(y) & " added to array; iterating to next i") 'sanity check to see if tickers are getting added, delete when running full program
                        y = y + 1
                        ReDim Preserve tickers(y)
                        'MsgBox ("iterating to next i after adding ticker to array")
                        Exit For
                     Else
                        'MsgBox ("Iterating to next j")
                     End If
                End If
            Next j
        End If
    Next i
    
    Worksheets("test").Activate
    
    'prints array to cells in a column
    For i = LBound(tickers) To y - 1 'loops through array, note that last position (y) is empty, so the last ticker is in y-1
        Cells(i + 1, 1).Value = tickers(i) 'i is initially 0, to need to add 1 (there is no "0" row to write to)
    Next i
    
    
End Sub
