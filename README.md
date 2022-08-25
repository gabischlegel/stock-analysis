# stock-analysis

## Oveview of Project
###
The purpose of this analysis was to help Steve’s parents decide which stock to invest in and which stock to not invest in from the years 2017 and 2018. It was also to refactor the code so it could become for efficient and run faster. 
## Analysis
###
The refactored code was extremely quick as it ran in .15625 seconds for 2017 and .1796875 amount of seconds for 201. The refactored code ran around a full second faster! In 2017, the stocks had a more favorable return rate, with only 1 losing money. In 2018, the stock were all losing money, with only 2 having a positive rate of return. 
###
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
    
    '1a) Create a ticker Index and set it to 0

    tickerIndex = 0

    
    '1b) Create three output arrays for tickerVolumes (As Long), tickerStartingPrices (As Single), and tickerEndingPrices (As Single)
    
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
            
        tickerVolumes(i) = 0
        
        'tickerStartingPrices(i) = 0
        
        'tickerEndingPrices(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.

    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
    
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
         
         'If  it is then the current price is assined to the tickerStartingPrices variable
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            'If  Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            '3d increase ticker Index
        
            tickerIndex = tickerIndex + 1
        
        
            End If
    

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
    Next i
    
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate

        
        'Inputs ticker name
        Cells(4 + i, 1).Value = tickers(i)
            
        'Inputs daily volume value
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'Inputs percentage of return value
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

## Summary
###
Refactoring the code helps it become easier to read, takes less memory, and takes fewer steps. Since the initial code is usually not the most efficient, refactor helps clean up the previous code. Due to the code being cleaner, it is easier to catch bugs in the code. Refactoring a code can pose problems if the first code was not well written. This means it could take more time to refactor a code than re write it. 
###
Compared to the original VBA script, the code ran much faster and is much easier to read. The refactored code ran in .15625 seconds for 2017 and .1796875 amount of seconds for 2018 (see images below), while the original code to 1.179668 seconds to run for 2018 and 1.2130928 for 2017. The only con was being new to refactoring codes, refactoring was slightly time consuming. 
