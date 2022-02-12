Sub AllStocksAnalysisRefactored()

    'Create variables to calculate macro execution time
    Dim startTime As Single
    Dim endTime As Single

    'Determine year for analysis
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Start clock on macro execution time
    startTime = Timer
    
    'Format output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(11) As String
    
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
    
    'Get number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker index variable and set it equal to 0
    Dim tickerIndex As Integer
    tickerIndex = 0

    'Create three output arrays
    
        'Create ticker volume array
        Dim tickerVolumes(11) As Long
        
        'Create ticker starting price array
        Dim tickerStartingPrices(11) As Single
        
        'Create ticker ending price array
        Dim tickerEndingPrices(11) As Single
    
    'Create a for loop to initialize tickerVolumes to zero
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
        
    'Loop over all rows in data worksheet
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
            
        'Check if current row is first row with selected ticker and get starting price
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Check if current row is last row with selected ticker and get ending price
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Increase tickerIndex if next row’s ticker doesn’t match current row's ticker
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerIndex = tickerIndex + 1
        End If
        
    Next i
    
    Next
    
    'Loop through arrays to output to Ticker, Total Daily Volume, and Return columns
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
            
    Next i
    
    'Format All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Bold header row
    Range("A3:C3").Font.FontStyle = "Bold"
    'Add bottom border to header row
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Format volume
    Range("B4:B15").NumberFormat = "#,##0"
    'Format return
    Range("C4:C15").NumberFormat = "0.0%"
    'Autofit Total Daily Volume column
    Columns("B").AutoFit

    'Format return cell color
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
            
        Else
            Cells(i, 3).Interior.Color = xlNone
            
        End If
        
    Next i
 
    'Stop clock on macro execution time
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
