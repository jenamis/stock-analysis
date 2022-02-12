Sub AllStocksAnalysis()

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
    
    'Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'Loop over tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
       
        'Loop over rows in data
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
    
            'Find total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
    
            'Find starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
    
            'Find ending price for current ticker
            If Cells(j, 1) = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If

        Next j

        'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
    
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
    
    Next i
    
    'Format All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Bold header row
    Range("A3:C3").Font.Bold = True
    'Add bottom border to header row
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Format volume
    Range("B4:B15").NumberFormat = "#,##0"
    'Format return
    Range("C4:C15").NumberFormat = "0.0%"
    'Autofit volume column
    Columns("B").AutoFit
    
    'Format return cell color
    DataRowStart = 4
    DataRowEnd = 15
    
    For i = DataRowStart To DataRowEnd
    
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
