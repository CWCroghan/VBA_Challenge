Sub AllStocksAnalysisRefactored()
    'Establishing the variables that will be needed in program
    Dim startTime As Single
    Dim endTime  As Single

    Dim tickers(12) As String
    Dim tickerIndex As Integer
    Dim tickerVolumn(12) As Long
    Dim tickerStartingPrice(12) As Single
    Dim tickerEndingPrice(12) As Single
    
    'Initalizing all of the tickerVolumns.  This variable is additive thus important to initalize.
    
    For k = 0 To 11
        tickerVolumn(k) = 0
    Next k
    
    'Getting year from the User
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers.  Hardcoding the values.  Future effort should be made to do based on data.
    
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
    
    'Starting the ticker index at the first element of the array
    tickerIndex = 0
            
        For i = 2 To RowCount
    
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                    'StartPrice
                    tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
             End If
            
            'Calculating Total Volumn by adding each Daily Return
             If Cells(i, 1).Value = tickers(tickerIndex) Then
                 tickerVolumn(tickerIndex) = tickerVolumn(tickerIndex) + Cells(i, 8)
                End If

             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                    'EndPrice
                    tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
                    'Last row for the ticker therefore need to move to next array element
                    tickerIndex = tickerIndex + 1
                End If
            
        Next i


    'Output the information for all the tickers
        
        Worksheets("All Stocks Analysis").Activate
        
    For k = 0 To 11
        
        Cells(4 + k, 1).Value = tickers(k)
        Cells(4 + k, 2).Value = tickerVolumn(k)
        Cells(4 + k, 3).Value = tickerEndingPrice(k) / tickerStartingPrice(k) - 1
    Next k
    
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
 
    'Ending timer and outputing the run time
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
