Sub stocks()
    'Loop through all sheets
    For Each ws In Worksheets
    
        'Create the new column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'define last row
        Dim lrow As Long
        lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'define variables
        Dim TickerName As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim StockVolume As LongLong
    
        'Keep track of the location of each ticker name in the summary table
        Dim SummaryTableRow As Integer
        SummaryTableRow = 2
    
        '------------------------------------
        'LOOP THROUGH ALL THE STOCKS
        '------------------------------------
        For i = 2 To lrow
            'look to see if this is a new ticker name (if the value of i,1 has changed from the row above)
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'set the ticker name
                TickerName = ws.Cells(i, 1).Value
                
                'Add the ticker name to the summary list
                ws.Range("I" & SummaryTableRow).Value = TickerName
    
                'set the open price (this is the first open day of the year)
                OpenPrice = ws.Cells(i, 3).Value
               
                'start adding up the stock volume
                StockVolume = 0
                StockVolume = ws.Cells(i, 7).Value + StockVolume
                 
            'if you look above and below and everything is the same
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value And ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                'keep adding to the stock volume
                StockVolume = CDec(ws.Cells(i, 7).Value + StockVolume)
            
            'if it's not a new ticker name, see if it's the last ticker name in the list (if the value of i,1 changes in the following row)
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'capture the close amount
                ClosePrice = ws.Cells(i, 6).Value
    
                'calculate the yearly change and add it to the summary list
                YearlyChange = ClosePrice - OpenPrice
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                If YearlyChange > 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4 'Green
                ElseIf YearlyChange < 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3 ' Red
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 2 'White
                End If
    
                'calculate the percent change and add it to the summary list
                If OpenPrice <> 0 Then
                   PercentChange = YearlyChange / OpenPrice
                   FormPercentChange = Format(PercentChange, "Percent")
                   ws.Range("K" & SummaryTableRow).Value = FormPercentChange
                Else
                    ws.Range("K" & SummaryTableRow).Value = "none"
                End If
            
                'add to the stock volume, then add it to the summary list
                StockVolume = Cells(i, 7).Value + StockVolume
                ws.Range("L" & SummaryTableRow).Value = StockVolume
                ws.Columns("L:L").EntireColumn.AutoFit
                
                'add a new row in the summary table
                SummaryTableRow = SummaryTableRow + 1
    
                'reset everything
                OpenPrice = 0
                ClosePrice = 0
            End If
        Next i
    Next ws

    '---------------------------------
    'LOOK THROUGH ALL WORKSHEETS FOR "GREATEST" AMOUNTS
    '---------------------------------
    
    'Set up the Greatest table
    Sheet1.Cells(2, 15).Value = "Greatest % Increase"
    Sheet1.Cells(3, 15).Value = "Greatest % Decrease"
    Sheet1.Cells(4, 15).Value = "Greatest Total Volume"
    Sheet1.Cells(1, 16).Value = "Ticker"
    Sheet1.Cells(1, 17).Value = "Value"
    Sheet1.Columns("O:O").EntireColumn.AutoFit
    
    'define "Greatest" variables
    Dim Increase As Double
    Increase = 0
    Dim IncreaseTicker As String
    Dim Decrease As Double
    Decrease = 0
    Dim DecreaseTicker As String
    Dim TotalVolume As LongLong
    TotalVolume = 0
    Dim TotalVolumeTicker as String
    
    'look through all the sheets
    For Each ws In Worksheets
        For i = 2 To lrow
            If ws.Cells(i, 10).Value > Increase Then
                Increase = ws.Cells(i, 10).Value
                IncreaseTicker = ws.Cells(i, 9).Value
            ElseIf Cells(i, 10).Value < Decrease Then
                Decrease = Cells(i, 10).Value
                DecreaseTicker = Cells(i, 9).Value            
            End If
        Next i
        For i = 2 To lrow
            If ws.Cells(i, 10).Value > TotalVolume Then
                TotalVolume = ws.Cells(i, 12).Value
                TotalVolumeTicker = ws.Cells(i, 9).Value
            End If
        Next i
    Next ws

    'Populate the "Greatest" values
    Sheet1.Cells(2, 16).Value = IncreaseTicker
    Sheet1.Cells(2, 17).Value = Increase
    Sheet1.Cells(3, 16).Value = DecreaseTicker
    Sheet1.Cells(3, 17).Value = Decrease
    Sheet1.Cells(4, 16).Value = TotalVolumeTicker
    Sheet1.Cells(4, 17).Value = TotalVolume
End Sub