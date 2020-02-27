Sub StockMarketData()
    Dim lastRow As Double
    Dim newChartRow As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearChange As Double
    Dim percentChange As Double
    Dim totalStock As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    
    For Each ws In Worksheets
        
        ActiveWorkbook.Sheets(ws.Name).Activate
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        newChartRow = 2
        openPrice = 0
        closePrice = 0
        totalStock = 0
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        increaseTicker = ""
        decreaseTicker = ""
        volumeTicker = ""
        For i = 2 To lastRow
        
            If openPrice = 0 Then
                'Get the open price for new ticker
                openPrice = Cells(i, 3).Value
            End If
            totalStock = totalStock + Cells(i, 7).Value
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'Get the close price for ticker
                closePrice = Cells(i, 6).Value
                Cells(newChartRow, 9).Value = Cells(i, 1).Value
                'Get yearChange for ticker
                yearChange = closePrice - openPrice
                Cells(newChartRow, 10).Value = yearChange
                'Fill cell depending on change
                If yearChange > 0 Then
                    Cells(newChartRow, 10).Interior.ColorIndex = 4
                ElseIf yearChange < 0 Then
                    Cells(newChartRow, 10).Interior.ColorIndex = 3
                End If
                If openPrice <> 0 Then
                    percentChange = ((closePrice - openPrice) / openPrice)
                Else
                    percentChange = (closePrice - openPrice)
                End If
                Cells(newChartRow, 11).Value = percentChange
                Cells(newChartRow, 12).Value = totalStock
                
                If totalStock > maxVolume Then
                    maxVolume = totalStock
                    volumeTicker = Cells(i, 1).Value
                ElseIf percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    increaseTicker = Cells(i, 1).Value
                ElseIf percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    decreaseTicker = Cells(i, 1).Value
                End If
                newChartRow = newChartRow + 1
                openPrice = 0
                closePrice = 0
                yearChange = 0
                percentChange = 0
                totalStock = 0
            End If
        Next i
        lastRow = Cells(Rows.Count, 11).End(xlUp).Row
        Range("K2:K" & lastRow).NumberFormat = "0.00%"
        'Formatting headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Formatting headers
        Cells(2, 14).Value = "Greatest % increase"
        Cells(3, 14).Value = "Greatest % decrease"
        Cells(4, 14).Value = "Greatest total volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        
        Cells(2, 15).Value = increaseTicker
        Cells(2, 16).Value = maxIncrease
        Cells(2, 16).NumberFormat = "0.00%"
        Cells(3, 15).Value = decreaseTicker
        Cells(3, 16).Value = maxDecrease
        Cells(3, 16).NumberFormat = "0.00%"
        Cells(4, 15).Value = volumeTicker
        Cells(4, 16).Value = maxVolume
        ws.Columns("I:P").AutoFit
    Next
End Sub