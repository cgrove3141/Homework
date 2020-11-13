Attribute VB_Name = "Module1"
Sub StockChanges()

    For Each ws In Worksheets

    
        Volume = 0
        YearlyStart = ws.Cells(2, 3).Value
        YearlyEnd = 0
        RowCounter = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Volume"
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For r = 2 To lastRow
    
        If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
            Ticker = ws.Cells(r, 1).Value
            Volume = Volume + ws.Cells(r, 7).Value
            YearlyEnd = ws.Cells(r, 6).Value
            ws.Cells(RowCounter, 9).Value = Ticker
            ws.Cells(RowCounter, 10).Value = YearlyEnd - YearlyStart
            If YearlyStart > 0 Then
                ws.Cells(RowCounter, 11).Value = (YearlyEnd - YearlyStart) / YearlyStart
            Else: ws.Cells(RowCounter, 11).Value = 0
            End If
            ws.Cells(RowCounter, 12).Value = Volume
            RowCounter = RowCounter + 1
            Volume = 0
            YearlyStart = ws.Cells(r + 1, 3).Value
            YearlyEnd = 0
    
        Else
    
            Volume = Volume + ws.Cells(r, 7).Value
    
        End If
    
        Next r
        
        
        
    Next ws

End Sub


