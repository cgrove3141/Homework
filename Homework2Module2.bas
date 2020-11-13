Attribute VB_Name = "Module2"
Sub Standouts()
    
    For Each ws In Worksheets
    
        biggestIncrease = 0
        biggestLoss = 0
        biggestVolume = 0
        biggestIncreaseTicker = 0
        biggestLossTicker = 0
        biggestVolumeTicker = 0
        
        
        lastStock = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For r = 2 To lastStock
            If ws.Cells(r, 11).Value > biggestIncrease Then
                biggestIncrease = ws.Cells(r, 11).Value
                biggestIncreaseTicker = ws.Cells(r, 9).Value
            End If
            
            If ws.Cells(r, 11).Value < biggestLoss Then
                biggestLoss = ws.Cells(r, 11).Value
                biggestLossTicker = ws.Cells(r, 9).Value
            End If
            
            If ws.Cells(r, 12).Value > biggestVolume Then
                biggestVolume = ws.Cells(r, 12).Value
                biggestVolumeTicker = ws.Cells(r, 9).Value
            End If
            
        Next r
            
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Loss"
        ws.Cells(4, 15).Value = "Largest Volume"
        ws.Cells(2, 16).Value = biggestIncreaseTicker
        ws.Cells(3, 16).Value = biggestLossTicker
        ws.Cells(4, 16).Value = biggestVolumeTicker
        ws.Cells(2, 17).Value = biggestIncrease * 100
        ws.Cells(3, 17).Value = biggestLoss * 100
        ws.Cells(4, 17).Value = biggestVolume
    
    Next ws

End Sub
