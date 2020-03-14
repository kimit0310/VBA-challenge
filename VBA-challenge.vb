Sub Pricechange()
    Dim yearChange, percentChange, totalVolume, openPrice, closePrice, lastrow As Double
    Dim ticker      As String
    For Each ws In Worksheets
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Volume"
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        totalVolume = 0
        Dim j As Double
        j = 2
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                ticker = ws.Cells(i, 1).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ws.Range("I" & j).Value = ticker
                ws.Range("L" & j).Value = totalVolume
                totalVolume = 0
                closePrice = ws.Cells(i, 6)
                If openPrice = 0 Then
                    percentChange = 0
                    yearChange = 0
                Else:
                    percentChange = (closePrice - openPrice) / openPrice
                    yearChange = closePrice - openPrice
                End If
                j = j + 1
                ws.Range("J" & j).Value = yearChange
                ws.Range("K" & j).Value = percentChange
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                openPrice = ws.Cells(i, 3)
            Else: totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        For i = 2 To lastrow
            If ws.Range("J" & i).Value > 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & i).Value < 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
End Sub
