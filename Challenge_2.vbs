Sub Ticker()

    For Each ws In Worksheets

        Dim r As Double
        r = 2
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Range("I1") = "Ticker"
    
        ' identify all stocks
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(r, 9).Value = ws.Cells(i, 1).Value
                r = r + 1
            End If
        Next i
         
    Next ws

End Sub

Sub TotalSum()

    For Each ws In Worksheets

        Dim Sum As Double
        Dim r As Double
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Range("L1") = "Total Stock Volume"
    
        r = 2
        Sum = 0
    
        ' if next row's stock equals current row, add volume to the summed volume
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                    Sum = Sum + ws.Cells(i, 7).Value
                ' else add that row's colume to the sum to get final stock sum
                ' then reset sum to zero and move on to next stock
                Else
                    Sum = Sum + ws.Cells(i, 7).Value
                    ws.Cells(r, 12).Value = Sum
                    Sum = 0
                    r = r + 1
            End If

        Next i
    
    Next ws

End Sub

Sub tickerchal()

        For Each ws In Worksheets

            LastRow = Cells(Rows.Count, 1).End(xlUp).Row

            Dim o As Double
            Dim c As Double
            Dim r As Double

    
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
    
            o = 0
            c = 0
            r = 2
 
            For i = 2 To LastRow
                ' if row's date is same as first date then set as opening value
                If ws.Cells(i, 2).Value = ws.Range("B2") Then
                    o = ws.Cells(i, 3).Value
                End If
                ' if the next row' stock does not equal current row,
                ' then set that row's closing price to closing value
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    c = ws.Cells(i, 6).Value
                End If
                ' if c and o were set values based on the ifs statements above,
                ' then find the yearly change and percent change
                If o <> 0 And c <> 0 Then
                    ws.Cells(r, 10).Value = c - o
                    ws.Cells(r, 11).Value = (c - o) / o
                    o = 0
                    c = 0
                    r = r + 1
                End If
        
            Next i
            
            ' set red if yearly change decreased, set green if yearly change increased
            For r = 2 To 3001
                If ws.Cells(r, 10).Value < 0 Then
                    ws.Cells(r, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(r, 10).Interior.ColorIndex = 4
                End If
        
            Next r

        Next ws
    
End Sub

Sub summary()

    For Each ws In Worksheets

        LastRow = Cells(Rows.Count, 9).End(xlUp).Row

        Dim m As Double
        Dim t As String
        Dim d As Double
        Dim v As Double
    
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
    
        m = 0
        
        ' find greatest percent increase
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value > m Then
                m = ws.Cells(i, 11).Value
                t = ws.Cells(i, 9).Value
            End If
        Next i
        ws.Cells(2, 16).Value = t
        ws.Cells(2, 17).Value = m
    
        ' find greatest percent decrease
        d = 0
        For i = 2 To LastRow
            If ws.Cells(i, 11).Value < d Then
                d = ws.Cells(i, 11).Value
                t = ws.Cells(i, 9).Value
            End If
        Next i
        ws.Cells(3, 16).Value = t
        ws.Cells(3, 17).Value = d
    
        ' find greatest total volume
        v = 0
        For i = 2 To LastRow
            If ws.Cells(i, 12).Value > v Then
                v = ws.Cells(i, 12).Value
                t = ws.Cells(i, 9).Value
            End If
        Next i
        ws.Cells(4, 16).Value = t
        ws.Cells(4, 17).Value = v

    Next ws
    
End Sub