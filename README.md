# Module-2-Challenge
Sub Alphabetical_Testing()
    Dim ws As Worksheet
    For Each ws In Worksheets
    ' Create variables to hold the value of the tickers
    Dim first_column_yearly_change As Double
    Dim second_column_percent_change As Double
    Dim third_column_total_stock As Double
    Dim tickerIndex As Double
    Dim total As Double
    Dim rowCount As Long
    Dim J As Double
    J = 0
    Dim orow As Double
    orow = 2
    
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "total volume"
    
    
    tickerIndex = 1
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To rowCount
        Value = ws.Cells(i, 1).Value
        Value2 = ws.Cells(i + 1, 1).Value
        ' If the next row does not equal this row then it needs to be in a different group
        If Value <> Value2 Then
        total = total + ws.Cells(i, 7).Value
        ' If total = 0 Then
        
        
        
        Change = ws.Cells(i, 6).Value - ws.Cells(orow, 3).Value
        PercentChange = Change / ws.Cells(orow, 3).Value
            ' print the results
            ws.Range("I" & 2 + J).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + J).Value = Change
            ws.Range("J" & 2 + J).NumberFormat = "0.00"
            ws.Range("K" & 2 + J).Value = PercentChange
            ws.Range("K" & 2 + J).NumberFormat = "0.00%"
            ws.Range("L" & 2 + J).Value = total
            

            
            If ws.Range("J" & 2 + J).Value > 0 Then
            ws.Range("J" & 2 + J).Interior.ColorIndex = 4
            End If
            If ws.Range("J" & 2 + J).Value < 0 Then
            ws.Range("J" & 2 + J).Interior.ColorIndex = 3
            
            
           End If
            total = 0
            J = J + 1
            orow = i + 1
            Change = 0
        Else
        total = total + ws.Cells(i, 7).Value
    End If
    Next i
    Next ws
End Sub
