Sub M2():

    For Each ws In Worksheets
    
    'Columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Assign Values
    Dim Worksheet As String
    Dim i As Long
    Dim j As Long
    Dim TCount As Long
    Dim LastRow1 As Long
    Dim LastRow2 As Long
    Dim PC As Double
    Dim Increase As Double
    Dim Decrease As Double
    Dim Volume As Double
    
    'Ticker Count
    TCount = 2
    j = 2
    
    'Blank Row
    LastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow1
    
    'Ticker Change
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    ws.Cells(TCount, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(TCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    
    'Formating
    If ws.Cells(TCount, 10).Value > 0 Then
    ws.Cells(TCount, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(TCount, 10).Interior.ColorIndex = 3
    End If
    
    'Percent Change
    If ws.Cells(j, 3).Value <> 0 Then
    PC = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
    ws.Cells(TCount, 11).Value = Format(PC, "Percent")
    
    
    End If
    
    'Volume Calculations
    ws.Cells(TCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
    
    'TCount Setup
    TCount = TCount + 1
    j = i + 1
    End If
    Next i
    
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Volume = ws.Cells(2, 12).Value
    Increase = ws.Cells(2, 11).Value
    Decrease = ws.Cells(2, 11).Value
    
    For i = 2 To LastRow1
    
    If ws.Cells(i, 12).Value > Volume Then
    Volume = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    End If
    
    If ws.Cells(i, 11) > Increase Then
    Increase = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value * 100
    End If
    
    
    If ws.Cells(i, 11) < Decrease Then
    Decrease = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value * 100
    End If
    
    ws.Cells(2, 17).Value = Format(Increase, "Percent")
    ws.Cells(3, 17).Value = Format(Decrease, "Percent")

Next i
Next ws

    For Each WS In ActiveWorkbook.Worksheets
    WS.Columns.AutoFit

Next WS

End Sub
