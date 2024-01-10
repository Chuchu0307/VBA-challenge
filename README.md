Sub Ticker()
    For Each WS In Worksheets
    
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
           OpenPrice = WS.Cells(2, 3).Value
    Dim YearlyChange As Double
    
    Dim SummaryTableRow As Long
           SummaryTableRow = 2
    Dim TotalStockVolume As Double
           TotalStockVolume = 0

 
    For i = 2 To LastRow
    
    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    Ticker = WS.Cells(i, 1).Value
    ClosePrice = WS.Cells(i, 6).Value
    YearlyChange = ClosePrice - OpenPrice
    WS.Range("J" & SummaryTableRow).Value = YearlyChange

      
       If OpenPrice <> 0 Then
       WS.Range("K" & SummaryTableRow).Value = YearlyChange / OpenPrice
       End If
       
    OpenPrice = WS.Cells(i + 1, 3).Value
    WS.Range("I" & SummaryTableRow).Value = Ticker
    TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
    WS.Range("L" & SummaryTableRow).Value = TotalStockVolume
    SummaryTableRow = SummaryTableRow + 1
    TotalStockVolume = 0
       
    Else
       TotalStockVolume = TotalStockVolume + WS.Cells(i, 7).Value
    
    End If
    
    If WS.Cells(i, 11).Value = Application.WorksheetFunction.Max(WS.Range("K:K")) Then
    
    WS.Range("P2").Value = Application.WorksheetFunction.Max(WS.Range("K:K"))
    WS.Range("O2").Value = WS.Cells(i, 9).Value
    
    ElseIf WS.Cells(i, 11).Value = Application.WorksheetFunction.Min(WS.Range("K:K")) Then
    
    WS.Range("P3").Value = Application.WorksheetFunction.Min(WS.Range("K:K"))
    WS.Range("O3").Value = WS.Cells(i, 9).Value
    
    ElseIf WS.Cells(i, 12).Value = Application.WorksheetFunction.Max(WS.Range("L:L")) Then
    WS.Range("P4").Value = Application.WorksheetFunction.Max(WS.Range("L:L"))
    WS.Range("O4").Value = WS.Cells(i, 9).Value
    
    End If
    
    If WS.Cells(i, 10).Value > 0 Then
    WS.Cells(i, 10).Interior.ColorIndex = 4
    Else
    WS.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
    Next i
    
    Next WS
    
End Sub
