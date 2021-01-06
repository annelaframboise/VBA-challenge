Sub stock_data()

For Each ws In Worksheets
    Dim TickerName As String
    Dim TickerTotal As Double
    TickerTotal = 0
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    Dim Openings As Double
    Openings = 0
    Dim Closings As Double
    Closings = 0
    Dim YearChange As Double
    Dim PercentChange As Double
        
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker Name"
    ws.Range("J1").Value = "Yearly Change from Opening to Closing of Year"
    ws.Range("K1").Value = "Percent Change from Opening to Closing of Year"
    ws.Range("L1").Value = "Ticker Total"


      For i = 2 To LastRow
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TickerName = ws.Cells(i, 1).Value
            TickerTotal = TickerTotal + ws.Cells(i, 7).Value
            Closings = ws.Cells(i, 6).Value
            ws.Range("I" & Summary_Table_Row).Value = TickerName
            ws.Range("L" & Summary_Table_Row).Value = TickerTotal
            YearChange = Closings - Openings
            ws.Range("J" & Summary_Table_Row).Value = YearChange
            If YearChange < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
               ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            If Openings > 0 Then
                PercentChange = (YearChange / Openings) * 100
            Else
                PercentChange = 0
            End If
            ws.Range("K" & Summary_Table_Row).Value = PercentChange
            TickerTotal = 0
            Summary_Table_Row = Summary_Table_Row + 1
          ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Openings = ws.Cells(i, 3).Value
          Else
            TickerTotal = TickerTotal + Cells(i, 7).Value
           End If
                    
  

      Next i
  

      
  Next ws

End Sub
    
    
