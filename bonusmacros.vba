Sub stockdata()

For Each ws In Worksheets

   Dim ticker As String
   Dim total As Double
   Dim change As Double
   Dim percentage As Double
   Dim tablerow As Integer
   Dim start As Double

   
   total = 0
   change = 0
   tablerow = 2
   start = 0
  
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = " Total Stock Volumn"
    
    For i = 2 To lastrow
     
      If i = 2 Then
      start = ws.Cells(i, 3).Value
      End If
     
     If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And i <> 2 Then
     
     ticker = ws.Cells(i - 1, 1).Value
     
     ws.Range("I" & tablerow).Value = ticker
     
     ws.Range("L" & tablerow).Value = total
     
     change = (Cells(i - 1, 6) - start)
     percentage = change / start
     start = ws.Cells(i, 3).Value
     ws.Range("J" & tablerow).Value = change
     ws.Range("K" & tablerow).Value = percentage
     ws.Range("K" & tablerow).NumberFormat = "0.00%"
     
     
     
    If change >= 0 Then
            ws.Cells(tablerow, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(tablerow, 10).Interior.ColorIndex = 3
            End If
    tablerow = tablerow + 1
    total = 0
    total = total + ws.Cells(i, 7).Value
    Else
         total = total + ws.Cells(i, 7).Value
     
     End If
   
     Next i
     ws.Range("O2") = " Greatest%Increase"
    ws.Range("O3") = " Greatest%Decrease"
    ws.Range("O4") = " Greatest Total Volumn"
    ws.Range("P1") = " Ticker"
    ws.Range("Q1") = " Value"
     ws.Range("Q2") = WorksheetFunction.Max(Range("K1:K" & tablerow))
     ws.Range("Q3") = WorksheetFunction.Min(Range("K1:K" & tablerow))
     ws.Range("Q2").NumberFormat = "0.00%"
     ws.Range("Q3").NumberFormat = "0.00%"
     ws.Range("Q4") = WorksheetFunction.Max(Range("L1:L" & tablerow))
     
     
       inc_per_row = WorksheetFunction.Match(ws.Range("Q2"), Range("K1:K" & tablerow), 0)
       dec_per_row = WorksheetFunction.Match(ws.Range("Q3"), Range("K1:K" & tablerow), 0)
       inc_vol_row = WorksheetFunction.Match(ws.Range("Q4"), Range("L1:L" & tablerow), 0)
    
    ws.Range("P2") = ws.Range("I" & inc_per_row)
    ws.Range("P3") = ws.Range("I" & dec_per_row)
    ws.Range("P4") = ws.Range("I" & inc_vol_row)
  
  Next ws
     
     
End Sub