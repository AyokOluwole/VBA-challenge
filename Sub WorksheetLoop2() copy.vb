Sub WorksheetLoop2()

         ' Declare Current as a worksheet object variable.
         Dim Current As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets
        ' Dim ws As Worksheet
        ' Set ws = wb.Sheets("B")
        ' ws.Activate
      Current.Activate

            ' Insert your code here.
            ' This line displays the worksheet name in a message box.
     
    Dim total_volume As Double
    Dim ticker As String
    Dim ticker_counter, ticker_open_close_counter As Double
    Dim yearly_open, yearly_close As Double
    Range("J1").Value = "Yearly Change"
    total_volume = 0
    ticker_counter = 2
    ticker_open_close_counter = 2
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow

        total_volume = total_volume + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(ticker_open_close_counter, 3)
        
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearly_close = Cells(i, 6)
            Cells(ticker_counter, 9).Value = ticker
            Cells(ticker_counter, 10).Value = yearly_close - yearly_open
                    
                 

        
            If yearly_open = 0 Then
                Cells(ticker_counter, 11).Value = Null
                Else
                Cells(ticker_counter, 11).Value = (yearly_close - yearly_open) / yearly_open
            End If
            Cells(ticker_counter, 12).Value = total_volume
            
          
            Cells(ticker_counter, 11).NumberFormat = "0.00%"
            
            
            total_volume = 0
            ticker_counter = ticker_counter + 1
            ticker_open_close_counter = i + 1
            
        End If


    Next i
    lastfork = Cells(Rows.Count, 11).End(xlUp).Row
        'MsgBox (lastfork)
   For j = 2 To lastfork
           If Cells(j, 11).Value >= 0 Then
           Cells(j, 11).Interior.ColorIndex = 4
          ElseIf Cells(j, 11).Value < 0 Then
          Cells(j, 11).Interior.ColorIndex = 3
           End If
           Next j
    
    
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))

  
    highest = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    lowest = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    highest_vol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)

 
    Range("P2") = Cells(highest + 1, 9)
    Range("P3") = Cells(lowest + 1, 9)
    Range("P4") = Cells(highest_vol + 1, 9)


       'MsgBox Current.Name
         Next


End Sub







