
Sub stock()

Dim ws As Worksheet

Dim ticker As String

Dim stock_volume As Double
stock_volume = 0

Dim starting_value As Double
starting_value = 0

Dim ending_value As Double
ending_value = 0

Dim yearly_change As Double
yearly_change = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In Worksheets
    ws.Activate
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Stock Volume"
    
    Summary_Table_Row = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value

         stock_volume = stock_volume + Cells(i, 7).Value

         Range("I" & Summary_Table_Row).Value = ticker

         Range("L" & Summary_Table_Row).Value = stock_volume

         Summary_Table_Row = Summary_Table_Row + 1
      
            stock_volume = 0
        
            If Cells(i, 2).Value = 20160101 Then
            starting_value = Cells(i, 3).Value
            Else: starting_value = starting_value
            End If
            
            If Cells(i, 2).Value = 20161230 Then
            ending_value = Cells(i, 6).Value
            Else: ending_value = ending_value
            End If
            
            If starting_value > 0 And ending_value > 0 Then
            yearly_change = starting_value / ending_value
            End If
            
            If yearly_change > 0 Then
            Range("J" & Summary_Table_Row).Value = yearly_change
            End If
            

            

    Else

      stock_volume = stock_volume + Cells(i, 7).Value

    End If

  Next i
Next

End Sub