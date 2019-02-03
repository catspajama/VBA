Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Select

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = " Total Volume"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "percentage change(%)"
Cells(2, 14).Value = "Greatest % change Increase"
Cells(3, 14).Value = "greatest % change Decrease"
Cells(4, 14).Value = "Greatest total volume change"

Dim Brand_Name As String
    
Dim Brand_Total As Double
    Brand_Total = 0
      
Dim Yearly_Change As Double
    Yearly_Change = 0
      
Dim Percentage As Double
    Percentage = 0
    
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

Dim max_vol As Double
    max_vol = 0

Dim max_year As Double
    max_year = 0

Dim min_year As Double
    min_year = 0
 
lastrow = ws.Cells(Rows.Count, 2).End(xlUp).Row
    
For i = 2 To lastrow
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 3).Value > 0 Then
          Brand_Name = Cells(i, 1).Value
          Brand_Total = Brand_Total + Cells(i, 7).Value
         
          Yearly_Change = Cells(i, 6).Value - Cells(i - 260, 3).Value
    
            If Cells(i - 260, 3).Value = 0 Then
               Percentage = 0
               Else: Percentage = (Yearly_Change / Cells(i - 260, 3).Value) * 100
               End If
          Range("I" & Summary_Table_Row).Value = Brand_Name
          Range("J" & Summary_Table_Row).Value = Brand_Total
          Range("K" & Summary_Table_Row).Value = Yearly_Change
          Range("L" & Summary_Table_Row).Value = Percentage
      If Range("K" & Summary_Table_Row).Value >= 0 Then
         Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
         Else
         Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
         Range("L" & Summary_Table_Row).Value = Percentage
         Summary_Table_Row = Summary_Table_Row + 1
         Brand_Total = 0
      Else
         Brand_Total = Brand_Total + Cells(i, 7).Value
      End If
Next i

max_vol = WorksheetFunction.Max(Range("J:J"))
Cells(4, 16).Value = max_vol

max_year = WorksheetFunction.Max(Range("L:L"))
Cells(2, 16).Value = max_year

min_year = WorksheetFunction.Min(Range("L:L"))
Cells(3, 16).Value = min_year

For J = 2 To lastrow
    If Cells(J, 10).Value = max_vol Then
    Cells(4, 15).Value = Cells(J, 9).Value
    End If
    If Cells(J, 12).Value = max_year Then
    Cells(2, 15).Value = Cells(J, 9).Value
    End If
    If Cells(J, 12).Value = min_year Then
    Cells(3, 15).Value = Cells(J, 9).Value
    End If
Next J

  Columns("A:P").EntireColumn.AutoFit
Next ws
End Sub


