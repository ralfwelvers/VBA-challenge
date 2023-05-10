
Option Explicit

Sub stock_calc()

Dim row_counter As Integer
Dim counter As Integer
Dim row_index As LongLong
Dim output_index As Integer
Dim stock_counter As Integer
Dim stock_ticker As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As LongLong
Dim greatest_percent_increase As Variant
Dim greatest_percent_decrease As Variant
Dim greatest_total_volume As LongLong
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease_ticker As String
Dim greatest_total_volume_ticker As String
Dim look_cell As Range
Dim ws As Worksheet

For Each ws In Worksheets

    Sheets(ws.Name).Select
    
    row_index = 2
    
    stock_counter = Application.WorksheetFunction.CountA(Application.WorksheetFunction.Unique(Range("A2:A1040000"))) - 1
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("J:J").Select
    Selection.NumberFormat = "0.00"
    Columns("K:K").Select
    Selection.NumberFormat = "0.00"
    Columns("L:L").Select
    Selection.NumberFormat = "0"
    Range("H6").Select
    
    For counter = 1 To stock_counter
    
        stock_ticker = Range("A" & row_index).Value
        row_counter = Application.WorksheetFunction.CountIf(Range("A2:A1040000"), stock_ticker) - 1
        
        open_price = Range("C" & row_index).Value
        close_price = Range("F" & row_index + row_counter).Value
        yearly_change = close_price - open_price
        percent_change = yearly_change / open_price
        
        total_volume = Application.WorksheetFunction.Sum(Range("G" & row_index & ":G" & row_index + row_counter))
        Range("I" & counter + 1).Value = stock_ticker
        Range("J" & counter + 1).Value = yearly_change
        Range("K" & counter + 1).Value = percent_change
        Range("L" & counter + 1).Value = total_volume
        
        row_index = row_index + row_counter + 1
    
    Next counter
    
    Dim MyRange As Range
    Set MyRange = Range("J2:J" & stock_counter + 1)
    MyRange.FormatConditions.Delete
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    MyRange.FormatConditions(1).Interior.Color = vbRed
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    MyRange.FormatConditions(2).Interior.Color = vbGreen
    
    Set MyRange = Range("K2:K" & stock_counter + 1)
    MyRange.FormatConditions.Delete
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    MyRange.FormatConditions(1).Interior.Color = vbRed
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    MyRange.FormatConditions(2).Interior.Color = vbGreen
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
     
    greatest_percent_increase = Application.WorksheetFunction.Max(Range("K2:K" & stock_counter + 1))
    Set look_cell = Range("K2:K" & stock_counter + 1).Find(greatest_percent_increase, Lookat:=xlWhole)
    greatest_percent_increase_ticker = look_cell.Offset(, -2)
    
    greatest_percent_decrease = Application.WorksheetFunction.Min(Range("K2" & ":K" & stock_counter + 1))
    Set look_cell = Range("K2:K" & stock_counter + 1).Find(greatest_percent_decrease, Lookat:=xlWhole)
    greatest_percent_decrease_ticker = look_cell.Offset(, -2)
    
    greatest_total_volume = Application.WorksheetFunction.Max(Range("L2" & ":L" & stock_counter + 1))
    Set look_cell = Range("L2:L" & stock_counter + 1).Find(greatest_total_volume, Lookat:=xlWhole)
    greatest_total_volume_ticker = look_cell.Offset(, -3)
    
    Range("Q2").Value = greatest_percent_increase
    Range("Q3").Value = greatest_percent_decrease
    Range("Q4").Value = greatest_total_volume
    
    Range("P2").Value = greatest_percent_increase_ticker
    Range("P3").Value = greatest_percent_decrease_ticker
    Range("P4").Value = greatest_total_volume_ticker
    
    Columns("J:J").Select
    Selection.NumberFormat = "0.00"
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    Columns("L:L").Select
    Selection.NumberFormat = "0"
    
    Range("Q2").Select
    Selection.NumberFormat = "0.00%"
    Range("Q3").Select
    Selection.NumberFormat = "0.00%"
    Range("Q4").Select
    Selection.NumberFormat = "0"
    
    Set MyRange = Range("Q2:Q3")
    MyRange.FormatConditions.Delete
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    MyRange.FormatConditions(1).Interior.Color = vbRed
    MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    MyRange.FormatConditions(2).Interior.Color = vbGreen
    
    Range("H6").Select

Next

MsgBox ("All done")

End Sub




