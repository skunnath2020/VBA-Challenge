Sub VBA_Challenge()

Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
'*******************************************************************
'Used the Advanced filters to get unique Tickers and then based on
'those tickers created 2 loops
'********************************************************************
'Unique Tickers
Dim RowCount, TickerCount As Long
Dim WorksheetName As String
WorksheetName = ws.Name
'MsgBox (WorksheetName)
'get unique tickers and place them on to column I
 'clear contents of the cells
   ws.Range("I:I").ClearContents
   ws.Range("A1").Select
   ws.Range(Selection, Selection.End(xlDown)).Select
   Selection.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True

'find last row of data
 RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
 'MsgBox "Total row count :" & rowCount - 1
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Stock Volume"
 TickerCount = ws.Cells(Rows.Count, 9).End(xlUp).Row
 'Clear_cells (TickerCount)
 'MsgBox "Total Ticker count :" & TickerCount



''''''''''''''''''''''''''''''''''''''''''''''''''
'Get Open_price, Close_price, Yearly change and Total Vol
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim pos As Integer
Dim ticker_position, r, c As Long
Dim Open_price, Close_price, Stock_vol, Yearly_change, Percent_change As Double

'Initialize to ticker position COLUMN value
pos = 1
ticker_position = 2
For r = 2 To TickerCount
'Init variables
   Stock_vol = Open_price = Close_price = Percent_change = 0
   Open_price = ws.Cells(ticker_position, pos + 2).Value
   For c = ticker_position To RowCount
    Close_price = ws.Cells(c, pos + 5).Value
    Stock_vol = Stock_vol + ws.Cells(c, pos + 6)
    'check condition for ticker change and print
    If ws.Cells(r, 9).Value <> ws.Cells(c + 1, pos).Value Then
      Close_price = ws.Cells(c, pos + 5).Value
     'Calculate print and format Annual change
      Yearly_change = Close_price - Open_price
      ws.Cells(r, 10).Value = Yearly_change 'Yearly change
      'color code Yearly change based on +/-
      If Yearly_change < 0 Then
        ws.Cells(r, 10).Interior.ColorIndex = 3
      Else
        ws.Cells(r, 10).Interior.ColorIndex = 4
      End If
      'Calculate Percentage and print with no formating
      If Open_price <> 0 Then
        Percent_change = (Yearly_change / Open_price) * 100
        ws.Cells(r, 11).Value = Round((Percent_change), 2)
      Else
        ws.Cells(r, 11).Value = 0
      End If
      'Calculate and print total stock volume
      ws.Cells(r, 12).Value = Stock_vol
      'Preserve the column position
      ticker_position = c + 1
      'Our c-iterator pick up the ticker from range(I) and loops till the end of all rows.
      'We don't want that to happen once next ticker in col A is encountered - hence
      'exit and return to next ticker in col I
      Exit For
    End If
   Next c
   'Format percentage
   ws.Cells(r, 11).NumberFormat = "0.00\%"
  Next r
  ' Bonus section for tickers with aggregates
  Bonus_totals
  
Next ws
MsgBox ("Processing Complete!!!")
End Sub

Sub Bonus_totals()
'Declare Variables
Dim rng As Range
Dim Max_percent, Min_percent, MaxVol As Double
Dim MaxCell, MinCell, VolCell As Range
'Calculate Values
Max_percent = Application.WorksheetFunction.Max(Range("K:K"))
Min_percent = Application.WorksheetFunction.Min(Range("K:K"))
MaxVol = Application.WorksheetFunction.Max(Range("L:L"))
Set MaxCell = Range("K:K").Find(Max_percent, Lookat:=xlWhole)
Set MinCell = Range("K:K").Find(Min_percent, Lookat:=xlWhole)
Set VolCell = Range("L:L").Find(MaxVol, Lookat:=xlWhole)
'Display Values
Range("P2") = "Ticker"
Range("Q2") = "Value"
Range("O3") = "Greatest % Increase"
Range("O3").Columns.AutoFit
Range("P3") = MaxCell.Offset(, -2)
Range("Q3") = MaxCell.Offset(, 0)
Range("O4") = "Greatest % Decrease"
Range("O4").Columns.AutoFit
Range("P4") = MinCell.Offset(, -2)
Range("Q4") = MinCell.Offset(, 0)
Range("Q3:Q4").NumberFormat = "0.00\%"
Range("O5") = "Greatest Total Volume"
Range("O5").Columns.AutoFit
Range("P5") = VolCell.Offset(, -3)
Range("Q5") = VolCell.Offset(, 0)
End Sub
Sub Clear_cells(Ticker_count As Long)
 Range("J2:L" & Ticker_count).ClearContents
 Range("O2:P5").ClearContents
 Range("J:J").ClearFormats
End Sub


