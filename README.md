'# 3YearsStockAnaylsis_VBA
'analysis of multiple year stock data via VBA


Sub stock_analysis()
  'set dimension
 
 Dim Total As Double
 Dim RowIndex As Long
 Dim Change As Double
 Dim ColumnIndex As Integer
 Dim Start As Long
 Dim RowCount As Long
 Dim PercentChange As Double
 Dim Days As Integer
 Dim DailyChange As Double
 Dim AverageChange As Double
 Dim ws As Worksheet
 Dim increase_number As Long
 Dim decrease_number As Long
 Dim volume_number As Long
 
 
For Each ws In Worksheets
  ColumnIndex = 0
  Total = 0
  Change = 0
  Start = 2
  DailyChange = 0
  
  'set title row
  
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  ws.Range("O2").Value = "Greated % increase"
  ws.Range("O3").Value = "Greated % decrease"
  ws.Range("O4").Value = "Greated Total Volume"

'get the last row number

RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

For RowIndex = 2 To RowCount

'if Ticker changes, print the data

If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
Total = Total + ws.Cells(RowIndex, 7).Value

If Total = 0 Then
   ws.Range("I" & 2 + ColumnIndex).Value = ws.Cells(RowIndex, 1).Value
   ws.Range("J" & 2 + ColumnIndex).Value = 0
   ws.Range("K" & 2 + ColumnIndex).Value = "%" & 0
   ws.Range("L" & 2 + ColumnIndex).Value = 0
   
Else
  If ws.Cells(Start, 3) = 0 Then
    For find_value = Start To RowIndex
    If ws.Cells(find_value, 3).Value <> 0 Then
    Start = find_value
    Exit For
    End If
    Next find_value
End If

'calculate the yearly change and percent change of stock prices

Change = (ws.Cells(RowIndex, 6) - ws.Cells(Start, 3))
PercentChange = Change / ws.Cells(Start, 3)

Start = RowIndex + 1
'color the cells based on whether the stock increased or decreased in value

ws.Range("I" & 2 + ColumnIndex) = ws.Cells(RowIndex, 1).Value
ws.Range("J" & 2 + ColumnIndex) = Change
ws.Range("J" & 2 + ColumnIndex).NumberFormat = "0.00'"
ws.Range("K" & 2 + ColumnIndex).Value = PercentChange
ws.Range("K" & 2 + ColumnIndex).NumberFormat = "0.00%"
ws.Range("L" & 2 + ColumnIndex).Value = Total


Select Case Change
  Case Is > 0
   ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 4
  Case Is < 0
  ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 3
  Case Else
  ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 0
  End Select
  
End If

  Total = 0
  Change = 0
  ColumnIndex = ColumnIndex + 1
  Days = 0
  DailyChange = 0
  
  Else
  Total = Total + ws.Cells(RowIndex, 7).Value
  
  End If
  
  Next RowIndex
  
  'find the max and min of percent change and max value and print the row
  
  ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
  ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("k2:k" & RowCount)) * 100
  ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
  
  increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("k2:k" & RowCount), 0)
  decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("k2:k" & RowCount), 0)
  volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
  
  ws.Range("P2") = ws.Cells(increase_number + 1, 9)
  ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
  ws.Range("P4") = ws.Cells(volume_number + 1, 9)
  
  Next ws
    
  End Sub
  
  
  
  


