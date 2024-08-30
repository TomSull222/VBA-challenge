Attribute VB_Name = "Module1"
Sub Stock_Data()

Dim i As Long
Dim j As Integer
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim WorksheetName As String
Dim Quarter As String
WorksheetName = ws.Name


Dim ticker As String
Dim quarterly_change As Double
Dim percent_change As Double
Dim open_price As Double
Dim close_price As Double
Dim stockvolume As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quarterly_Change"
Cells(1, 11).Value = "Percent_Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim TickerRow As Long
TickerRow = 2
quarterly_change = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
ticker = Cells(i, 1).Value
stockvolume = stockvolume + Cells(i, 7).Value
Range("i" & TickerRow).Value = ticker
Range("l" & TickerRow).Value = stockvolume
open_price = Cells(i, 3).Value
close_price = Cells(i, 6).Value
quarterly_change = (open_price - close_price)
Range("j" & TickerRow).Value = quarterly_change

If (open_price = 0) Then
percent_change = 0
Else
percent_change = (quarterly_change / open_price)
End If

Range("k" & TickerRow).Value = percent_change
Range("k" & TickerRow).NumberFormat = "0.00%"
TickerRow = TickerRow + 1
stockvolume = 0
open_price = Cells(i + 1, 3)
Else
stockvolume = stockvolume + Cells(i, 7).Value
End If
Next i

lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To lastrow_summary_table

If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
Else
Cells(i, 10).Interior.ColorIndex = 3
End If
Next i

Cells(2, 15).Value = "Greatest % Increase"
Cells(2, 15).Value = "Greatest % Decrease"
Cells(2, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For i = 2 To lastrow_summary_table
If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("k2:k" & lastrow_summary_table)) Then
Cells(2, 16).Value = Cells(i, 9).Value
Cells(2, 17).Value = Cells(i, 11).Value
Cells(2, 17).NumberFormat = "0.00%"

ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("k2:k" & lastrow_summary_table)) Then
Cells(3, 16).Value = Cells(i, 9).Value
Cells(3, 17).Value = Cells(i, 11).Value
Cells(3, 17).NumberFormat = "0.00%"

ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("l2:l" & lastrow_summary_table)) Then
Cells(4, 16).Value = Cells(i, 9).Value
Cells(4, 17).Value = Cells(i, 12).Value
End If
Next i

Next ws

End Sub

