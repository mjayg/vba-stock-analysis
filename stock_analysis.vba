Sub alphatesting()

For Each ws In Worksheets

Dim WorksheetName As String
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name

ws.Range("H1").EntireColumn.Insert
ws.Cells(1, 8).Value = "Ticker"

ws.Range("I1").EntireColumn.Insert
ws.Cells(1, 9).Value = "Total"


LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column


Dim current_ticker As String
Dim volume_total As Double
volume_total = 0
Dim total_column As Integer
total_column = 2

For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


current_ticker = ws.Cells(i, 1).Value
volume_total = volume_total + ws.Cells(i, 7).Value

Range("H" & total_column).Value = current_ticker
Range("I" & total_column).Value = volume_total

total_column = total_column + 1

volume_total = 0

Else

volume_total = volume_total + ws.Cells(i, 7).Value


End If



Next i

    
Next ws
    
    
End Sub
