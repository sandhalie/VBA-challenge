Sub challenge():
Dim ws As Worksheet
For Each ws In Sheets

Dim count As Integer
' select the unique tickers and paste it in column I

ws.Range("A1").EntireColumn.Copy ws.Range("I1")
ws.Columns("I").RemoveDuplicates Columns:=1, Header:=xlYes

'format date column

 With ws.Range("B:B")
  .NumberFormat = "General"
  .Value = .Value
 End With
 
' add the column headers

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Opening Price"
ws.Range("K1").Value = "Closing Price"
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percentage Change"
ws.Range("N1").Value = "Total Stock Volume"

Next ws
End Sub

Sub challenge2():
Dim ws As Worksheet
Dim iter As Integer
Dim skip As Integer
For Each ws In Sheets
' loop through the data to find the opening price and closing price for each ticker


opening_date = Application.WorksheetFunction.Min(ws.Range("B:B"))
closing_date = Application.WorksheetFunction.Max(ws.Range("B:B"))
numrowA = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.count
numrowI = ws.Range("I1", ws.Range("I1").End(xlDown)).Rows.count

' calculate skip size

For j = 3 To 365
    If Range("B" & j).Value = Range("B" & 2).Value Then
    skip = j - 2
    End If
Next j
MsgBox (skip)

For j = 2 To numrowA Step skip
    iter = 2
    Do Until ws.Range("I" & iter).Value = ws.Range("A" & j).Value
        iter = iter + 1
    Loop
    ws.Range("J" & iter).Value = ws.Range("C" & j).Value
Next j

For j = skip + 1 To numrowA Step skip
    iter = 2
    Do Until ws.Range("I" & iter).Value = ws.Range("A" & j).Value
        iter = iter + 1
    Loop
    ws.Range("K" & iter).Value = ws.Range("F" & j).Value
Next j
Next ws
End Sub



Sub challenge3():
Dim ws As Worksheet

For Each ws In Sheets
numrowI = ws.Range("I1", ws.Range("I1").End(xlDown)).Rows.count
' calculate the yearly change from: closing - opening price & format

For i = 2 To numrowI
    ws.Range("L" & i).Value = ws.Range("K" & i).Value - ws.Range("J" & i).Value
     ws.Range("L" & i).NumberFormat = "$#,##0.00"
    If ws.Range("L" & i).Value >= 0 Then
    ws.Range("L" & i).Interior.ColorIndex = 4
    Else: ws.Range("L" & i).Interior.ColorIndex = 3
    End If
Next i


' calculate the percentage total: = yearly change / opening price & format

For i = 2 To numrowI
    ws.Range("M" & i).Value = ws.Range("L" & i).Value / ws.Range("J" & i).Value
    ws.Range("M" & i).Value = Format(ws.Range("M" & i).Value, "0.00%")
    If ws.Range("M" & i).Value >= 0 Then
    ws.Range("M" & i).Interior.ColorIndex = 4
    Else: ws.Range("M" & i).Interior.ColorIndex = 3
    End If
Next i


' loop through the data for each ticker to find the total volume

For i = 2 To numrowI
    ws.Range("n" & i).Formula = "=SUMIF(C[-13],RC[-5],C[-7])"
Next i

' opening and closing price columns and format columns
ws.Columns([10]).EntireColumn.Delete
ws.Columns([10]).EntireColumn.Delete
ws.Columns("I:L").AutoFit

Next ws
End Sub



