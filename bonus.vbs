Sub bonus():
' Loop through the worksheets

Dim ws As Worksheet
For Each ws In Sheets

' Label the rows/ columns

ws.Range("Q2").Value = "Greatest % Increase"
ws.Range("Q3").Value = "Greatest % Decrease"
ws.Range("Q4").Value = "Greatest Total Volume"
ws.Range("R1").Value = "Ticker"
ws.Range("S1").Value = "Value"

' Loop through the list to find the values required

numrowI = ws.Range("I1", ws.Range("I1").End(xlDown)).Rows.Count

For i = 2 To numrowI
    If ws.Range("K" & i) = Application.WorksheetFunction.Max(ws.Range("K:K")) Then
    ws.Range("R2").Value = ws.Range("I" & i).Value
    ws.Range("S2").Value = ws.Range("K" & i).Value
    ElseIf ws.Range("K" & i) = Application.WorksheetFunction.Min(ws.Range("K:K")) Then
    ws.Range("R3").Value = ws.Range("I" & i).Value
    ws.Range("S3").Value = ws.Range("K" & i).Value
    ElseIf ws.Range("L" & i) = Application.WorksheetFunction.Max(ws.Range("L:L")) Then
    ws.Range("R4").Value = ws.Range("I" & i).Value
    ws.Range("S4").Value = ws.Range("L" & i).Value
    End If
Next i

' Format the values

ws.Range("S2").Value = Format(ws.Range("S2").Value, "0.00%")
ws.Range("S3").Value = Format(ws.Range("S3").Value, "0.00%")

Next ws
End Sub
