Sub stockprice()

'Defining Variables
Dim TickerName As String
Dim Tickerrow As Integer
Dim TickerCount As Long
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim lastrow As Long
Dim ws As Worksheet
Dim Summary_lastrow As Long
Dim m As Long

'prevents overflow error
On Error Resume Next

'looping through sheets
For Each ws In ThisWorkbook.Worksheets

'Summary Headers
ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
ws.Range("o2").Value = "Greatest % increase"
ws.Range("o3").Value = "Greatest % decrease"
ws.Range("o4").Value = "Greatest Total Volume"

'Header Names
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percentage Change"
ws.Range("l1").Value = "Total Stock Volume"

TotalVolume = 0
Tickerrow = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Columns("K").NumberFormat = "0.00%"
Summary_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
ws.Range("Q2:q3").NumberFormat = "0.00%"

'looping through rows
For i = 2 To lastrow
   

'Conditionals
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'finding values
    Ticker = ws.Cells(i, 1).Value
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    ClosingPrice = ws.Cells(i, 6).Value
    TickerCount = WorksheetFunction.CountIf(ws.Range("a:a"), Ticker)
    OpeningPrice = ws.Cells((i - TickerCount + 1), 3).Value
    YearlyChange = ClosingPrice - OpeningPrice
    PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
    
    'assigning values
    ws.Range("i" & Tickerrow).Value = Ticker
    ws.Range("l" & Tickerrow).Value = TotalVolume
    ws.Range("j" & Tickerrow).Value = YearlyChange
    ws.Range("k" & Tickerrow).Value = PercentChange

    'resetting
    Tickerrow = Tickerrow + 1
    TotalVolume = 0
    TickerCount = 0
Else
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value

End If
Next i


'conditional formatting
For m = 2 To WorksheetFunction.CountA(ws.Range("i:i"))

If ws.Cells(m, 10) >= 0 Then
    ws.Cells(m, 10).Interior.ColorIndex = 4
Else
    ws.Cells(m, 10).Interior.ColorIndex = 3

End If
Next m

'Conditionals for finding max and min
Dim g As Long
Dim ValueRange As Range
Dim VolumeRange As Range

Set ValueRange = ws.Range("K:k")
Set VolumeRange = ws.Range("L:L")


ws.Range("q2").Value = WorksheetFunction.Max(ValueRange)
ws.Range("q3").Value = WorksheetFunction.Min(ValueRange)
ws.Range("q4").Value = WorksheetFunction.Max(VolumeRange)

For g = 2 To WorksheetFunction.CountA(ws.Range("i:i"))

If ws.Range("q2").Value = ws.Cells(g, 11).Value Then
    ws.Range("p2").Value = ws.Cells(g, 9).Value
End If
If ws.Range("q3").Value = ws.Cells(g, 11).Value Then
    ws.Range("p3").Value = ws.Cells(g, 9).Value
End If
If ws.Range("q4").Value = ws.Cells(g, 12).Value Then
    ws.Range("p4").Value = ws.Cells(g, 9).Value
End If

Next g
ws.Columns("A:Q").AutoFit
Next ws
End Sub




