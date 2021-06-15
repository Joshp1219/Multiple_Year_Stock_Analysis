Attribute VB_Name = "Module1"
Sub Stocks()
For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim Volume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim summarytablerow As Integer

Volume = 0
YearlyChange = 0
PercentChange = 0
summarytablerow = 2

    For p = 2 To LastRow
        If ws.Cells(p + 1, 1).Value <> ws.Cells(p, 1).Value Then
            Ticker = ws.Cells(p, 1).Value
            Volume = Volume + ws.Cells(p, 7).Value
            YearlyChange = YearlyChange + (ws.Cells(p, 6).Value - ws.Cells(p, 3).Value)
            PercentChange = PercentChange + ((ws.Cells(p, 6).Value - ws.Cells(p, 3).Value) / ws.Cells(p, 3).Value)
            ws.Range("I" & summarytablerow).Value = Ticker
            ws.Range("J" & summarytablerow).Value = YearlyChange
            ws.Range("K" & summarytablerow).Value = PercentChange
            ws.Range("L" & summarytablerow).Value = Volume
            summarytablerow = summarytablerow + 1
            Volume = 0
            YearlyChange = 0
            PercentChange = 0
        Else
            Volume = Volume + ws.Cells(p, 7).Value
            YearlyChange = YearlyChange + (ws.Cells(p, 6).Value - ws.Cells(p, 3).Value)
            PercentChange = PercentChange + ((ws.Cells(p, 6).Value - ws.Cells(p, 3).Value) / ws.Cells(p, 3).Value)
        End If
       Next p
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        For j = 2 To LastRow
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
    'Bonus
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L:L"))
        For q = 1 To LastRow
            If ws.Cells(q, 11).Value = ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = ws.Cells(q, 9).Value
            ElseIf ws.Cells(q, 11).Value = ws.Cells(3, 17).Value Then
                ws.Cells(3, 16).Value = ws.Cells(q, 9).Value
            ElseIf ws.Cells(q, 12).Value = ws.Cells(4, 17).Value Then
                ws.Cells(4, 16).Value = ws.Cells(q, 9).Value
            End If
        Next q
Next ws
End Sub
