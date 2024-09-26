Attribute VB_Name = "Module1"
Sub LoopThroughStocks()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If i = 2 Then
            openingPrice = ws.Cells(i, 3).Value
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i - 1, 1).Value
            closingPrice = ws.Cells(i - 1, 6).Value
            quarterlyChange = closingPrice - openingPrice
            If openingPrice <> 0 And closingPrice <> 0 Then
                percentChange = (quarterlyChange / openingPrice) * 100
            Else
                percentChange = 0
            End If
            totalVolume = totalVolume + ws.Cells(i - 1, 7).Value
            ws.Cells(summaryTableRow, 9).Value = ticker
            ws.Cells(summaryTableRow, 10).Value = quarterlyChange
            ws.Cells(summaryTableRow, 11).Value = percentChange
            ws.Cells(summaryTableRow, 12).Value = totalVolume
            summaryTableRow = summaryTableRow + 1
            openingPrice = ws.Cells(i, 3).Value
            totalVolume = 0
        Else
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        End If
    Next i
End Sub

