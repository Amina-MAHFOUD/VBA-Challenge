Attribute VB_Name = "Module3"
Sub AnalyzeStocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Initialize variables
    greatestIncrease = -999999
    greatestDecrease = 999999
    greatestVolume = 0
    
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        Dim percentChange As Double
        Dim totalVolume As Double
        
        ' Calculate percent change
        percentChange = ws.Cells(i, 11).Value
        
        ' Get total volume
        totalVolume = ws.Cells(i, 12).Value
        
        ' Check for greatest percent increase
        If percentChange > greatestIncrease Then
            greatestIncrease = percentChange
            greatestIncreaseTicker = ticker
        End If
        
        ' Check for greatest percent decrease
        If percentChange < greatestDecrease Then
            greatestDecrease = percentChange
            greatestDecreaseTicker = ticker
        End If
        
        ' Check for greatest total volume
        If totalVolume > greatestVolume Then
            greatestVolume = totalVolume
            greatestVolumeTicker = ticker
        End If
    Next i
    
    ' Output the results
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(2, 17).Value = greatestIncrease
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = greatestDecreaseTicker
    ws.Cells(3, 17).Value = greatestDecrease
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = greatestVolumeTicker
    ws.Cells(4, 17).Value = greatestVolume
End Sub

