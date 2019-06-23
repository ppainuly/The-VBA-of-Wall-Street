Attribute VB_Name = "Module1"
Sub summarize():

Dim TotalVol As Double
Dim ticker As String
Dim dateValue As String
Dim i As Long
Dim TickerRow As Integer
Dim add As Long
Dim x As Long
Dim lRow As Long
Dim change As Single

Dim openPrice As Double
Dim changePercent As Double
Dim changePercentStr As String

Dim GreatestVol As Double
Dim GreatestPercentage As Double
Dim LowestPercentage As Double
Dim ws As Worksheet

For Each ws In Worksheets
    openPrice = 0
    TickerRow = 2
    TotalVol = 0
    x = 0
    change = 0
    openPrice = ws.Cells(2, 3).Value
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    GreatestPercentage = 0
    
    'Set Headings
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Set Summary Headings
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

    For i = 2 To lRow
        add = ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                 TotalVol = TotalVol + ws.Cells(i, 7).Value
                 change = ws.Cells(i, 6).Value - openPrice

                     ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
                     ws.Cells(TickerRow, 12).Value = TotalVol
                     ws.Cells(TickerRow, 10).Value = change
                     If change >= 0 Then
                        ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
                    End If
                     
                     'If Open Price happens to be zero, the  change would be change * 100 Percent
                     If openPrice = 0 Then
                        changePercent = 0
                    Else
                        changePercent = (change / openPrice) * 100
                    End If
                     changePercentStr = Round(changePercent, 2) & "%"
                     ws.Cells(TickerRow, 21).Value = changePercent
                     ws.Cells(TickerRow, 11).Value = changePercentStr
                     TickerRow = TickerRow + 1
                 
                 TotalVol = 0
                 change = 0
                 openPrice = Cells(i + 1, 3).Value
            Else
                TotalVol = TotalVol + add
                    
            End If
            
        Next i
        
     'Assign Greatest Vol Increase to Summary
    GreatestVol = ws.Range("L:L").Find(WorksheetFunction.Max(ws.Range("L:L"))).Row
    ws.Cells(4, 16).Value = ws.Cells(GreatestVol, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(GreatestVol, 12).Value

    'Assign Greatest Percentage
    GreatestPercentage = ws.Range("U:U").Find(WorksheetFunction.Max(ws.Range("U:U"))).Row
    ws.Cells(2, 16).Value = ws.Cells(GreatestPercentage, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(GreatestPercentage, 21).Value & "%"

    'Assign Lowest Percentage
    LowestPercentage = ws.Range("U:U").Find(WorksheetFunction.Min(ws.Range("U:U"))).Row
    ws.Cells(3, 16).Value = ws.Cells(LowestPercentage, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(LowestPercentage, 21).Value & "%"

    'Delete temp column storing change percentages
      ws.Columns(21).EntireColumn.Delete

Next ws
End Sub

