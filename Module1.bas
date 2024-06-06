Attribute VB_Name = "Module1"
Sub CalculateQuarterlyChangeAndFormatCells()

    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim percentChange As Double
    Dim quarterlyChange As Double
    Dim lastRow As Long
    Dim outputRow As Integer
    Dim quarter As Integer
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        outputRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        For i = 2 To lastRow
            If i = 2 Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                If i > 2 Then
                    ws.Cells(outputRow, 10).Value = ticker
                    ws.Cells(outputRow, 11).Value = quarterlyChange
                    ws.Cells(outputRow, 12).Value = percentChange
                    ws.Cells(outputRow, 13).Value = totalVolume
                    outputRow = outputRow + 1
                    
                    
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        increaseTicker = ticker
                    ElseIf percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        decreaseTicker = ticker
                    End If
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        volumeTicker = ticker
                    End If
                End If
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            closePrice = ws.Cells(i, 5).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            quarterlyChange = closePrice - openPrice
            percentChange = (quarterlyChange / openPrice) * 100
        Next i
        
        ws.Cells(outputRow, 10).Value = ticker
        ws.Cells(outputRow, 11).Value = quarterlyChange
        ws.Cells(outputRow, 12).Value = percentChange
        ws.Cells(outputRow, 13).Value = totalVolume
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Volume"
        ws.Range("J1:M1").Font.Bold = True
        
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(2, 11), ws.Cells(outputRow, 11))
        
        
        rng.FormatConditions.Delete
        
        
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
            .Interior.Color = RGB(0, 255, 0)
        End With
        
        
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
            .Interior.Color = RGB(255, 0, 0)
        End With

        
        ws.Cells(2, 18).Value = "Greatest % Increase"
        ws.Cells(2, 19).Value = increaseTicker
        ws.Cells(2, 20).Value = greatestIncrease
        ws.Cells(3, 18).Value = "Greatest % Decrease"
        ws.Cells(3, 19).Value = decreaseTicker
        ws.Cells(3, 20).Value = greatestDecrease
        ws.Cells(4, 18).Value = "Greatest Total Volume"
        ws.Cells(4, 19).Value = volumeTicker
        ws.Cells(4, 20).Value = greatestVolume

        
        With ws.Range("L2:L" & ws.Cells(ws.Rows.Count, "L").End(xlUp).Row)
            .NumberFormat = "0%"
        End With
    Next ws

End Sub


