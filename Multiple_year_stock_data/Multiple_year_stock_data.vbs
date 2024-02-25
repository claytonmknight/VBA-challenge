Sub CalculateStockMetrics()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim wsIndex As Integer
    
    Dim maxPercentIncrease As Double
    Dim minPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim minPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row of data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 10).Interior.Color = RGB(255, 255, 255) ' White background
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize variables
        outputRow = 2
        totalVolume = 0 ' Initialize total volume
        maxPercentIncrease = 0
        minPercentDecrease = 0
        maxTotalVolume = 0
        maxPercentIncreaseTicker = ""
        minPercentDecreaseTicker = ""
        maxTotalVolumeTicker = ""
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Check if the ticker symbol has changed
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
                ' If it's not the first row, calculate and output previous ticker's metrics
                If i > 2 Then
                    ' Calculate yearly and percentage change
                    yearlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = (yearlyChange / openPrice) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Output the previous ticker's metrics to the next row
                    ws.Cells(outputRow, 9).Value = ticker
                    ws.Cells(outputRow, 10).Value = yearlyChange
                    
                    ' Format yearly change based on positive or negative value
                    If yearlyChange > 0 Then
                        ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Bright green
                    ElseIf yearlyChange < 0 Then
                        ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Bright red
                    End If
                    
                    ws.Cells(outputRow, 11).Value = percentChange
                    ws.Cells(outputRow, 12).Value = totalVolume  ' Make sure total volume is correctly assigned here
                    
                    ' Check for max percent increase, min percent decrease, and max total volume
                    If percentChange > maxPercentIncrease Then
                        maxPercentIncrease = percentChange
                        maxPercentIncreaseTicker = ticker
                    End If
                    
                    If percentChange < minPercentDecrease Then
                        minPercentDecrease = percentChange
                        minPercentDecreaseTicker = ticker
                    End If
                    
                    If totalVolume > maxTotalVolume Then
                        maxTotalVolume = totalVolume
                        maxTotalVolumeTicker = ticker
                    End If
                    
                    ' Move to the next output row
                    outputRow = outputRow + 1
                End If
                
                ' Set the new ticker symbol and open price
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                
                ' Reset total volume for the new ticker
                totalVolume = 0
            End If
            
            ' Get the close price and update total volume
            closePrice = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 4).Value  ' Update to column D
        Next i
        
        ' Output the last ticker's metrics
        yearlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentChange = (yearlyChange / openPrice) * 100
        Else
            percentChange = 0
        End If
        
        ws.Cells(outputRow, 9).Value = ticker
        ws.Cells(outputRow, 10).Value = yearlyChange
        
        ' Format yearly change based on positive or negative value using conditional formatting
        Dim conditionRange As Range
        Set conditionRange = ws.Cells(outputRow, 10)

        With conditionRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0) ' Bright green
        End With

        With conditionRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0) ' Bright red
        End With
        
        ws.Cells(outputRow, 11).Value = percentChange
        ws.Cells(outputRow, 12).Value = totalVolume
        
        ' Output greatest % increase, greatest % decrease, and greatest total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(2, 16).Value = maxPercentIncreaseTicker
        ws.Cells(3, 16).Value = minPercentDecreaseTicker
        ws.Cells(4, 16).Value = maxTotalVolumeTicker

        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 17).Value = minPercentDecrease
        ws.Cells(4, 17).Value = maxTotalVolume

        ' Resize all columns to fit all text
        ws.Columns("A:Q").AutoFit
    Next ws
End Sub