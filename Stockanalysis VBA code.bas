Attribute VB_Name = "Module1"
Sub ticker()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputCell As Range
    Dim i As Long
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Set outputCell = ws.Cells(2, 9)
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                outputCell.Value = ws.Cells(i, 1).Value
                Set outputCell = outputCell.Offset(1, 0)
            End If
        Next i
    Next ws
End Sub

Sub Pricechanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlychange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim ticker As String
    Dim firstRow As Long
    Dim currentrow As Long
    Dim ticketcount As Long
    
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        firstRow = 2
        currentrow = 2
        ticketcount = 2
        Do While currentrow <= lastRow
            ticker = ws.Cells(currentrow, 1).Value
            'Conditional formatting
            If ws.Cells(ticketcount, 10).Value > 0 Then
            ws.Cells(ticketcount, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(ticketcount, 10).Interior.ColorIndex = 3
            End If
    
            'Calculating the opening price
            openingPrice = Val(ws.Cells(currentrow, 3).Value)
            totalVolume = 0
            Do While ws.Cells(currentrow, 1).Value = ticker
                totalVolume = totalVolume + Val(ws.Cells(currentrow, 7).Value)
                currentrow = currentrow + 1
                If currentrow > lastRow Then Exit Do
            Loop
            
        
          'Calculating the closing price and yearlychange
            
            closingPrice = Val(ws.Cells(currentrow - 1, 6).Value)
            yearlychange = closingPrice - openingPrice
            percentChange = (yearlychange / openingPrice) * 100
            ws.Cells(ticketcount, 10).Value = yearlychange
            ws.Cells(ticketcount, 11).Value = percentChange
            ws.Cells(ticketcount, 12).Value = totalVolume '
            ticketcount = ticketcount + 1
        Loop
    
    Next ws

End Sub

Sub columnheader()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Sheets
'Add column headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Next ws
End Sub

Sub Greaterpercentanaysis()
    Dim ws As Worksheet
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    ' Looping  through all worksheets
    
    For Each ws In ThisWorkbook.Sheets
        For currentrow = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            percentChange = ws.Cells(currentrow, 11).Value
            totalVolume = ws.Cells(currentrow, 12).Value
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                greatestIncreaseTicker = ws.Cells(currentrow, 9).Value
            End If
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                greatestDecreaseTicker = ws.Cells(currentrow, 9).Value
            End If
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                greatestVolumeTicker = ws.Cells(currentrow, 9).Value
            End If
        Next currentrow
        ' Produce the results on a specific row in each worksheet
        Dim resultrow As Long
    resultrow = 2
        ' Write the results to each worksheet
        ws.Cells(resultrow, 16).Value = greatestIncreaseTicker
        ws.Cells(resultrow, 17).Value = maxIncrease
        ws.Cells(resultrow, 16).Offset(0, 1).Value = maxIncrease
        ws.Cells(resultrow + 1, 16).Value = greatestDecreaseTicker
        ws.Cells(resultrow + 17).Value = maxDecrease
        ws.Cells(resultrow + 1, 16).Offset(0, 1).Value = maxDecrease
        ws.Cells(resultrow + 2, 16).Value = greatestVolumeTicker
        ws.Cells(resultrow + 17).Value = maxVolume
        ws.Cells(resultrow + 2, 16).Offset(0, 1).Value = maxVolume
    Next ws
End Sub
