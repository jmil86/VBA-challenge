# VBA-challenge

Descriptions of each category notated

'Set Variables
Sub StockData()
    Dim lastRow As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim i As Long
    Dim ws As Worksheet
    Dim startRow As Long
    Dim maxIncrease As Double
    Dim maxIncreaseTicker As String
    Dim minIncrease As Double
    Dim minIncreaseTicker As String
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    Dim lastRowSummary As Long
    
    

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Find the last row with data
        outputRow = 2 ' Start output at row 2
        totalVolume = 0
        startRow = 2
        

        ' Headers for the output columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        

        ' Loop through all rows of stock data
        For i = 2 To lastRow
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(startRow, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                

                ' Calculate quarterly change
                quarterlyChange = closingPrice - openingPrice

                ' Avoid division by zero
                If openingPrice <> 0 Then
                    percentChange = (quarterlyChange / openingPrice)
                Else
                    percentChange = 0
                End If

                ' Write results to output columns
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 10).NumberFormat = "0.00"
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalVolume
                
                Select Case quarterlyChange
                    Case Is > 0
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                    Case Else
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 0
                End Select
                
                
                ' Move to the next output row
                outputRow = outputRow + 1
                totalVolume = 0
                startRow = i + 1
                
            End If
        
        Next i
        
'initialize variables
maxIncrease = 0
minIncrease = 0
maxVolume = 0

lastRowSummary = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'loop through the summary and find the min and max
    For x = 2 To lastRowSummary
    
    
        'look for the greatest increase percentage
        If ws.Cells(x, 11).Value > maxIncrease Then
            maxIncrease = ws.Cells(x, 11).Value
            maxIncreaseTicker = ws.Cells(x, 9).Value
        End If
        
        'look for the greatest decrease percentage
        If ws.Cells(x, 11).Value < minIncrease Then
            minIncrease = ws.Cells(x, 11).Value
            minIncreaseTicker = ws.Cells(x, 9).Value
        End If
    
        'look for the max total volume
        If ws.Cells(x, 12).Value > maxVolume Then
            maxVolume = ws.Cells(x, 12).Value
            maxVolumeTicker = ws.Cells(x, 9).Value
            
        End If
        
    

        'set the ticker and values
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(3, 16).Value = minIncreaseTicker
        ws.Cells(3, 17).Value = minIncrease
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
        
        Next x
         'set the ticker and values
        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxIncrease
        ws.Cells(3, 16).Value = minIncreaseTicker
        ws.Cells(3, 17).Value = minIncrease
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
        
    Next ws

End Sub
