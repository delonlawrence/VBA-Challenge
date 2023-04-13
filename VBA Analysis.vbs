Attribute VB_Name = "Module1"
Sub StockMarketDataAnalysis()
    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim ticker As String
    Dim summaryTableRowIndex As Long
    Dim greatestPercentIncrease As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    Dim tickerGreatestPercentIncrease As String
    Dim tickerGreatestPercentDecrease As String
    Dim tickerGreatestTotalVolume As String
    
   ' Loop for applying to all worksheets
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row of data in column A
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' positions of headers for summary tables
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Row where summary table data will begin
    summaryTableRowIndex = 2
    
    ' starting point for tracking greatest percentage increase, decrease, and total volume
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    tickerGreatestPercentIncrease = ""
    tickerGreatestPercentDecrease = ""
    tickerGreatestTotalVolume = ""
    
    ' Loop through each row of data
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            
            openingPrice = ws.Cells(i, 3).Value
            
            ' For Resetting the total volume
            totalVolume = 0
        End If
        
        ' stock volume + total volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        ' Check if the ticker symbol changes or if it's the last row of data
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
    
            closingPrice = ws.Cells(i, 6).Value
            
            ' yearly change calculation
            yearlyChange = closingPrice - openingPrice
            
            ' percent change calculation
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If
            
            ' Output the results to the summary table
            ws.Cells(summaryTableRowIndex, 9).Value = ticker
            ws.Cells(summaryTableRowIndex, 10).Value = yearlyChange
            ws.Cells(summaryTableRowIndex, 11).Value = percentChange
            ws.Cells(summaryTableRowIndex, 12).Value = totalVolume
            
            ' Format the percent change as percentage
            ws.Cells(summaryTableRowIndex, 11).NumberFormat = "0.00%"
            
            ' Apply conditional formatting to highlight positive and negative yearly change
                If yearlyChange > 0 Then
                    ws.Cells(summaryTableRowIndex, 10).Interior.Color = RGB(198, 239, 206) ' Light Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryTableRowIndex, 10).Interior.Color = RGB(255, 199, 206) ' Light Red
                    
            End If
            
            ' Update the stock with greatest percent increase
            If percentChange > greatestPercentIncrease Then
                greatestPercentIncrease = percentChange

                tickerGreatestPercentIncrease = ticker
                
                
            End If
            
            ' Update the stock with greatest percent decrease
            If percentChange < greatestPercentDecrease Then
                greatestPercentDecrease = percentChange
                tickerGreatestPercentDecrease = ticker
            End If
            
            ' Update the stock with greatest total volume
            If totalVolume > greatestTotalVolume Then
                greatestTotalVolume = totalVolume
                tickerGreatestTotalVolume = ticker
            End If
            
            ' Move to the next row in the summary table
            summaryTableRowIndex = summaryTableRowIndex + 1
        End If
    Next i
    
    ' Output the stocks with greatest percent increase, decrease, and total volume
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    ws.Cells(2, 17).Value = tickerGreatestPercentIncrease
    ws.Cells(3, 17).Value = tickerGreatestPercentDecrease
    ws.Cells(4, 17).Value = tickerGreatestTotalVolume
    
    ws.Cells(2, 18).Value = greatestPercentIncrease
    ws.Cells(3, 18).Value = greatestPercentDecrease
    ws.Cells(4, 18).Value = greatestTotalVolume
    
    ' Format the greatest percent increase and decrease as percentage
    ws.Cells(2, 18).NumberFormat = "0.00%"
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    Next ws

End Sub

