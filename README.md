Got help with these parts of my code from Chat GPT
' Apply conditional formatting to highlight positive and negative yearly change
                If yearlyChange > 0 Then
                    ws.Cells(summaryTableRowIndex, 10).Interior.Color = RGB(198, 239, 206) ' Light Green
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryTableRowIndex, 10).Interior.Color = RGB(255, 199, 206) ' Light Red
  
  I knew how to do this but saw it last minute asked chat for help
  
  
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
     
    I really struggled with this part so I went to chat
    
    The way I use chat is I ask it for what I want with an explanation. Then I try to duplicate it based on what I understand. My understanding of this is better but not perfect. This will not be a frequent practice.
    
