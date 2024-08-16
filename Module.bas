Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Declare variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim lastRow As Long
    Dim summaryRow As Integer
    
    ' Variables to track greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Loop through each worksheet
    For Each ws In Worksheets
        
        ' Reset variables
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Add column titles for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Find the last row with data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize the first open price
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through all rows of data
        For i = 2 To lastRow
            
            ' Check if the ticker changes (end of quarter for that stock)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Set the close price for the last day of the quarter
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate quarterly change
                quarterlyChange = closePrice - openPrice
                
                ' Calculate percentage change
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice) * 100
                Else
                    percentageChange = 0
                End If
                
                ' Calculate total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output the data to the summary table
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = quarterlyChange
                ws.Cells(summaryRow, 11).Value = Format(percentageChange, "0.00") & "%"
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Apply conditional formatting for positive/negative changes
                If quarterlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = vbRed
                End If
                
                If percentageChange >= 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(summaryRow, 11).Interior.Color = vbRed
                End If
                
                ' Track the greatest percentage increase
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    greatestIncreaseTicker = ticker
                End If
                
                ' Track the greatest percentage decrease
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    greatestDecreaseTicker = ticker
                End If
                
                ' Track the greatest total volume
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                ' Move to the next summary row
                summaryRow = summaryRow + 1
                
                ' Reset for the next ticker
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Accumulate volume if still within the same ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        ' Output the greatest values to the worksheet
        ws.Cells(1, 14).Value = "Greatest % Increase"
        ws.Cells(2, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 14).Value = "Greatest Total Volume"
        
        ws.Cells(1, 15).Value = greatestIncreaseTicker
        ws.Cells(2, 15).Value = greatestDecreaseTicker
        ws.Cells(3, 15).Value = greatestVolumeTicker
        
        ws.Cells(1, 16).Value = Format(greatestIncrease, "0.00") & "%"
        ws.Cells(2, 16).Value = Format(greatestDecrease, "0.00") & "%"
        ws.Cells(3, 16).Value = greatestVolume
        
    Next ws
    
    ' Format the column headers
    For Each ws In Worksheets
        ws.Range("I1:L1").Font.Bold = True
        ws.Range("I1:L1").Interior.Color = RGB(200, 200, 200) ' Light gray background
        ws.Range("N1:O1").Font.Bold = True
        ws.Range("N1:O1").Interior.Color = RGB(200, 200, 200) ' Light gray background
    Next ws
    
End Sub
