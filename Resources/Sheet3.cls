VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockTracker()

    ' total up the volume, then compare to next stock
    ' write to column I and L
    ' reset counter
    
    Dim stockOne As String
    Dim totalStockVol As Double
    Dim stockTwo As String
    Dim row As Long
    Dim summaryRow As Integer
    Dim lastRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    
        'set my defaults
        totalStockVol = 0
        summaryRow = 2
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        
        'this is the opening price of the first stock
        openPrice = ws.Cells(2, 3).Value
        
        'this loops through the worksheets
        For row = 2 To lastRow
            stockOne = ws.Cells(row, 1).Value
            stockTwo = ws.Cells(row + 1, 1).Value
            
            totalStockVol = totalStockVol + ws.Cells(row, 7).Value
        
            If stockTwo <> stockOne Then
                closePrice = ws.Cells(row, 6).Value
                yearlyChange = closePrice - openPrice
                
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If
            
                'this writes to summary table
                ws.Cells(summaryRow, 9).Value = stockOne
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalStockVol
                
                'this section handles the conditional formatting
                If yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 4
                End If
                
                'resets the volume and moves summaryRow
                totalStockVol = 0
                summaryRow = summaryRow + 1
                
                'this checks for the price of next stock
                openPrice = ws.Cells(row + 1, 3).Value
            End If
        
        Next row
        
        'this second loop works to get the summaries
        Dim lastSummaryRow As Long
        Dim greatestVolume As Double
        Dim greatestVolumeStock As String
        Dim greatestIncrease As Double
        Dim greatestIncreaseStock As String
        Dim greatestDecrease As Double
        Dim greatestDecreaseStock As String
        
        'this sets the defaults
        lastSummaryRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        greatestVolume = ws.Cells(2, 12).Value
        greatestVolumeStock = ws.Cells(2, 9).Value
        
        greatestIncrease = ws.Cells(2, 11).Value
        greatestIncreaseStock = ws.Cells(2, 9).Value
        
        greatestDecrease = ws.Cells(2, 11).Value
        greatestDecreaseStock = ws.Cells(2, 9).Value
        
        For row = 2 To lastSummaryRow
        
            'this finds the greatest volume
            If ws.Cells(row, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(row, 12).Value
                greatestVolumeStock = ws.Cells(row, 9).Value
            End If
            
            'this looks for the greatest change
            If ws.Cells(row, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(row, 11).Value
                greatestIncreaseStock = ws.Cells(row, 9).Value
            End If
            
            'this looks for the worst decrease
            If ws.Cells(row, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(row, 11).Value
                greatestDecreaseStock = ws.Cells(row, 9).Value
            End If
        
        Next row
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseStock
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseStock
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeStock
        ws.Cells(4, 17).Value = greatestVolume
        
        ws.Range("J2:J" & summaryRow).NumberFormat = "0.00"
        ws.Range("K2:K" & summaryRow).NumberFormat = "0.00%"
    Next
End Sub
