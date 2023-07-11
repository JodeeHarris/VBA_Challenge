Attribute VB_Name = "Module1"
Sub Analytics()

    Dim i As Long
    Dim n As Double
    Dim ticker As Long
    Dim SumOpen As Double
    Dim SumClose As Double
    Dim OpenMinusClose As Double
    Dim PercentSign As String
    Dim StockTot As LongLong
    Dim SummaryTable As Long
   
    For Each ws In Worksheets
    ws.Activate
    Dim GreatestPercent As Double
    Dim GreatestTicker As String
    Dim LowestPercent As Double
    Dim LowestTicker As String
    Dim StockTicker As String
    Dim GreatestStock As LongLong
    
    GreatestPercent = 0
    
    SummaryTable = 2
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    LastPercent = Cells(Rows.Count, 11).End(xlUp).Row
    LastValue = Cells(Rows.Count, 12).End(xlUp).Row
    StockTot = 0
        
                
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
    
                
        Cells(2, 15) = "Greateast % Increase"
        Cells(2 + 1, 15) = "Greatest % Decrease"
        Cells(2 + 2, 15) = "Greatest Total Volume"
        
    
        For i = 2 To LastRow
            
            StockTot = StockTot + Cells(i, 7).Value
                
                If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                SumClose = Cells(i, 6).Value
                OpenMinusClose = SumClose - SumOpen
        'May have to switch Sumopen and Sumclose values
        
                Cells(SummaryTable, 9).Value = Cells(i, 1).Value
                Cells(SummaryTable, 10).Value = OpenMinusClose
                If OpenMinusClose < 0 Then
                
                    Cells(SummaryTable, 10).Interior.Color = RGB(255, 0, 0)
                
                ElseIf Cells(SummaryTable, 10).Value > 0 Then
                
                    Cells(SummaryTable, 10).Interior.Color = RGB(0, 255, 0)
                    
                End If
                
                
                PercentChange = (OpenMinusClose / SumOpen) * 100
                
                Cells(SummaryTable, 11).Value = PercentChange & "%"
                
                Cells(SummaryTable, 12).Value = StockTot
                
                
                If GreatestPercent < PercentChange Then
                    
                    GreatestPercent = PercentChange
                
                    GreatestTicker = Cells(i, 1).Value
                
                End If
                
                If LowestPercent > PercentChange Then
                    
                    LowestPercent = PercentChange
                
                    LowestTicker = Cells(i, 1).Value
                
                End If
                
                If GreatestStock < StockTot Then
                    
                    GreatestStock = StockTot
                
                    StockTicker = Cells(i, 1).Value
                
                End If
                
                SummaryTable = SummaryTable + 1
                
                StockTot = 0
                
                
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
                SumOpen = Cells(i, 3).Value
                
            End If
    
        Next i
    
    Cells(2, 16).Value = GreatestTicker
    Cells(2, 17).Value = GreatestPercent
    Cells(3, 16).Value = LowestTicker
    Cells(3, 17).Value = LowestPercent
    Cells(4, 16).Value = StockTicker
    Cells(4, 17).Value = GreatestStock
    
    ws.UsedRange.EntireColumn.AutoFit
    Next ws

End Sub


