Attribute VB_Name = "Module1"
Sub WallStreet_StockMarket()

For Each ws In Worksheets

    ' Declare variables
    
    Dim i As Long
    Dim TotalStockVolume As LongLong
    Dim TickerCount As Long
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim greatestPerInc As Double
    Dim greatestPerDec As Double
    Dim GreatestTotalVolume As LongLong
    Dim PerIncTicker As String
    Dim PerDecTicker As String
    Dim StockVolumeTicker As String
    
    
    
    'Get the WorksheetName
    
    WorksheetName = ws.Name
    
    
    ' Create summary table
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Create bonus table
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ' Beginning values
    
    TickerCount = 2
    greatestPerInc = 0
    greatestPerDec = 0
    GreatestTotalVolume = 0
    OpenPrice = ws.Range("C2").Value
    
    ' Determine the Last Row
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To lastrow
    
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Fill in summary table
            
            ' Get closing price
            
            ClosingPrice = ws.Cells(i, 6).Value
            
            ' Ticker
            
            ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
            
            ' Yearly change from opening price to closing price
            
            YearlyChange = ClosingPrice - OpenPrice
            
            ' Formatting for yearly change
            
            If YearlyChange < 0 Then
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 3

            Else
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 4

            End If
            
            ws.Cells(TickerCount, 10).Value = YearlyChange
            
            ' Percent change from opening price to closing price
            
            PercentChange = (ClosingPrice - OpenPrice) / OpenPrice
            ws.Cells(TickerCount, 11).Value = PercentChange
            
            ' Total stock volume
            
            ws.Cells(TickerCount, 12).Value = TotalStockVolume
            
            ' Bonus section
            
            If PercentChange > greatestPerInc Then
                greatestPerInc = PercentChange
                PerIncString = ws.Cells(i, 1).Value
                
            ElseIf PercentChange < greatestPerDec Then
                greatestPerDec = PercentChange
                PerDecString = ws.Cells(i, 1).Value
            End If
            
            If TotalStockVolume > GreatestTotalVolume Then
                GreatestTotalVolume = TotalStockVolume
                StockVolumeString = ws.Cells(i, 1).Value
            End If
            
            ' Reset for next round
            
            TotalStockVolume = 0
            OpenPrice = ws.Cells(i + 1, 3).Value
            TickerCount = TickerCount + 1
            
        End If
    
    Next i
    
    ' Fill out bonus section
    
    ws.Range("P2").Value = PerIncString
    ws.Range("Q2").Value = greatestPerInc
    
    ws.Range("P3").Value = PerDecString
    ws.Range("Q3").Value = greatestPerDec
    
    ws.Range("P4").Value = StockVolumeString
    ws.Range("Q4").Value = GreatestTotalVolume
    
     ' Formatting
     
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Columns("A:Q").AutoFit

Next ws

End Sub
