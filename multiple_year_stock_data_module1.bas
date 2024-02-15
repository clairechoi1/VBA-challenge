Attribute VB_Name = "Module1"
Sub StockData()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change($)"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim TickerName As String
        Dim TotalStockVolume, i, lastrow As LongLong
        Dim SummaryTableRow As Integer
        Dim OpenValue, CloseValue, YearlyChange, PercentChange As Double
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        SummaryTableRow = 2
        TotalStockVolume = 0
        OpenValue = 0
        CloseValue = 0
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                ws.Range("I" & SummaryTableRow).Value = TickerName
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                TotalStockVolume = 0
                
                CloseValue = ws.Cells(i, 6).Value
                YearlyChange = CloseValue - OpenValue
                If OpenValue <> 0 Then
                    PercentChange = YearlyChange / OpenValue
                Else
                    PercentChange = 0
                End If
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                
                SummaryTableRow = SummaryTableRow + 1
                OpenValue = ws.Cells(i + 1, 3).Value
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                If OpenValue = 0 Then
                    OpenValue = ws.Cells(i, 3).Value
                End If
            End If
        Next i
        
        ws.Range("K:K").NumberFormat = "0.00%"
        
        For i = 2 To lastrow
            If ws.Cells(i, "J").Value < 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 3 ' Red
            ElseIf ws.Cells(i, "J").Value > 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 4 ' Green
            End If
        Next i
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim MaxPerIncr, MaxPerDecr, MaxTotalVol As Double
        Dim MaxPerIncr_ticker, MaxPerDecr_ticker, MaxTotalVol_ticker As String
        
        MaxPerIncr = 0
        MaxPerDecr = 0
        MaxTotalVol = 0
        
        For i = 2 To lastrow
            If ws.Cells(i, "K").Value > MaxPerIncr Then
                MaxPerIncr = ws.Cells(i, "K").Value
                MaxPerIncr_ticker = ws.Cells(i, "I").Value
            End If
            
            If ws.Cells(i, "K").Value < MaxPerDecr Then
                MaxPerDecr = ws.Cells(i, "K").Value
                MaxPerDecr_ticker = ws.Cells(i, "I").Value
            End If
            
            If ws.Cells(i, "L").Value > MaxTotalVol Then
                MaxTotalVol = ws.Cells(i, "L").Value
                MaxTotalVol_ticker = ws.Cells(i, "I").Value
    End If
Next i
   
ws.Cells(2, 17).Value = MaxPerIncr
ws.Cells(3, 17).Value = MaxPerDecr
ws.Cells(4, 17).Value = MaxTotalVol

ws.Cells(2, 16).Value = MaxPerIncr_ticker
ws.Cells(3, 16).Value = MaxPerDecr_ticker
ws.Cells(4, 16).Value = MaxTotalVol_ticker

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

Next ws
End Sub


