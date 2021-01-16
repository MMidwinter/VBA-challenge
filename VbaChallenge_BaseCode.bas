Attribute VB_Name = "Module1"
Sub VBAChallenge()

    'Establish Dimensions and Variables
    Dim Ticker As String
    Dim Year As String
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim StockVolume As Double
    Dim StockYoYChange As Double
    Dim StockPercentChange As Double
    
    'Set initial values of Stock Data
    StockOpen = 0
    StockClose = 0
    StockVolume = 0
    StockYoYChange = 0
    StockPercentChange = 0
    
    'Worksheet For Loop
    For Each ws In Worksheets
    
        'Establish ForLoop Dimensions & Variables
        Dim LastRow As Double
        Dim LastColumn As Double
        Dim WorksheetName As String
        Dim TickerRow As Double
        
        WorksheetName = ws.Name
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = LastColumn = ws.Cells(Rows.Count, "A").End(xlUp).Row
        TickerRow = 2
        
        'Add Headers
            ws.Range("I1") = "Ticker"
            ws.Range("j1") = "Yearly Change"
            ws.Range("k1") = "Percent Changed"
            ws.Range("l1") = "Total Stock Volume"
        
        'For loop for Ticker
            For i = 2 To LastRow
            
                Ticker = ws.Cells(i, 1).Value
                
                'Debug.Print (Ticker)
                
                'Print Ticker Value when it Changes
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Debug.Print (Ticker)
                    ws.Range("I" & TickerRow).Value = Ticker
                    
                    'Add Current Row to Stock Volume
                    StockVolume = StockVolume + ws.Cells(i, 7).Value
                    
                    'Print StockVolume
                    ws.Range("L" & TickerRow).Value = StockVolume
                    'Debug.Print (StockVolume)
                
                    'Record Stock Close
                    StockClose = ws.Cells(i, 6).Value
                    'Debug.Print (StockClose)
                    
                    'Calculate Stock Yearly Change
                    StockYoYChange = StockClose - StockOpen
                    'Debug.Print (StockYoYChange)
                    
                    'Print Stock YoY Change
                    ws.Range("J" & TickerRow).Value = StockYoYChange
                    
                    'Formatting Stock YoY Change Color
                    
                    If ws.Range("J" & TickerRow).Value >= 0 Then
                    
                        ws.Range("J" & TickerRow).Interior.ColorIndex = 4
                        
                    Else
                    
                        ws.Range("J" & TickerRow).Interior.ColorIndex = 3
                    
                    End If
                    
                        If StockYoYChange <> 0 Then
                        
                            'Calculate Percent Changed
                            StockPercentChange = StockYoYChange / StockOpen
                        
                            'Print Percent Changed
                            ws.Range("k" & TickerRow).Value = StockPercentChange
                            
                        Else
                        
                           'Calculate Percent Changed
                            StockPercentChange = 0
                        
                            'Print Percent Changed
                            ws.Range("k" & TickerRow).Value = StockPercentChange
                            
                        End If
                        
                    'Format Stock YoY Change as %
                    ws.Range("k" & TickerRow).Value = Format(ws.Range("k" & TickerRow).Value, "0.00%")
                    
                    'Reset StockVolume
                    StockVolume = 0
                    
                    'Reset StockOpen
                    'StockOpen = 0
                    
                    'Reset Stock YoY Change
                    'StockYoYChange = 0
                    
                    'Reset Percent Changed
                    'StockPercentChange = 0
                    
                    'Increase the Ticker Row
                    TickerRow = TickerRow + 1
                
                    
                
                'If the Ticker doesnt change and remains the same
                Else
                
                    'Add to Stock Open
                    StockVolume = StockVolume + ws.Cells(i, 7).Value
                    
                    'Record Stock Open
                    
                        If ws.Cells(i, 3).Value > 0 Then
                
                            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                            StockOpen = ws.Cells(i, 3).Value
                            'Debug.Print (StockOpen)
                            
                            End If
                        
                        End If
                    
                   
                End If
                
            Next i
    
    'Close Worksheet For Loop
    Next ws

End Sub



