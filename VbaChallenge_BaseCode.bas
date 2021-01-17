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
            For I = 2 To LastRow
            
                Ticker = ws.Cells(I, 1).Value
                
                'Debug.Print (Ticker)
                
                'Print Ticker Value when it Changes
                If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                    
                    'Debug.Print (Ticker)
                    ws.Range("I" & TickerRow).Value = Ticker
                    
                    'Add Current Row to Stock Volume
                    StockVolume = StockVolume + ws.Cells(I, 7).Value
                    
                    'Print StockVolume
                    ws.Range("L" & TickerRow).Value = StockVolume
                    'Debug.Print (StockVolume)
                
                    'Record Stock Close
                    StockClose = ws.Cells(I, 6).Value
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
                    StockVolume = StockVolume + ws.Cells(I, 7).Value
                    
                    'Record Stock Open
                    
                        If ws.Cells(I, 3).Value > 0 Then
                
                            If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value Then
                
                            StockOpen = ws.Cells(I, 3).Value
                            'Debug.Print (StockOpen)
                            
                            End If
                        
                        End If
                    
                   
                End If
                
            Next I
            
            LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            'Set Variable Values
            Dim GreatestPercentIncrease As Double
            Dim GreatestPercentIncreaseTicker As String
            Dim GreatestPercentDecrease As Double
            Dim GreatestPercentDecreaseTicker As String
            Dim GreatestTotalVolume As Double
            Dim GreatestTotalVolumeTicker As String

            
            GreatestPercentIncrease = 0
            GreatestPercentIncreaseTicker = ""
            GreatestPercentDecrease = 0
            GreatestPercentDecreaseTicker = ""
            GreatestTotalVolume = 0
            GreatestTotalVolumeTicker = ""
            
            For B = 2 To LastRow
                
                'Bonus Work Assignment
                ws.Range("o2") = "Greatest % Increase"

                'If Statement to find greatest increase
                If ws.Cells(B, 11) > GreatestPercentIncrease Then
                
                    GreatestPercentIncreaseTicker = ws.Cells(B, 9)
                    GreatestPercentIncrease = ws.Cells(B, 11)
                    
                End If
                
                If ws.Cells(B, 11) < GreatestPercentDecrease Then
                
                    GreatestPercentDecreaseTicker = ws.Cells(B, 9)
                    GreatestPercentDecrease = ws.Cells(B, 11)
                    
                End If
                
                If ws.Cells(B, 12) > GreatestTotalVolume Then
                
                    GreatestTotalVolumeTicker = ws.Cells(B, 9)
                    GreatestTotalVolume = ws.Cells(B, 12)
                    
                End If
                
                ws.Range("o3") = "Greatest % Decrease"
                ws.Range("o4") = "Greatest Total Volume"
                ws.Range("p1") = "Ticker"
                ws.Range("q1") = "Value"
                ws.Range("P2") = GreatestPercentIncreaseTicker
                ws.Range("Q2") = GreatestPercentIncrease
                ws.Range("P3") = GreatestPercentDecreaseTicker
                ws.Range("Q3") = GreatestPercentDecrease
                ws.Range("P4") = GreatestTotalVolumeTicker
                ws.Range("Q4") = GreatestTotalVolume
                ws.Range("Q2").Value = Format(ws.Range("q2").Value, "0.00%")
                ws.Range("Q3").Value = Format(ws.Range("q3").Value, "0.00%")
                
            Next B
            
            ws.Columns("A:Q").EntireColumn.AutoFit
            
    'Close Worksheet For Loop
    Next ws

End Sub



