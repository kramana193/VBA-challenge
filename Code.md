# VBA-challenge
Sub Multiple_Year_Stock_data_Ramana()

    For Each ws In Worksheets
    
        Dim j As Integer
        j = 2
        
        Dim YearlyOpening As Double
        Dim YC As Double
        Dim Total As Double
        Total = 0
        
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
        
        Dim GITicker As String
        Dim GDTicker As String
        Dim GTVTicker As String
        
    
        ws.Range("M1").Value = "Unique Ticker"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("O1").Value = "Yearly Percentage Change"
        ws.Range("P1").Value = "Total Stock Volume"
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastRowNew = ws.Cells(Rows.Count, 14).End(xlUp).Row
        
        For i = 2 To LastRow
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                YC = ws.Cells(i, 3).Value
                YearlyOpening = ws.Cells(i, 3).Value
                
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                YC = ws.Cells(i, 6) - YC
                ws.Cells(j, 13).Value = ws.Cells(i, 1).Value
                ws.Cells(j, 14).Value = YC
                ws.Cells(j, 15).Value = (YC / YearlyOpening)
                ws.Cells(j, 16).Value = Total + ws.Cells(i, 7)
                YC = 0
                YearlyOpening = 0
                Total = 0
                j = j + 1
            Else: Total = Total + ws.Cells(i, 7).Value
                
            End If
            
        Next i
            
        For k = 2 To LastRowNew
                ws.Cells(k, 15).NumberFormat = "0.00%"
                
                If ws.Cells(k, 14).Value > 0 Then
                    ws.Cells(k, 14).Interior.ColorIndex = 4
                End If
                
                If ws.Cells(k, 14).Value < 0 Then
                    ws.Cells(k, 14).Interior.ColorIndex = 3
                End If
                
                If ws.Cells(k, 14).Value = 0 Then
                    ws.Cells(k, 14).Interior.ColorIndex = 0
                End If
                
                If ws.Cells(k, 15) > GreatestIncrease Then
                    GreatestIncrease = ws.Cells(k, 15).Value
                    GITicker = ws.Cells(k, 13).Value
                End If
                
                If ws.Cells(k, 15) < GreatestDecrease Then
                    GreatestDecrease = ws.Cells(k, 15).Value
                    GDTicker = ws.Cells(k, 13).Value
                End If
                
                If ws.Cells(k, 16) > GreatestTotalVolume Then
                    GreatestTotalVolume = ws.Cells(k, 16).Value
                    GTVTicker = ws.Cells(k, 13).Value
                End If
                          
        Next k
        
        ws.Range("T2:T3").NumberFormat = "0.00%"
        ws.Range("S2").Value = GITicker
        ws.Range("S3").Value = GDTicker
        ws.Range("S4").Value = GTVTicker
        ws.Range("T2").Value = GreatestIncrease
        ws.Range("T3").Value = GreatestDecrease
        ws.Range("T4").Value = GreatestTotalVolume
        
    Next ws

End Sub


