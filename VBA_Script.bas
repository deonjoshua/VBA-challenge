Attribute VB_Name = "Module1"
Sub StockData():
    
    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim totalstock As Double
        totalstock = 0
        
        Dim rowcount As Integer
        rowcount = 2
        
        Dim openprice As Double
        openprice = ws.Range("C2").Value
        
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        
        For i = 2 To lastrow
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                totalstock = ws.Cells(i, 7).Value + totalstock
                
            Else
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                
                closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
                ws.Cells(rowcount, 10).Value = yearlychange
                
                If yearlychange > 0 Then
                    ws.Cells(rowcount, 10).Interior.Color = vbGreen
                    
                ElseIf yearlychange < 0 Then
                    ws.Cells(rowcount, 10).Interior.Color = vbRed
                End If
                
                percentchange = yearlychange / openprice
                ws.Cells(rowcount, 11).Value = FormatPercent(percentchange)
                
                openprice = ws.Cells(i + 1, 3).Value
                
                totalstock = ws.Cells(i, 7).Value + totalstock
                ws.Cells(rowcount, 12).Value = totalstock
                totalstock = 0
                
                rowcount = rowcount + 1
            End If
        Next i
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim greatinc As Double
        greatinc = 0
        
        Dim greatdec As Double
        greatdec = 0
        
        Dim greatvol As Double
        greatvol = 0
        
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value > greatinc Then
                greatinc = ws.Cells(i, 11).Value
                ws.Range("Q2").Value = FormatPercent(greatinc)
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < greatdec Then
                greatdec = ws.Cells(i, 11).Value
                ws.Range("Q3").Value = FormatPercent(greatdec)
                ws.Range("P3").Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > greatvol Then
                greatvol = ws.Cells(i, 12).Value
                ws.Range("Q4").Value = greatvol
                ws.Range("P4").Value = ws.Cells(i, 9).Value
            End If
        Next i
            
        ws.Columns("I:Q").AutoFit
    
    Next ws
        
End Sub

