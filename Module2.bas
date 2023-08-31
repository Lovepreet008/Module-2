
Sub StockAnalyze()
Attribute StockAnalyze.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ws As Worksheet
    Dim TotalClose As Double
    Dim TotalOpen As Double
    Dim LastRow  As Long
    Dim i As Long
    Dim j As Integer
    Dim YearlyChange As Double
    Dim TotalStock As LongLong
    
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
    
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     j = 2
     k = 2
     s = 0
     t = 0
     u = 10000000
     TotalClose = 0
     TotalOpen = 0
     TotalStock = 0
        For i = 2 To LastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                '1. Ticker Symbol
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                
                '2.Yearly change
                
                openvalue = ws.Cells(k, 3).Value
                closevalue = ws.Cells(i, 6).Value
                YearlyChange = closevalue - openvalue
                ws.Cells(j, 10).Value = YearlyChange
                
                
                If YearlyChange < 0 Then
                    ws.Cells(j, 10).Interior.Color = vbRed
                Else:
                    ws.Cells(j, 10).Interior.Color = vbGreen
                End If

                '3. Percentage change
                
                          
                
                ws.Cells(j, 11).Value = ((closevalue - openvalue) / openvalue)
                k = i + 1
                '4. Total Stock Volume
                
                TotalStock = TotalStock + ws.Cells(i, 7).Value
                ws.Cells(j, 12).Value = TotalStock
                
                
                bottom = j
                j = j + 1
                TotalClose = 0
                TotalOpen = 0
                TotalStock = 0
            Else:
                TotalOpen = ws.Cells(i, 3).Value + TotalOpen
                TotalClose = ws.Cells(i, 6).Value + TotalClose
                TotalStock = TotalStock + ws.Cells(i, 7).Value
                End If
          Next i
          
          ws.Range("K2:K" & bottom).NumberFormat = "0.00%"
            
            
            For m = 2 To bottom
            '5. Greatest Increase
                If ws.Cells(m, 11).Value > s Then
                    Maxvalue = ws.Cells(m, 11).Value
                    s = ws.Cells(m, 11).Value
                    ticker1 = ws.Cells(m, 9).Value
                End If
                
                '6. Greatest Volume
                
                If ws.Cells(m, 12).Value > t Then
                    Maxtotal = ws.Cells(m, 12).Value
                    t = ws.Cells(m, 12).Value
                    ticker2 = ws.Cells(m, 9).Value
                End If
                
                '7.Greatest decrease
                
                If ws.Cells(m, 11).Value < u Then
                    Minvalue = ws.Cells(m, 11).Value
                    u = ws.Cells(m, 11).Value
                    ticker3 = ws.Cells(m, 9).Value
                End If
                Next m
                
                
                ws.Range("Q2").Value = Maxvalue
                ws.Range("P2").Value = ticker1
                
                ws.Range("Q3").Value = Minvalue
                ws.Range("P3").Value = ticker3
                
                ws.Range("Q4").Value = Maxtotal
                ws.Range("P4").Value = ticker2
            
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("Q3").NumberFormat = "0.00%"
                
                ws.Columns("A:Q").AutoFit
         
    
    Next ws
    
        
        
            
        
    
End Sub
