Attribute VB_Name = "Module1"
Sub stocks()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        Dim rowcounter As Integer
        Dim ticker As String
        Dim changeval As Double
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim volume As Double
        Dim maxincrease As Double
        Dim maxdecrease As Double
        Dim maxvolume As Double
        Dim maxticker As String
        Dim minticker As String
        Dim maxvolticker As String
        
        
        
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        rowcounter = 2
        volume = 0
        maxincrease = 0
        maxdecrease = 0
        maxvolume = 0
        minticker = ""
        maxvolticker = ""
        
        
        For i = 2 To lastrow
        
            volume = volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                openprice = ws.Cells(i, 3).Value
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & rowcounter).Value = ticker
                rowcounter = rowcounter + 1
                
            End If
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
                ws.Range("J" & rowcounter - 1).Value = yearlychange
                ws.Range("J:J").NumberFormat = "#.00"
                
                If openprice <> 0 Then
                    percentchange = yearlychange / openprice
                    ws.Range("K" & rowcounter - 1).Value = percentchange
                    ws.Range("K:K").NumberFormat = "0.00%"
                    
                    If yearlychange > 0 Then
                        ws.Range("J" & rowcounter - 1).Interior.ColorIndex = 4
                    ElseIf yearlychange < 0 Then
                        ws.Range("J" & rowcounter - 1).Interior.ColorIndex = 3
                    End If
                
                    If percentchange > maxincrease Then
                        maxincrease = percentchange
                        maxticker = ticker
                    End If
                    
                    If percentchange < maxdecrease Then
                        maxdecrease = percentchange
                        minticker = ticker
                    End If
                    
                End If
                
                ws.Range("L" & rowcounter - 1).Value = volume
                
                If volume > maxvolume Then
                    maxvolume = volume
                    maxvolticker = ticker
                End If
                
                volume = 0
                
                
                
            End If
        Next i
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Yearly Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest Percentage Increase"
        ws.Range("O3").Value = "Greatest Percentage Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("P2").Value = maxticker
        ws.Range("P3").Value = minticker
        ws.Range("P4").Value = maxvolticker
        ws.Range("Q2").Value = Format(maxincrease, "0.00%")
        ws.Range("Q3").Value = Format(maxdecrease, "0.00%")
        ws.Range("Q4").Value = maxvolume
        
       ws.Columns("I:L").AutoFit
       ws.Columns("O:Q").AutoFit
        
    Next ws
End Sub


