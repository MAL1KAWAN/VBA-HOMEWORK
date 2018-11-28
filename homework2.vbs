Public Sub FinalSummary()
    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
        Worksheets(i).Select
        
        
        Formating
        StockSummary
        GreatestIncreaseAndDecrease
    
        i = i + 1
        
    Loop
    
End Sub


Public Sub GreatestIncreaseAndDecrease()
    Dim i As Long
    Dim max As Double
    Dim min As Double
    Dim maxVolume As Double
    
    LastRw = Cells(Rows.Count, 10).End(xlUp).Row
    
    max = WorksheetFunction.max(Range("J:J"))
    min = WorksheetFunction.min(Range("J:J"))
    maxVolume = WorksheetFunction.max(Range("K:K"))
    
        For i = 2 To LastRw
            If Cells(i, 10).Value = max Then
                    Range("O2").Value = Cells(i, 8).Value
                    Range("P2").Value = max
            End If
            
            If Cells(i, 10).Value = min Then
                    Range("O3").Value = Cells(i, 8).Value
                    Range("P3").Value = min

            End If
            
            If Cells(i, 11).Value = maxVolume Then
                    Range("O4").Value = Cells(i, 8).Value
                    Range("P4").Value = maxVolume
            End If
            
            
            If Cells(i, 9).Value = "" Then
                    Cells(i, 9).Interior.ColorIndex = 0
                
            ElseIf Cells(i, 9).Value >= 0 Then
                    Cells(i, 9).Interior.ColorIndex = 4
                    
            ElseIf Cells(i, 9).Value < 0 Then
                    Cells(i, 9).Interior.ColorIndex = 3
            End If
            
    
        Next i
        
End Sub
Public Sub StockSummary()

    Dim i As Long
    Dim RowNumber As Long
    Dim stockVolume As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim OpeningPrice As Double
    Dim closingPrice As Double
    Dim max As Double
    Dim min As Double
    Dim maxVolume As Double
    Dim entryNumber As Long
    
    
    RowNumber = 2

    LastRw = Cells(Rows.Count, 1).End(xlUp).Row
        entryNumber = 0
        For i = 2 To LastRw
            If entryNumber = 0 Then
                OpeningPrice = Cells(i, 3).Value
            End If
            entryNumber = entryNumber + 1
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
                


                closingPrice = Cells(i, 6).Value
                stockVolume = stockVolume + Cells(i, 7).Value
                YearlyChange = closingPrice - OpeningPrice
                
                If OpeningPrice = 0 Then
                PercentChange = 0
                Else
                
                PercentChange = (YearlyChange / OpeningPrice)
                End If
                
                
        
                Range("H" & RowNumber).Value = Cells(i, 1).Value
                Range("K" & RowNumber).Value = stockVolume
                Range("I" & RowNumber).Value = YearlyChange
                Range("j" & RowNumber).Value = PercentChange
                
                    RowNumber = RowNumber + 1
                    stockVolume = 0
                    entryNumber = 0
                
                

            Else
    
                stockVolume = stockVolume + Cells(i, 7).Value
            
            End If
                           
           
        Next i
        
    
    
    
    

End Sub

Sub Formating()
'
' Formating Macro
' This macro add headers and format cells
'

'
    'Sheets("A").Select
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Yearly change"
    Range("J1").Select
    Columns("I:I").EntireColumn.AutoFit
    ActiveCell.FormulaR1C1 = "Percent change"
    Range("K1").Select
    Columns("J:J").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.NumberFormat = "0.00%"
    ActiveCell.FormulaR1C1 = "Total stock volume"
    Range("K2").Select
    Columns("K:K").EntireColumn.AutoFit
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Greatest % increase"
    Range("N3").Select
    ActiveCell.FormulaR1C1 = "Greatest % decrease"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "Greatest total voulme"
    Range("N5").Select
    Columns("N:N").EntireColumn.AutoFit
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("P1").Select
    Range("P2:P3").NumberFormat = "0.00%"
    
    ActiveCell.FormulaR1C1 = "Value"
    Rows("1:1").Select
    Selection.Font.Bold = True
    
    
End Sub


