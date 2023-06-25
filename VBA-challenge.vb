Sub Ticker_Stocks()
    For Each ws In Worksheets
        ws.Activate
        '------------------------------------------------
        ' Name of cells
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change ($)"
        Range("K1").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Cells(1, "P").Value = "Ticker"
        Cells(1, "Q").Value = "Value"
        '-------------------------------------------------
        ' Adding Variables
        Dim TotalVolume As Double
        Dim OpenPointer As Long
        Dim SummaryPointer As Long
        Dim Row_Count As Long
     '-------------------------------------------------
        ' Defining Variables
        TotalVolume = 0
        OpenPointer = 2
        SummaryPointer = 2
        Row_Count = Cells(Rows.Count, "A").End(xlUp).Row
        '-------------------------------------------------
        ' Start Loop
        For i = 2 To Row_Count
        '-------------------------------------------------
            ' Start function
            ' If next cell in column A matches current cell in column A, then do this...
            If Cells(i + 1, "A").Value = Cells(i, "A").Value Then
                TotalVolume = TotalVolume + Cells(i, "G").Value
            
            ' Otherwise, If next cell in column A does NOT match current cell in column A, then do this...
            Else
                
                TotalVolume = TotalVolume + Cells(i, "G").Value
                
                OpenPrice = Cells(OpenPointer, "C").Value
                
                ClosePrice = Cells(i, "F").Value
                
                YearlyChange = ClosePrice - OpenPrice
                
                PercentageChange = YearlyChange / OpenPrice * 100
                
                Cells(SummaryPointer, "I").Value = Cells(i, "A").Value
                
                Cells(SummaryPointer, "J").Value = YearlyChange
                
                Cells(SummaryPointer, "K").Value = PercentageChange & "%"
                
                ' Put in TotalVolume in Cell L2
                Cells(SummaryPointer, "L").Value = TotalVolume
                
                ' Format the cell to display entire numbers
                Cells(SummaryPointer, "L").NumberFormat = "0"
           '-------------------------------------------------
           
            ' Apply conditional formatting based on yearly change
            If YearlyChange > 0 Then
                    
                    Cells(SummaryPointer, "J").Interior.ColorIndex = 4 ' Green
            
            ElseIf YearlyChange < 0 Then
                    
                    Cells(SummaryPointer, "J").Interior.ColorIndex = 3 ' Red
            
            ' End Function
            End If
            
            '-------------------------------------------------
             
             ' Apply conditional formatting based on percentage change
            If PercentageChange > 0 Then
                    
                    Cells(SummaryPointer, "K").Interior.ColorIndex = 4 ' Green
            
            ElseIf PercentageChange < 0 Then
                    
                    Cells(SummaryPointer, "K").Interior.ColorIndex = 3 ' Red
                    
            End If
                    
                TotalVolume = 0
                OpenPointer = i + 1
                SummaryPointer = SummaryPointer + 1
            
            End If

        ' Next iteration
        Next i
        '-------------------------------------------------
        ' Adding Variables
        Dim MaxPercentage As Double
        Dim MaxTicker As String
        Dim MinPercentage As Double
        Dim MinTicker As String
        Dim MaxVolume As Double
        Dim MaxVolumeTicker As String
        Dim SummaryLastRow As Long
        
        SummaryLastRow = Cells(Rows.Count, "I").End(xlUp).Row
        
        ' Find the stock with the greatest percentage increase
        MaxPercentage = Application.WorksheetFunction.Max(Range("K2:K" & SummaryLastRow))
        MaxTicker = WorksheetFunction.Index(Range("I2:I" & SummaryLastRow), WorksheetFunction.Match(MaxPercentage, Range("K2:K" & SummaryLastRow), 0))
        
        ' Write the values to cells P2 and Q2
        Range("P2").Value = MaxTicker
        Range("Q2").Value = Format(MaxPercentage, "0.00%")
        
        ' Find the stock with the greatest percentage decrease
        MinPercentage = Application.WorksheetFunction.Min(Range("K2:K" & SummaryLastRow))
        MinTicker = WorksheetFunction.Index(Range("I2:I" & SummaryLastRow), WorksheetFunction.Match(MinPercentage, Range("K2:K" & SummaryLastRow), 0))
       
       ' Write the values to cells P3 and Q3
        Range("P3").Value = MinTicker
        Range("Q3").Value = Format(MinPercentage, "0.00%")
        
        ' Find the stock with the greatest total volume
        MaxVolume = Application.WorksheetFunction.Max(Range("L2:L" & SummaryLastRow))
        MaxVolumeTicker = WorksheetFunction.Index(Range("I2:I" & SummaryLastRow), WorksheetFunction.Match(MaxVolume, Range("L2:L" & SummaryLastRow), 0))
        
        ' Write the values to cells P4 and Q4
        Range("P4").Value = MaxVolumeTicker
        Range("Q4").Value = MaxVolume
    
    Next ws
End Sub

