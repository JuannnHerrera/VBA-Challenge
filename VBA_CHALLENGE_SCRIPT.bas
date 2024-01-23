Attribute VB_Name = "Module1"
Sub VBA_CHALLENGE_SCRIPT()

    Dim ws As Worksheet
    Dim ColumnI As String
    Dim ColumnJ As String
    Dim ColumnK As String
    Dim ColumnL As String
    Dim YearlyChange As Double
    Dim TotalVolume As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim TickerIncrease As String
    Dim TickerDecrease As String
    Dim TickerVolume As String
    Dim LastRow As Long
    Dim rng As Range
    Dim cell As Range

    ' Added to autorun on all worksheetsa
     For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        Range("A2:A22771").AdvancedFilter _
            Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True

        ColumnI = "Ticker"
        ColumnJ = "Yearly Change"
        ColumnK = "Percentage Yield"
        ColumnL = "Total Stock Volume"
        Cells(1, 9).Value = ColumnI
        Cells(1, 10).Value = ColumnJ
        Cells(1, 11).Value = ColumnK
        Cells(1, 12).Value = ColumnL

        ' First Loop by Ticker symbol
        For i = 2 To Range("I" & Rows.Count).End(xlUp).Row
            Ticker = Cells(i, 9).Value
            YearlyChange = 0
            TotalVolume = 0
            OpeningPrice = 0
            LastClosingPrice = 0
            LastClosingFound = False

            ' Column J Column K by Ticker
            For j = 2 To 22771
                If Cells(j, 1).Value = Ticker Then
                
                    If OpeningPrice = 0 Then
                        OpeningPrice = Cells(j, 3).Value
                        
                    End If
                    ClosingPrice = Cells(j, 6).Value
                    YearlyChange = ClosingPrice - OpeningPrice
                    TotalVolume = TotalVolume + Cells(j, 7).Value
                    
                    If Cells(j, 2).Value = Cells(i, 9).Value Then
                        LastClosingPrice = ClosingPrice
                        LastClosingFound = True
                        
                    End If
                    
                End If
                
            Next j

            ' Column K Percentage Yield
            If OpeningPrice <> 0 Then
                PercentageYield = (YearlyChange / OpeningPrice) * 100
                
            Else
                PercentageYield = 0
                
            End If

            ' Output results by ticker needs to stay within loop
            Cells(i, 10).Value = YearlyChange
            Cells(i, 11).Value = PercentageYield
            Cells(i, 12).Value = TotalVolume
            
        Next i

        ' Find the last row in column J ('Yearly Change')
        LastRow = Cells(Rows.Count, "J").End(xlUp).Row
        
        ' Define the range to apply conditional formatting
        Set rng = Range("J2:J" & LastRow)
        
        ' Clear any existing conditional formatting
        rng.FormatConditions.Delete
        
        ' Red fill color
        rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        rng.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        
        ' Green fill color
        rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        rng.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
        
        ' Second Loop for second chart Columns O, P, Q
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0

        For i = 2 To Range("I" & Rows.Count).End(xlUp).Row
            Ticker = Cells(i, 9).Value
            PercentageYield = Cells(i, 11).Value
            TotalVolume = Cells(i, 12).Value

            If PercentageYield > MaxIncrease Then
                MaxIncrease = PercentageYield
                TickerIncrease = Ticker
            End If

            If PercentageYield < MaxDecrease Then
                MaxDecrease = PercentageYield
                TickerDecrease = Ticker
            End If

            If TotalVolume > MaxVolume Then
                MaxVolume = TotalVolume
                TickerVolume = Ticker
            End If
            
        Next i

        ' Create all the labels
        Cells(2, 16).Value = TickerIncrease
        Cells(2, 17).Value = MaxIncrease
        Cells(3, 16).Value = TickerDecrease
        Cells(3, 17).Value = MaxDecrease
        Cells(4, 16).Value = TickerVolume
        Cells(4, 17).Value = MaxVolume
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
    Next ws
    
End Sub


