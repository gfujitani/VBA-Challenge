Attribute VB_Name = "Module1"
Sub Stocks()

For Each ws In Worksheets


        Count = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim total As Double
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestIncreaseT As String
        Dim GreatestDecrease As Double
        Dim GreatestDecreaseT As String
        Dim GreatestVolume As Double
        Dim GreatestVolumeT As String
        total = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        OpeningPrice = ws.Cells(2, 3)
        ClosingPrice = 0
        PercentChange = 0
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Chang"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"

            For i = 2 To lastrow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(Count, 9) = ws.Cells(i, 1)
                    total = ws.Cells(i, 7) + total
                    ws.Cells(Count, 12) = total
            
                    If GreatestVolume < total Then
                        GreatestVolume = total
                        GreatestVolumeT = ws.Cells(i, 1)
                    End If
                
                    ClosingPrice = ws.Cells(i, 6)
                    ws.Cells(Count, 10) = (ClosingPrice - OpeningPrice)
                    If OpeningPrice = 0 Then
                        ws.Cells(Count, 11) = 0
                    Else
                        ws.Cells(Count, 11) = ws.Cells(Count, 10) / OpeningPrice
                    End If
                
                    If GreatestIncrease < ws.Cells(Count, 11) Then
                        GreatestIncrease = ws.Cells(Count, 11)
                        GreatestIncreaseT = ws.Cells(i, 1)
                    
                    End If
                
                    If GreatestDecrease > ws.Cells(Count, 11) Then
                        GreatestDecrease = ws.Cells(Count, 11)
                        GreatestDecreaseT = ws.Cells(i, 1)
                    
                    End If
                OpeningPrice = ws.Cells(i + 1, 3)
                Count = Count + 1
                total = 0
                   
        Else
            total = total + ws.Cells(i, 7)
        End If
     Next i
    ws.Columns("K").Style = "percent"
    
    ws.Cells(2, 16) = GreatestIncreaseT
    ws.Cells(2, 17) = GreatestIncrease
    ws.Cells(3, 16) = GreatestDecreaseT
    ws.Cells(3, 17) = GreatestDecrease
    ws.Cells(4, 16) = GreatestVolumeT
    ws.Cells(4, 17) = GreatestVolume

    Dim rng As Range
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    Set rng = ws.Range("J2:J" & lastrow)
    Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
    With condition1
    .Interior.Color = vbGreen
   End With

   With condition2
     .Interior.Color = vbRed
     
   End With
   
   ws.Range("Q2:Q3").NumberFormat = "0.00%"
   ws.Columns("I:Q").AutoFit
   
Next ws
End Sub


