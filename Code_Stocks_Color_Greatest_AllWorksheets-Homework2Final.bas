Sub Stocks_Greatest_Color_AllWorksheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

    Dim FinalCell As String
    Dim Tickers As Variant
    Dim NewRange As String
    Dim UTickers As Variant
    Dim NumRows As Integer
    
    Range("A2").Select
    Selection.End(xlDown).Select
    FinalCell = ActiveCell.Address(False, False)
    
    NewRange = "A2:" & FinalCell
    
    Tickers = Range(NewRange)
    
    UTickers = Application.WorksheetFunction.Unique(Tickers)
    
    NumRows = UBound(UTickers) - LBound(UTickers) + 1
    
    Dim FirstDate As Double
    Dim LastDate As Double
    Dim CurrentDate As Double
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim TotalVolume As Double
    Dim i As Integer
    Dim CurrentRow As Variant
    Dim YearlyChange As Double
    
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxTotalVolume As Double
    Dim MaxPercentIncreaseTicker As String
    Dim MaxPercentDecreaseTicker As String
    Dim MaxTotalVolumeTicker As String
    
    MaxPercentIncrease = 0
    MaxPercentDecrease = 0
    MaxTotalVolume = 0
    
    i = 1
    
    CurrentRow = 2
    
    For i = 1 To NumRows
    
        Cells(i + 1, 9).Value = UTickers(i, 1)
        FirstDate = Range("B2").Value
        TotalVolume = 0
        
        Do While Cells(CurrentRow, 1).Value = UTickers(i, 1)
        
            If Cells(CurrentRow, 1).Value = UTickers(i, 1) Then
            
                CurrentDate = Cells(CurrentRow, 2).Value
                
                    If CurrentDate <= FirstDate Then
                    
                        FirstDate = CurrentDate
                        OpenValue = Cells(CurrentRow, 3).Value
                        
                    End If
                    
                    If CurrentDate >= LastDate Then
                    
                        LastDate = CurrentDate
                        CloseValue = Cells(CurrentRow, 6).Value
                        
                    End If
            
            TotalVolume = TotalVolume + Cells(CurrentRow, 7).Value
                
            End If
            
            CurrentRow = CurrentRow + 1
            
        Loop
        
        YearlyChange = CloseValue - OpenValue
        
        If OpenValue = 0 Then
            PercentChange = 0
        
        Else
            PercentChange = YearlyChange / OpenValue * 100
        
        End If
        
        Cells(i + 1, 10).Value = OpenValue
        Cells(i + 1, 11).Value = CloseValue
        Cells(i + 1, 12).Value = YearlyChange
        Cells(i + 1, 13).Value = PercentChange
        Cells(i + 1, 14).Value = TotalVolume
        
        If PercentChange > MaxPercentIncrease Then
            
            MaxPercentIncrease = PercentChange
            MaxPercentIncreaseTicker = UTickers(i, 1)
        
        End If
        
        If PercentChange < MaxPercentDecrease Then
            
            MaxPercentDecrease = PercentChange
            MaxPercentDecreaseTicker = UTickers(i, 1)
        
        End If
        
        If TotalVolume > MaxTotalVolume Then
            
            MaxTotalVolume = TotalVolume
            MaxTotalVolumeTicker = UTickers(i, 1)
        
        End If
        
        Cells(2, 18).Value = MaxPercentIncrease
        Cells(3, 18).Value = MaxPercentDecrease
        Cells(4, 18).Value = MaxTotalVolume
        Cells(2, 17).Value = MaxPercentIncreaseTicker
        Cells(3, 17).Value = MaxPercentDecreaseTicker
        Cells(4, 17).Value = MaxTotalVolumeTicker
        
        Next i
    
        Dim cell As Range
        
            For Each cell In Range("L2:L" & Cells(Rows.Count, "A").End(xlUp).Row)

                If cell.Value < 0 Then
                    cell.Interior.Color = vbRed
    
                ElseIf cell.Value > 0 Then
                    cell.Interior.Color = vbGreen
    
                Else
                    cell.Interior.Color = vbWhite
    
                End If

        Next cell
        
    Next ws
    
End Sub

