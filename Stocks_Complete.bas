Sub Stocks()

    Dim FinalCell As String
    Dim Tickers As Variant
    Dim NewRange As String
    Dim UTickers As Variant
    Dim Rows As Integer
    
    Range("A2").Select
    Selection.End(xlDown).Select
    FinalCell = ActiveCell.Address(False, False)
    
    NewRange = "A2:" & FinalCell
    
    Tickers = Range(NewRange)
    
    UTickers = Application.WorksheetFunction.Unique(Tickers)
    
    Rows = UBound(UTickers) - LBound(UTickers) + 1
    
    Dim FirstDate As Double
    Dim LastDate As Double
    Dim CurrentDate As Double
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim TotalVolume As Double
    Dim i As Integer
    Dim CurrentRow As Variant
    
    i = 1
    
    CurrentRow = 2
    
    For i = 1 To Rows
    
        Cells(i + 1, 9).Value = UTickers(i, 1)
        FirstDate = 20180601
        
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
        
        Cells(i + 1, 10).Value = OpenValue
        Cells(i + 1, 11).Value = CloseValue
        Cells(i + 1, 12).Value = CloseValue - OpenValue
        Cells(i + 1, 13).Value = (CloseValue - OpenValue) / OpenValue * 100
        Cells(i + 1, 14).Value = TotalVolume
        
        Next i
        
        
        
End Sub
