Attribute VB_Name = "Module3"

Sub most_increase()

    Dim max As Double

    
    max = Application.WorksheetFunction.max(Range("k:k"))
    
    Cells(2, 15).Value = max
    
    Cells(2, 14).Value = "Greatest % Increase"
    
    Cells(2, 15).NumberFormat = "0.00%"
    
        For i = 2 To Range("k2").CurrentRegion.End(xlDown).Row
        
            If Cells(i, 11).Value = Cells(2, 15).Value Then
            
                Cells(2, 16).Value = Cells(i, 9).Value
                
            End If
            
        Next i
    
    
End Sub


Sub most_decrease()

    Dim min As Double
    
    
    min = Application.WorksheetFunction.min(Range("k:k"))
    
    Cells(3, 15).Value = min
    
    Cells(3, 14).Value = "Greatest % Decrease"
    
    Cells(3, 15).NumberFormat = "0.00%"
    
    
        For i = 2 To Range("k2").CurrentRegion.End(xlDown).Row
        
            If Cells(i, 11).Value = Cells(3, 15).Value Then
            
                Cells(3, 16).Value = Cells(i, 9).Value
                
            End If
            
        Next i

End Sub


Sub greatest_volume()

    Dim vol As Double
    
    vol = Application.WorksheetFunction.max(Range("l:l"))
    
    Cells(4, 15).Value = vol
    
    Cells(4, 14).Value = "Greatest Total Volume"
    
    Cells(4, 15).NumberFormat = "0"
    
    
        For i = 2 To Range("l2").CurrentRegion.End(xlDown).Row
        
            If Cells(i, 12).Value = Cells(4, 15).Value Then
            
                Cells(4, 16).Value = Cells(i, 9).Value
                
            End If
            
        Next i
    
End Sub


Sub step1()

    stock_project
    
    color_change
    
    percent_change

    
End Sub

Sub step2()

    volume
    
    most_increase
    
    most_decrease
    
    greatest_volume
    
    add_headers
    
    Columns("a:o").AutoFit
    
End Sub

Sub add_headers()

    Cells(1, 9).Value = "Ticker Symbol"
    
    Cells(1, 10).Value = "Yearly Change"
    
    Cells(1, 11).Value = "Percent Change"
    
    Cells(1, 12).Value = "Volume"
    
    Range("a1:l1").Interior.ColorIndex = 37
    
    Range("a1:l1").Font.Bold = True
    
    
End Sub

