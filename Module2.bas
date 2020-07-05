Attribute VB_Name = "Module2"
 
  
  Sub color_change()
  
     
     For q = 2 To Range("j2").CurrentRegion.End(xlDown).Row
     
        If Cells(q, 10).Value < 0 Then
     
            Cells(q, 10).Interior.ColorIndex = 53
        
        
        ElseIf Cells(q, 10).Value > 0 Then
        
            Cells(q, 10).Interior.ColorIndex = 10
        
        Else
        
            Cells(q, 10).Interior.ColorIndex = normal
            
        End If
        
        Next q
        
    
    
End Sub

Sub volume()

'integer
Dim volume As Double

count = 0

Row = 2



  For r = 2 To Range("a2").CurrentRegion.End(xlDown).Row

    'For r = 2 To 2000
    
    

    If Cells(r, 1) <> Cells(r + 1, 1) Then
    
    'next row for next ticker
    
        Row = Row + 1
        
        
    'reset count for next ticker
        
        count = 0
        
        
    Else
    
    'counter   IS MISSING LAST VALUE OF Ticker
    
        volume = Cells(r + 1, 7).Value
        
        count = count + volume
        
        Cells(Row, 12).Value = count
        
           
     End If
     
        Next r
     
End Sub




