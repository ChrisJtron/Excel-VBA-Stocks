Attribute VB_Name = "Module1"
Sub stock_project()

Dim ticker As String
Dim change As Double
Dim ticker_row As Integer
Dim count As Double
Dim change_row As Integer


count = 0

ticker_row = 2

change_row = 2


  For r = 2 To Range("a2").CurrentRegion.End(xlDown).Row

    'For r = 2 To 2000
    
    

    If Cells(r + 1, 1) <> Cells(r, 1) Then
    
    'next row for next ticker
    
        ticker_row = ticker_row + 1
        
    'next row for next ticker's change
        
        change_row = change_row + 1
        
    'reset count for next ticker
        
        count = 0
        
        
        
    'else portion works correctly
        
    Else
    
    'counter works
    
        change = Cells(r + 1, 6).Value - Cells(r, 6).Value
    
        'change_row = 2
        
        count = count + change
        
        Cells(change_row, 10).Value = count
        
    'add ticker symbol
    
        Cells(ticker_row, 9).Value = Cells(r, 1).Value
           
        
     End If
     
        Next r
     
     
     
     
   
    

End Sub



Sub percent_change()

    Dim first_open As Double
    Dim yearly_change As Double
    Dim j_row As Integer
    Dim percent As Variant
    
    

        j_row = 2
        
        'for a = 2 to 2000

        For a = 1 To Range("a2").CurrentRegion.End(xlDown).Row
    

    'find first row of ticker
         If Cells(a, 1) <> Cells(a + 1, 1) Then
    
    
    'extract opening price on first day of ticker
        first_open = Cells(a + 1, 3).Value
        


    'define yearly change for ticker
        
            yearly_change = Cells(j_row, 10).Value
            
            
            If first_open = 0 Then
                Cells(j_row, 11).Value = "N/A"
            
            
            ElseIf yearly_change <> 0 Then
                

    'divide yearly change by opening price
             percent = yearly_change / first_open
             
    'put percent in row K
          Cells(j_row, 11).Value = percent
        
        
    'format column k to percent
         Range("k:k").NumberFormat = "0.00%"
         
         
            'If first_open = 0 Then
                'Cells(j_row, 11).Value = "N/A"
         
         End If
        
    'add 1 to j-row for next ticker
            j_row = j_row + 1
            
            
        
            
        Else
        
    End If
    
    Next a
    
            




End Sub



