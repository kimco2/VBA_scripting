Sub bonus_table()

'-------------------------------------------
'BONUS PART 1
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
'---------------------------------------------

'Create the summary table labels
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'Change column width to display text clearly
Columns("N").AutoFit

'---------------------------------------------
'FINDING GREATEST % INCREASE AND DECREASE
'---------------------------------------------

    'Set a variable for the greatest % increase and set initial value
    Dim Greatest_Increase As Double
    Greatest_Increase = Cells(2, 11).Value
    
    'Set a variable for the greatest % decrease and set initial value
    Dim Greatest_Decrease As Double
    Greatest_Decrease = Cells(2, 11).Value
        
    'Set a variable for the number of rows and establish the value
    Dim LastRow_sum_table As Long
    LastRow_sum_table = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through all yearly change values
         Dim l As Long
         For l = 2 To LastRow_sum_table

            'Check if the cells value is greater than the Max value, if it is...
            If Cells(l, 11).Value > Greatest_Increase Then
              
            'Store this value as the greatest % increase
            Greatest_Increase = Cells(l, 11)
                  
            'Retrieve ticker name and value associated with the Greatest % Increase, print into summary table and format the value
            Range("O2").Value = Cells(l, 9).Value
            Range("P2").Value = Cells(l, 11).Value
            Range("P2").NumberFormat = "0.00%"
                
            'If the value is smaller, see whether it is smaller than current Greatest_Decrease value, if it is...
            ElseIf Cells(l, 11).Value < Greatest_Decrease Then
            
            'Store this value as the greastest % decrease
            Greatest_Decrease = Cells(l, 11)
                
            'Retrieve ticker name and value associated with the Greatest % decrease, print to summary table, and format the value
            Range("O3").Value = Cells(l, 9).Value
            Range("P3").Value = Cells(l, 11).Value
            Range("P3").NumberFormat = "0.00%"
                                        
            End If
               
        Next l

'---------------------------------------------
'FINDING GREATEST TOTAL VOLUME
'---------------------------------------------
  
    'Set a variable for total greatest volume and set an initial value
    Dim volume As Double
    volume = Cells(2, 12).Value
               
        'Loop through all total stock volumes
         For m = 2 To LastRow_sum_table

             'Check if the cell value is greater than volume value, if it is...
              If Cells(m, 12).Value > volume Then
              
              'Store this higher value as the  volume
              volume = Cells(m, 12)
                  
              'Retrieve ticker name and value associated with the greatest total volume, print to summary table, and format the value
              Range("O4").Value = Cells(m, 9).Value
              Range("P4").Value = Cells(m, 12).Value
              Range("P4").NumberFormat = "0"
                
            End If
        
        Next m
End Sub




