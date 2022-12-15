Attribute VB_Name = "Module2"
Sub VBA_Challenge_Bonus_Part_1()

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

'Set a variable for the number of rows and establish the value
Dim LastRow_sum_table As Long
LastRow_sum_table = Cells(Rows.Count, 1).End(xlUp).Row

'GREATEST % INCREASE
'-------------------
'Set a variable for the greatest % increase and set initial value
Dim Greatest_Increase As Double
Greatest_Increase = Cells(2, 11).Value
            
    'Loop through all percent change values
    Dim l As Long
    For l = 2 To LastRow_sum_table
            
        'Check if the cell value is larger than the current greatest % increase value, if it is...
        If Cells(l, 11).Value > Greatest_Increase Then
              
        'Store this value as the greatest % increase
        Greatest_Increase = Cells(l, 11).Value
                  
        'Retrieve ticker name and value associated with the greatest % Increase, print into summary table and format the value
        Range("O2").Value = Cells(l, 9).Value
        Range("P2").Value = Cells(l, 11).Value
        Range("P2").NumberFormat = "0.00%"
              
        'Otherwise if the initial value is the greatest % increase, then...
        ElseIf Greatest_Increase <= Cells(2, 11).Value Then
        
        'Retrieve ticker name and value associated with the greatest % Increase, print into summary table and format the value
        Range("O2").Value = Cells(2, 9).Value
        Range("P2").Value = Cells(2, 11).Value
        Range("P2").NumberFormat = "0.00%"
              
        End If
            
    Next l

'GREATEST % DECREASE
'-------------------
'Set a variable for the greatest % decrease and set initial value
Dim Greatest_Decrease As Double
Greatest_Decrease = Cells(2, 11).Value
            
    'Loop through all percent change values
    Dim m As Long
    For m = 2 To LastRow_sum_table
            
        'Check if the cell value is less than the current greatest % decresae value, if it is...
        If Cells(m, 11).Value < Greatest_Decrease Then
              
        'Store this value as the greatest % decrease
        Greatest_Decrease = Cells(m, 11).Value
                  
        'Retrieve ticker name and value associated with the greatest % Decrease, print into summary table and format the value
        Range("O3").Value = Cells(m, 9).Value
        Range("P3").Value = Cells(m, 11).Value
        Range("P3").NumberFormat = "0.00%"
        
        'Otherwise if the initial value is the greatest % decrease, then...
        ElseIf Greatest_Decrease >= Cells(2, 11).Value Then
            
       'Retrieve ticker name and value associated with the greatest % decrease, print to summary table, and format the value
        Range("O3").Value = Cells(2, 9).Value
        Range("P3").Value = Cells(2, 11).Value
        Range("P3").NumberFormat = "0.00%"
                                        
        End If
            
    Next m
        
'GREATEST TOTAL VOLUME
'---------------------
'Set a variable for the greatest total volume and set an initial value
Dim Greatest_Volume As Double
Greatest_Volume = Cells(2, 12).Value
            
    'Loop through all total volume values
    Dim n As Long
    For n = 2 To LastRow_sum_table
            
        'Check if the cell value is larger than the current greatest total volume, if it is...
        If Cells(n, 12).Value > Greatest_Volume Then
              
        'Store this value as the greatest total volume
        Greatest_Volume = Cells(n, 12).Value
                  
        'Retrieve ticker name and value associated with the greatest total volume, print into summary table, and format value
        Range("O4").Value = Cells(n, 9).Value
        Range("P4").Value = Cells(n, 12).Value
            
        'Otherwise if the intital value is the greatest total volume, then...
        ElseIf Greatest_Volume <= Cells(2, 12) Then
            
        'Otherwise if the intital value is the greatest total volume, then ...
        'Retrieve ticker name and value associated with the greatest total volume print to summary table, and format value
        Range("O4").Value = Cells(2, 9).Value
        Range("P4").Value = Cells(2, 12).Value
                                                           
        End If
            
    Next n

End Sub
    


