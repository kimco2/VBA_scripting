Sub VBA_challenge_Bonus_Part_2()

'--------------------------------------------------------------------------------------------
'BONUS - PART 2 - MAKE ADJUSTS TO YOUR SCRIPT SO IT RUNS ON EVERY WORKSHEET
'--------------------------------------------------------------------------------------------

For Each ws In Worksheets

'Name the columns for the output table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Change column width to display text clearly
ws.Columns("J").AutoFit
ws.Columns("K").AutoFit
ws.Columns("L").AutoFit

'Set a variable for holding the ticker name
Dim ticker_name As String

'Set a variable for holding the total volume of the ticker and set the value to 0
Dim ticker_volume_total As Double
ticker_volume_total = 0

'Set a variable for holding the opening value for a ticker
Dim opening_value As Double

'Set a variable for holding the closing vaue for a ticker
Dim closing_value As Double

'Set a variable to count the number of rows and establish the value
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set a variable for keeping count of the number of rows for each ticker and set the value to 0
Dim row_count As Long
row_count = 0

'Set a variable to keep track of the location of each ticker name in the output table and set the value to 2
Dim output_table_row As Integer
output_table_row = 2

        'Loop through all tickers
        Dim i As Long
        For i = 2 To LastRow

        'Check if we are still within the same ticker name, if not...
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Set the ticker name
            ticker_name = ws.Cells(i, 1).Value
         
            'Print the ticker name in the ouput table
            ws.Range("I" & output_table_row).Value = ticker_name
         
            'Add to the ticker volume total
            ticker_volume_total = ticker_volume_total + ws.Cells(i, 7).Value
                     
            'Print the ticker volume total in the output table and format the number
            ws.Range("L" & output_table_row).Value = ticker_volume_total
            ws.Range("l" & output_table_row).NumberFormat = 0
            
            'Find the opening value for the ticker
            opening_value = ws.Cells(i - row_count, 3).Value
                     
            'Find closing value for ticker
            closing_value = ws.Cells(i, 6).Value
                     
            'Calculate yearly change and print value in output table
            yearly_change = closing_value - opening_value
            ws.Range("J" & output_table_row).Value = yearly_change
            ws.Range("J" & output_table_row).NumberFormat = "0.00"
           
            'Calcuate percent change, print value in the output table, and format the number
            percent_change = yearly_change / opening_value
            ws.Range("K" & output_table_row).Value = percent_change
            ws.Range("K" & output_table_row).NumberFormat = "0.00%"
                                 
            'Add one to the output table row
            output_table_row = output_table_row + 1
    
            'Reset the ticker total volume
            ticker_volume_total = 0
           
            'Reset row count to 0
             row_count = 0
                      
                'If the cell immediately following a row is the same brand, then...
                Else
    
                  'Add to the ticker volume total
                  ticker_volume_total = ticker_volume_total + ws.Cells(i, 7).Value
         
                 'Add one to the row count
                  row_count = row_count + 1
              
       
           End If
           
      Next i

'---------------------------------------------------------------
'CONDITIONAL FORMAT THE YEARLY CHANGE AND PERCENT CHANGE COLUMN
'--------------------------------------------------------------

'Set a variable to count the number of rows in the output table and establish the value
Dim LastRow_table As Long
LastRow_table = ws.Cells(Rows.Count, 1).End(xlUp).Row

           'Loop through all yearly change values
            Dim j As Long
            For j = 2 To LastRow_table
             
             'Check whether yearly change is greater than 0, if it is...
              If ws.Cells(j, 10).Value > 0 Then
            
                    'Change the colour of the cell to green
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    ws.Cells(j, 11).Interior.ColorIndex = 4
            
              'Check if the yearly change is less than 0, if it is...
              ElseIf ws.Cells(j, 10).Value < 0 Then
         
                    'Change the colour of the cell to red
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    ws.Cells(j, 11).Interior.ColorIndex = 3
    
               End If
               
        Next j
                          
'-------------------------------------------
'BONUS PART 1
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
'---------------------------------------------

'Create the summary table labels
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Change column width to display text clearly
ws.Columns("N").AutoFit

'Set a variable for the number of rows and establish the value
Dim LastRow_sum_table As Long
LastRow_sum_table = ws.Cells(Rows.Count, 1).End(xlUp).Row

'GREATEST % INCREASE
'-------------------
'Set a variable for the greatest % increase and set initial value
Dim Greatest_Increase As Double
Greatest_Increase = ws.Cells(2, 11).Value
            
    'Loop through all percent change values
    Dim l As Long
    For l = 2 To LastRow_sum_table
            
        'Check if the cell value is larger than the current greatest % increase value, if it is...
        If ws.Cells(l, 11).Value > Greatest_Increase Then
              
        'Store this value as the greatest % increase
        Greatest_Increase = ws.Cells(l, 11).Value
                  
        'Retrieve ticker name and value associated with the greatest % Increase, print into summary table and format the value
        ws.Range("O2").Value = ws.Cells(l, 9).Value
        ws.Range("P2").Value = ws.Cells(l, 11).Value
        ws.Range("P2").NumberFormat = "0.00%"
              
        'Otherwise if the initial value is the greatest % increase, then...
        ElseIf Greatest_Increase <= ws.Cells(2, 11).Value Then
        
        'Retrieve ticker name and value associated with the greatest % Increase, print into summary table and format the value
        ws.Range("O2").Value = ws.Cells(2, 9).Value
        ws.Range("P2").Value = ws.Cells(2, 11).Value
        ws.Range("P2").NumberFormat = "0.00%"
              
        End If
            
    Next l

'GREATEST % DECREASE
'-------------------
'Set a variable for the greatest % decrease and set initial value
Dim Greatest_Decrease As Double
Greatest_Decrease = ws.Cells(2, 11).Value
            
    'Loop through all percent change values
    Dim m As Long
    For m = 2 To LastRow_sum_table
            
        'Check if the cell value is less than the current greatest % decresae value, if it is...
        If ws.Cells(m, 11).Value < Greatest_Decrease Then
              
        'Store this value as the greatest % decrease
        Greatest_Decrease = ws.Cells(m, 11).Value
                  
        'Retrieve ticker name and value associated with the greatest % Decrease, print into summary table and format the value
        ws.Range("O3").Value = ws.Cells(m, 9).Value
        ws.Range("P3").Value = ws.Cells(m, 11).Value
        ws.Range("P3").NumberFormat = "0.00%"
        
        'Otherwise if the initial value is the greatest % decrease, then...
        ElseIf Greatest_Decrease >= ws.Cells(2, 11).Value Then
            
       'Retrieve ticker name and value associated with the greatest % decrease, print to summary table, and format the value
        ws.Range("O3").Value = ws.Cells(2, 9).Value
        ws.Range("P3").Value = ws.Cells(2, 11).Value
        ws.Range("P3").NumberFormat = "0.00%"
                                        
        End If
            
    Next m
        
'GREATEST TOTAL VOLUME
'---------------------
'Set a variable for the greatest total volume and set an initial value
Dim Greatest_Volume As Double
Greatest_Volume = ws.Cells(2, 12).Value
            
    'Loop through all total volume values
    Dim n As Long
    For n = 2 To LastRow_sum_table
            
        'Check if the cell value is larger than the current greatest total volume, if it is...
        If ws.Cells(n, 12).Value > Greatest_Volume Then
              
        'Store this value as the greatest total volume
        Greatest_Volume = ws.Cells(n, 12).Value
                  
        'Retrieve ticker name and value associated with the greatest total volume, print into summary table
        ws.Range("O4").Value = ws.Cells(n, 9).Value
        ws.Range("P4").Value = ws.Cells(n, 12).Value
            
        'Otherwise if the intital value is the greatest total volume, then...
        ElseIf Greatest_Volume <= ws.Cells(2, 12) Then
            
       'Retrieve ticker name and value associated with the greatest total volume print to summary table
        ws.Range("O4").Value = ws.Cells(2, 9).Value
        ws.Range("P4").Value = ws.Cells(2, 12).Value
                                                           
        End If
            
    Next n
    
  Next ws

End Sub
    





