Attribute VB_Name = "Module1"
Sub VBAChallenge()
'Instructions
    'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.
    'You should also have conditional formatting that will highlight positive change in green and negative change in red.
    'The result should look as follows.

'CHALLENGES
'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
'Other Considerations
'Use the sheet alphabetical_testing.xlsx while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.
'Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.
    
    For Each ws In Worksheets
    
    ' Set initial variables for holding the open value, the close value, the yearly change and the percent change
    Dim OpenVal, CloseVal, YearlyChange, PercentChange As Double
    
    ' Set an initial variable for holding the ticker
    Dim Ticker As String
    
    ' Set an initial variable for holding the total stock volume for each ticker
    Dim TotalStockVol As Double
    TotalStockVol = 0
    
    'Set an initial variable for the first row number of the ticker
    Dim First_B_RowNumber As Long
    First_Ticker_RowNumber = 2
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Add Headers.
    ws.Range("j1").Value = "Ticker"
    ws.Range("k1").Value = "Yearly Change"
    ws.Range("l1").Value = "Percent Change"
    ws.Range("m1").Value = "Total Stock Volume"
    
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"
    
    
    'CREATE A SCRIPT THAT WILL LOOP THROUGH ALL THE STOCKS FOR ONE YEAR AND OUTPUT THE FOLLOWING INFORMATION.
        'The ticker symbol.
        'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        'The total stock volume of the stock.
    
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all rows
    For i = 2 To LastRow
        
    
        ' Check if we are still within the same ticker for all other values, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Set the ticker name
            Ticker = ws.Cells(i, 1).Value
            
            'Set the opening value
            With ActiveWorkbook.ActiveSheet
                Set First_Ticker_Row = ws.Range("A:A").Find(What:=Ticker, LookIn:=xlValues)
            End With

            First_Ticker_RowNumber = First_Ticker_Row.Row
            
            ActiveWorkbook.ActiveSheet.Rows(First_Ticker_RowNumber).Select
            
            OpenVal = ws.Cells(First_Ticker_RowNumber, 3)
            
            'Set Close Value
            CloseVal = ws.Cells(i, 6).Value
        
            'Calculate the difference between the opening value at the beginning of the year and the closing value at the end of the year.
            YearlyChange = CloseVal - OpenVal
            
            'Calculate the percent change
            If OpenVal = 0 Then
                OpenVal = PercentChange = 0
                
            Else
                PercentChange = YearlyChange / OpenVal
            
            End If
            
            ' Add to the total stock volume
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
            ' Print the ticker name in the summary table
            ws.Range("j" & Summary_Table_Row).Value = Ticker
            
            ' Print the ticker name in the summary table
            ws.Range("k" & Summary_Table_Row).Value = YearlyChange
            
            ' Print the percent change to the summary table
            ws.Range("l" & Summary_Table_Row).Value = PercentChange
                        
            ' Print the total stock volume to the summary table
            ws.Range("m" & Summary_Table_Row).Value = TotalStockVol

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the total stock volume
            TotalStockVol = 0
        
        ' If the cell immediately following a row is the same ticker...
        Else
        
        'The total stock volume of the stock.
        TotalStockVol = TotalStockVol + Cells(i, 7).Value
                
        End If
   
    Next i
   
   
   'FORMAT THE SUMMARY TABLE
      
      ' Determine the last row of the summary table
    LastRow2 = ws.Cells(Rows.Count, 12).End(xlUp).Row
 
   
       ' Loop through all rows
    For i = 2 To LastRow2
        
    
        'for those that are negative
        If ws.Cells(i, 12).Value < 0 Then
            
            ' Turn cell red
            ws.Cells(i, 12).Interior.ColorIndex = 3
       
        ' for those that are positive
        ElseIf ws.Cells(i, 12).Value > 0 Then
                    
            'turn cell green
            ws.Cells(i, 12).Interior.ColorIndex = 4
                
        'if the change is 0, leave them as is
        Else
        
        End If
        
        'Change the "Percent Change" to percent
        ws.Cells(i, 12).NumberFormat = "0.00%"
        
         
    Next i
    

'CHALLENGES
    'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
    
    ' Determine the last row of the summary table
    LastRow3 = ws.Cells(Rows.Count, 12).End(xlUp).Row
 
    'set initial values for the Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    Dim GreatInc, GreatDec, GreatTotVol As Double
    GreatInc = -1
    GreatDec = 1
    GreatTotVol = 0
     
    ' Loop through all the rows of the summary table to get the greatest increase, decrease and total volume
    For i = 2 To LastRow3
           
        'for those that are less than the current value of greatest increase
        If ws.Cells(i, 12).Value < GreatInc Then
            
            ' leave value of greatest increase as is
            GreatInc = GreatInc
                    
        ' change the value to greatest increase and put it and ticker in the table
        Else
                    
            'replace the current greatest increase value with the current value
            GreatInc = ws.Cells(i, 12).Value
       
        ' Print the ticker and the value to the table
        ws.Range("q2").Value = GreatInc
        
        ' Print the percent change to the summary table
        ws.Range("p2").Value = ws.Cells(i, 10).Value

        End If
                 
                 
        'for those that are greater than the current value of greatest decrease
        If ws.Cells(i, 12).Value > GreatDec Then
            
            ' leave value of greatest decrease as is
            GreatDec = GreatDec
                    
        ' change the value to greatest decrease and put it and ticker in the table
        Else
                    
            'replace the current greatest decrease value with the current value
            GreatDec = ws.Cells(i, 12).Value
       
        ' Print the ticker and the value to the table
        ws.Range("q3").Value = GreatDec
        
        ' Print the percent change to the summary table
        ws.Range("p3").Value = ws.Cells(i, 10).Value

        End If
                  
                  
        'for those that are greater than current value of greatest total volume
        If ws.Cells(i, 13).Value < GreatTotVol Then
            
            ' leave value of greatest total volume as is
            GreatTotVol = GreatTotVol
                    
        ' change the value to greatest total volume and put it and ticker in the table
        Else
                    
            'replace the current greatest value with the current value
            GreatTotVol = ws.Cells(i, 13).Value
       
        ' Print the ticker and the value to the table
        ws.Range("q4").Value = GreatTotVol
        
        ' Print the total volume to the summary table
        ws.Range("p4").Value = ws.Cells(i, 10).Value

        End If
        
 
        
    Next i
    
    'autofit the width of the columns of the summary table
    ws.Columns("J:q").AutoFit
    
    'Change the "Percent Change" to percent for the greatest increase and decrease
    ws.Range("q2:q3").NumberFormat = "0.00%"

    'have the script end in the final table
    Range("n4").End(xlToRight).Select
    
    'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

    
Next ws

End Sub
  
      
