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
    Range("j1").Value = "Ticker"
    Range("k1").Value = "Yearly Change"
    Range("l1").Value = "Percent Change"
    Range("m1").Value = "Total Stock Volume"
    
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    
    
    'Create a script that will loop through all the stocks for one year and output the following information.
    
    ' Determine the Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all rows
    For i = 2 To LastRow
        
    
        ' Check if we are still within the same ticker for all other values, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set the ticker name
            Ticker = Cells(i, 1).Value
            
            'Set the opening value
            With ActiveWorkbook.ActiveSheet
                Set First_Ticker_Row = .Range("A:A").Find(What:=Ticker, LookIn:=xlValues)
            End With

            First_Ticker_RowNumber = First_Ticker_Row.Row
            
            ActiveWorkbook.ActiveSheet.Rows(First_Ticker_RowNumber).Select
            
            OpenVal = Cells(First_Ticker_RowNumber, 3)
            
            'Set Close Value
            CloseVal = Cells(i, 6).Value
        
            'Calculate the difference between the opening value at the beginning of the year and the closing value at the end of the year.
            YearlyChange = CloseVal - OpenVal
            
            'Calculate the percent change
            PercentChange = YearlyChange / OpenVal
            
            ' Add to the total stock volume
            TotalStockVol = TotalStockVol + Cells(i, 7).Value
            
            ' Print the ticker name in the summary table
            Range("j" & Summary_Table_Row).Value = Ticker
            
            ' Print the ticker name in the summary table
            Range("k" & Summary_Table_Row).Value = YearlyChange
            
            ' Print the percent change to the summary table
            Range("l" & Summary_Table_Row).Value = PercentChange
                        
            ' Print the total stock volume to the summary table
            Range("m" & Summary_Table_Row).Value = TotalStockVol

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
   
End Sub
  

      ' Add to the Total Stock Volume
   'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
       
    
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    
       
        'OpenVal , CloseVal, YearlyChange, PercentChange

  
    ' If the cell immediately following a row is the same ticker...
    
    
    
    'You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    

    
    

'CHALLENGES
    'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
    
    
    
    
    'Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

