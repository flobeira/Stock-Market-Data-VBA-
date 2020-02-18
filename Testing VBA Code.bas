Attribute VB_Name = "Module1"
Sub Testing()
    'Loop through all worksheets
    For Each ws In Worksheets
        'Variable created to refer to each Ticket Symbol
        Dim Stock_Ticket As String
        'Variable created to refer to Opening Price
        Dim Opening_Price As Double
            Opening_Price = 0
        'Variable created to refer to Closing Price
        Dim Closing_Price As Double
            Closing_Price = 0
        'Variable Created to refer to the Stock Volume
        Dim Stock_Volume As Double
            Stock_Volume = 0
        'Variable created to go down rows when creating the table
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
        'Determine Last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop Through all Stocks
        For i = 2 To LastRow
        
            'Check Every time the loop passes to a different Stock Ticket
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Aknowledge each time the Ticket Symbol changes for i
                Stock_Ticket = ws.Cells(i, 1).Value
                'Add to the Total Stock Volume
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                
                'Save Opening price for each sytock
                Opening_Price = ws.Cells(i - 1, 3).Value
                'Save Closing price for each stock
                Closing_Price = ws.Cells(i, 6).Value
                
               
                
                'Start Printing Stock ticket on the Summary table
                ws.Range("I" & Summary_Table_Row).Value = Stock_Ticket
                'Start printing the total volume for each stock ticket
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                'Start Printing yearly change = Clocing_Price - Opening_Price
                ws.Range("J" & Summary_Table_Row).Value = (Closing_Price) - (Opening_Price)
                'Start Printing yearly change = Clocing_Price - Opening_Price
                'ws.Range("K" & Summary_Table_Row).Value = (((Closing_Price) - (Opening_Price)) / Opening_Price)
                
            
                
                'Go down one Row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset The stock Volume for next Stock Ticket
                Stock_Volume = 0
                Opening_Price = 0
                Closing_Price = 0
                
                'Formatting
                
                    'If ws.Cells(i, 10).Value > 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                    'Else
                        'ws.Cells(i, 10).Interior.ColorIndex = 3
                    'End If
                    
                    'Percent for all Changes
                   ' ws.Cells(i, 11).Style = Percent
                
                
                
                
            Else
            
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                
            End If
              
        Next i
        
        'Add Titles to the table
        ws.Cells(1, 9).Value = "Ticket"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Volume"
    
    Next ws
            
  
End Sub
