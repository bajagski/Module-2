Attribute VB_Name = "Module2"
Sub test_Stock_2()

         ' Declare Current as a worksheet object variable.
         Dim Current As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets

            'Set a inital vairable for holding the stock name
                Dim Stock_Name As String
                
                'Set an inital variable for holding the total volume per stock name
                Dim Stock_Vol As Double
                Stock_Vol = 0
                
                'Keep track of stock name in summary table
                Dim Summary_Table_Row As Double
                Summary_Table_Row = 2
                
                'Find the last row on each worksheet
                lastrow = Cells(Rows.Count, 1).End(xlUp).Row
                
                 ' Loop through all Stocks on Sheet 1
                  For i = 2 To lastrow
                  
                      ' Check if we are still within the same stock name, if it is not...
                    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                      ' Set the Stock name
                      Stock_Name = Cells(i, 1).Value
                
                      ' Add to the Stock Total
                      Stock_Vol = Stock_Vol + Cells(i, 7).Value
                
                      ' Print the Stock Name in the Summary Table
                      Range("I" & Summary_Table_Row).Value = Stock_Name
                
                      ' Print the Stock Volume to the Summary Table
                      Range("J" & Summary_Table_Row).Value = Stock_Vol
                
                      ' Add one to the summary table row
                      Summary_Table_Row = Summary_Table_Row + 1
                      
                      ' Reset the Brand Total
                      Stock_Vol = 0
                    
                    ' If the cell immediately following a row is the same Stock Name...
                    Else
                
                      ' Add to the Stock Volumne
                      Stock_Vol = Stock_Vol + Cells(i, 7).Value
                
                    End If
                
                  Next i
                  
                  'print last row that was calculated
                    MsgBox lastrow
            
            
            ' This line displays the worksheet name in a message box.
            MsgBox Current.Name
         Next

      End Sub
