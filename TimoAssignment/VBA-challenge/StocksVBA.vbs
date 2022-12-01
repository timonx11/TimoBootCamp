Attribute VB_Name = "Module1"
'Please Have a look at README.md for all the instructions / questions
'-------------------------------------------------------------------------------------------------------------------------------------------
Sub Stocks()
    'declare and create variable for worksheet or to hold worksheet name
    Dim ws As Worksheet

    'looping through all worksheets
    For Each ws In Worksheets
        
        'Declaring Variable needed for the main code to work
        '----------------------------------------------------------------------------------------------------------------------------------
        'declaring variable column for column pointer purposes
        Dim column As Integer
        'set/assigned value of column = 1
        column = 1
        
        'declaring variable tickercounter for cells and rows comparison
        Dim tickercounter As Integer
        'set/assigned value of tickercounter = 1
        tickercounter = 2
        
        'declaring variable to hold firstopenprice value
        Dim firstopenprice As Double
        
        'declaring variable to hold firstopenprice value
        Dim stockstotal As Double
        'set value of stockstotal = 0
        stockstotal = 0
        
        'declare and set variable lastrow to get the value of the last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'declare boolean to hold first opening price value
        Dim holdfirstvalue As Boolean
        
        'Set header to specific cells to provide results/summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'MAIN FOR LOOP TO PROCESS THE STOCKS DATA
        For i = 2 To lastrow
            ' if function to hold/reserve first value of stocks
            If holdfirstvalue = False Then
                 
                 'Set opening price value
                 firstopenprice = ws.Cells(i, 3).Value
        
                 'ensures no future prices captured until condition met.
                 holdfirstvalue = True
            End If
            
            ' if function to compare current/selected cells with the cells below it
            ' if current cells value is not the same with the value of cells below it
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                'add value to stockstotal
                stockstotal = stockstotal + ws.Cells(i, 7).Value
                
                'print ticker in the result/summary table
                ws.Cells(tickercounter, 9).Value = ws.Cells(i, column).Value
                
                'print
                ws.Cells(tickercounter, 10).Value = ws.Cells(i, 6).Value - firstopenprice
                
                If firstopenprice <> 0 Then
                ws.Cells(tickercounter, 11).Value = ((ws.Cells(i, 6).Value - firstopenprice) / firstopenprice)
                Else
                End If
                
                'print volumetotal in the result/summary table
                ws.Cells(tickercounter, 12).Value = stockstotal
                
                'proceed to the next cells/rows
                tickercounter = tickercounter + 1
                
                'reset stocktotal for the next ticker
                stockstotal = 0
                
                'reset the first value
                holdfirstvalue = False
                
            Else
                'add value to stockstotal
                stockstotal = stockstotal + ws.Cells(i, 7).Value
                
            End If
                       
        Next i
        
        'format column percent change in results/summary table
        'looping through all rows
        For i = 2 To lastrow
                'add % format to cells
                ws.Cells(i, 11).NumberFormat = "0.00%"

        Next i
        
        'format yearly change and percent change colum in results/summary table to show red when had negative value and green if positive
        'looping through all rows
        For i = 2 To lastrow
            For j = 10 To 11
                If ws.Cells(i, j).Value < 0 Then
                    ws.Cells(i, j).Interior.ColorIndex = 3
                ElseIf ws.Cells(i, j).Value > 0 Then
                    ws.Cells(i, j).Interior.ColorIndex = 4
                End If
            Next j
        Next i
        
        'VARIABLE FOR BONUS QUESTIONS/TASK
        'Set header to specific cells to provide results/summary table for BONUS QUESTIONS
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
             
        'set myrange variable as range
        Dim myrange As Range
        'set column K to myrange
        Set myrange = ws.Range("K:K")
        'set myrange2 variable as range
        Dim myrange2 As Range
        'set column L to myrange2
        Set myrange2 = ws.Range("L:L")
        
        'set variable lowestpercentdecrease to hold the lowestpercentdecrease value
        Dim lowestpercentdecrease As Double
        'set variable higherpercentincrease to hold the higherpercentincrease value
        Dim highestpercentincrease As Double
        'set variable highesttotalvolume to hold the highesttotalvolume value
        Dim highesttotalvolume As Double
        
        'functions to find highest and lowest value on column K and L
         lowestpercentdecrease = Application.WorksheetFunction.Min(myrange)
         highestpercentincrease = Application.WorksheetFunction.Max(myrange)
         highesttotalvolume = Application.WorksheetFunction.Max(myrange2)
        
        'FOOR LOOP FOR BONUS QUESTIONS TO show stocks that had lowest % decrease, highest % increase, and highest volume stocks
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value = lowestpercentdecrease Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9)
                    ws.Cells(3, 17).Value = lowestpercentdecrease
                    ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(i, 11).Value = highestpercentincrease Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9)
                    ws.Cells(2, 17).Value = highestpercentincrease
                    ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(i, 12).Value = highesttotalvolume Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 9)
                    ws.Cells(4, 17).Value = ws.Cells(i, 12)
            End If
        Next i
        
    Next ws
    
    MsgBox ("Data Compiled")

End Sub
