Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock.

NOTE
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

The result should match the following image:

Moderate solution
Bonus
Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

Hard solution
Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

Other Considerations
Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

Some assignments, like this one, contain a bonus. It is possible to achieve proficiency for this assignment without completing the bonus, but the bonus is an opportunity to further develop your skills and receive extra points for doing so.

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
VBA Scripts for multiple sheets
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
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
VBA scripts for single sheet - not properly commented, for testing purposes
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Stocks()
    
    Dim ws As Worksheet

    For Each ws In Worksheets

        Dim column As Integer
        column = 1
        Dim tickercounter As Integer
        tickercounter = 2
        Dim firstopenprice As Double
        Dim stockstotal As Double
        stockstotal = 0
        'Dim stocksname As String 
        'total
        'Dim Brand_Total As Double
        'Brand_Total = Brand_Total + Cells(i, 3).Value
        'in if reset total
        
        'declare last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        

        
        'declare boolean to hold first opening price value
        Dim holdfirstvalue As Boolean
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To lastrow
            'hold/reserve first value of stocks
            If holdfirstvalue = False Then
                 'Set opening price
                 
                 firstopenprice = Cells(i, 3).Value
        
                 'ensures no future prices captured until condition met.
                 holdfirstvalue = True
            End If
            
            
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                'MsgBox (("next cells value will be ") & Cells(i, column).Value & " and then " & Cells(i + 1, column).Value)
                stockstotal = stockstotal + Cells(i, 7).Value
                
                'print ticker in the result/summary table
                Cells(tickercounter, 9).Value = Cells(i, column).Value
                
                Cells(tickercounter, 10).Value = Cells(i, 6).Value - firstopenprice
                
                If firstopenprice <> 0 Then
                Cells(tickercounter, 11).Value = ((Cells(i, 6).Value - firstopenprice) / firstopenprice)
                Else
                End If
                
                'print volumetotal in the result/summary table
                Cells(tickercounter, 12).Value = stockstotal
                
                'keep track of the rowcounter
                tickercounter = tickercounter + 1
                
                'reset stock total
                stockstotal = 0
                
                'reset the first opening price value
                holdfirstvalue = False
                
                'MsgBox ("last value of " + Cells(i, column).Value)
                'MsgBox (("next cells value will be ") & Cells(i + 1, column).Value)
                
            Else
                
                stockstotal = stockstotal + Cells(i, 7).Value
                
            End If
            
                
        Next i
        
        'format column percent change
        For i = 2 To lastrow

                Cells(i, 11).NumberFormat = "0.00%"

        Next i
        
        
        
        For i = 2 To lastrow
            For j = 10 To 11
                If Cells(i, j).Value < 0 Then
                    Cells(i, j).Interior.ColorIndex = 3
                ElseIf Cells(i, j).Value > 0 Then
                    Cells(i, j).Interior.ColorIndex = 4
                End If
            Next j
        Next i
        
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
		
		Dim myrange As Range
        Set myrange = Range("K:K")
        Set myrange2 = Range("L:L")
        Dim lowestpercentdecrease As Double
        Dim highestpercentincrease As Double
        Dim highesttotalvolume As Double
		
		lowestpercentdecrease = Application.WorksheetFunction.Min(myrange)
        highestpercentincrease = Application.WorksheetFunction.Max(myrange)
        highesttotalvolume = Application.WorksheetFunction.Max(myrange2)
        
        For i = 2 To lastrow
        
            If Cells(i, 11).Value = lowestpercentdecrease Then
                    Cells(3, 16).Value = Cells(i, 9)
                    Cells(3, 17).Value = lowestpercentdecrease
                    Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(i, 11).Value = highestpercentincrease Then
                    Cells(2, 16).Value = Cells(i, 9)
                    Cells(2, 17).Value = highestpercentincrease
                    Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(i, 12).Value = highesttotalvolume Then
                    Cells(4, 16).Value = Cells(i, 9)
                    Cells(4, 17).Value = Cells(i, 12)
            End If
        Next i
        'MsgBox Name
        'MsgBox (lowestpercentdecrease)
        'MsgBox (highestpercentincrease)
    Next ws
    
    MsgBox ("compile complete")

End Sub



