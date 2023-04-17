Sub VBA_Summary_Table()

'define variables
Dim SummaryTable As Integer
Dim StockVol As LongLong
Dim TickType As String
Dim OpenPrice As Double
Dim ClosePrice As Double

'to cycle through each worksheet
For Each ws In Worksheets
    
    'once in the worksheet assign these values to the varibles
    SummaryTable = 2 'this helps start the filling of the table below the header line
    StockVol = 0 'this resets the stock volume when changing worksheets

    'populate the header row for the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'define the last row of data so that it will run the loop as per each worksheetsdata size
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'starting at row 2 run through each row of the given dataset
    For rowi = 2 To lastRow
    
    'assigning the value of ticker symbol to the row we are in, column 1
    TickType = ws.Cells(rowi, 1).Value
        
        'once cycling through each row of the data do one of the following IFS
        
        'starting with if the new row is a different ticker symbol to the previous one
        If ws.Cells(rowi, 1).Value <> ws.Cells(rowi - 1, 1).Value Then 'this row ticker symbol is different to the one in the row above
            ws.Range("N1").Value = ws.Cells(rowi, 3).Value 'put this value aside to use later as the opening price
            StockVol = StockVol + ws.Cells(rowi, 7).Value 'start calculating the accumulative stock count
            
        'to continue growing the stock count we need to grab values for when the first IF statement is not true, but the ticker is the same as the one below it
        ElseIf ws.Cells(rowi, 1).Value = ws.Cells(rowi + 1, 1).Value Then 'because the above statement was past, this grabs rows where the ticker is the same above and below
            StockVol = StockVol + ws.Cells(rowi, 7).Value 'continue with the accumlation of stock count
            
        'grabbing all other options. As the above statements have been past over this works with the row that is the same as the one above it
        'but is different to the one below it, eg the last row for that ticker symbol
        Else: ws.Cells(SummaryTable, 9).Value = TickType 'assign this rows ticker symbol to the summary table
            StockVol = StockVol + ws.Cells(rowi, 7).Value 'add this rows stock to the stock count
            ws.Cells(SummaryTable, 12).Value = StockVol 'populate the summary table with the final stock count
            ClosePrice = ws.Cells(rowi, 6).Value 'assign this rows price as the closing price
            StockVol = 0 'reset the stock count back to zero ready for the next ticker symbol's accumulation
            ws.Cells(SummaryTable, 10).Value = ClosePrice - ws.Range("N1") 'using the closing price we assigned a few lines above and the opening price we set aside from the first row of the ticker symbol
            ws.Cells(SummaryTable, 11).Value = (ClosePrice - ws.Range("N1")) / (ws.Range("N1").Value) 'using the assigned values calulate the percentage change between opening and closing prices for the year
            ws.Cells(SummaryTable, 11).NumberFormat = "0.00%" 'format the percentages
                
                'set conditional formatting to highlight increases and decreases in price
                If ws.Cells(SummaryTable, 10).Value < 0 Then 'if the end of year price is less than the start of year price
                    ws.Cells(SummaryTable, 10).Interior.ColorIndex = 3 'format the cell to be red
                    
                'if the above is not true then the change in price will be zero or above
                Else: ws.Cells(SummaryTable, 10).Interior.ColorIndex = 4 'format the cell to be green
                
                End If
                
                'repeat conditional formatting to highlight increases and decreases in price percentage
                If ws.Cells(SummaryTable, 11).Value < 0 Then 'if the price change percentage has decreased
                    ws.Cells(SummaryTable, 11).Interior.ColorIndex = 3 'format the cell to be red
                    
                'if the above is not true then the change in price percentage will be zero or above
                Else: ws.Cells(SummaryTable, 11).Interior.ColorIndex = 4 'format the cell to be green
                
                End If
               
             'still in the ELSE statement for when we are at the last row of the ticker
             SummaryTable = SummaryTable + 1 'add 1 to the table count so that the next ticker symbol populates in the next row of the summary
             ws.Range("N1").ClearContents 'clear the held opening value ready for the next ticker
             'ws.Range("I:L").EntireColumn.AutoFit
             
           End If
           
    Next rowi
       
    'with the table updated format columns to show all values clearly
    ws.Range("I:L").EntireColumn.AutoFit
   
Next ws

End Sub

