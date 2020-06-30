Attribute VB_Name = "Module1"
Sub stock_analysis():
    
    'To run on all sheets
    For Each ws In Worksheets

        'Write column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Autofit appropriate column spacing
        ws.Columns("I:L").AutoFit
    
        'Assigning ticker column
        
        Dim tickercol As Integer
        tickercol = 9
        
        Dim tickerrow As Integer
        tickerrow = 2
        
        'Find last row
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Assign Total Stock Volume Column
        Dim totalvolcol As Integer
        totalvolcol = 12
        
        'Initial Stock Volume for each stock
        Dim totalvolume As Double
        totalvolume = 0
        
        Dim i As Double
        
        Dim openingvalue As Double
        Dim closingvalue As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        
        
        'Assign Yearly Change Column
        yearlychangecol = 10
        
        'Assign Percent Change Column
        percentchangecol = 11
        
        'Find value of opening price for only the first stock
        openingvalue = ws.Cells(2, 3).Value
        
        ws.Cells(2, 10).Value = openingvalue
        
        
    
            For i = 2 To lastrow
    
                Dim currentcell As String
                currentcell = ws.Cells(i, 1).Value
    
                Dim nextcell As String
                nextcell = ws.Cells(i + 1, 1).Value
    
                'To identify new stock ticker
                If currentcell <> nextcell Then
        
                            'Write new ticker in ticker column
                            ws.Cells(tickerrow, tickercol).Value = currentcell
                            
                            'Add last stock volume value for current ticker
                            totalvolume = totalvolume + ws.Cells(i, 7).Value
                
                            'Write total volume to total volume column
                            ws.Cells(tickerrow, totalvolcol).Value = totalvolume
                
                            'Reads the value of closing stock price for current ticker
                            closingvalue = ws.Cells(i, 6).Value
                        
                            'Calculates the yearly change in stock price
                            yearlychange = closingvalue - openingvalue
                            
                            
                                        'To avoid dividing by zero
                                            If openingvalue <> 0 Then
                        
                                                percentchange = yearlychange / openingvalue
                        
                                            Else
                                                percentchange = 0
                        
                                            End If
                    
                    
                            'Writes yearly change to yearly change column
                            ws.Cells(tickerrow, yearlychangecol).Value = yearlychange
                            
                            'Writes the percentage change to percent change column
                            ws.Cells(tickerrow, percentchangecol).Value = percentchange
                    
                                        'Assign color formating
                                            If yearlychange < 0 Then
                                                ws.Cells(tickerrow, yearlychangecol).Interior.ColorIndex = 3
                            
                                            Else
                                                ws.Cells(tickerrow, yearlychangecol).Interior.ColorIndex = 4
                            
                                            End If
                            
                            'Write number as percentage
                            ws.Cells(tickerrow, percentchangecol).NumberFormat = "###,##0.00%"
                    
                            'So that we don't overwrite out tickers in the ticker column
                            tickerrow = tickerrow + 1
                    
                             'Find value of opening price for each new stock
                            openingvalue = ws.Cells(i + 1, 3).Value
                            
                            'Resets total stock volume for next ticker
                            totalvolume = 0
        
              Else
        
                            'Accumulate stock volume total for each stock
                            totalvolume = totalvolume + ws.Cells(i, 7).Value
  
        
        
        
            End If
    
    
            Next i
    
    
    Next ws

End Sub
