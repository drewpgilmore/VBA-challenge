Sub StockSummary()
    Dim WorksheetCount as Integer 'set worksheet variable for worksheet loop
        WorksheetCount = ActiveWorkbook.Worksheets.Count
    Dim x as Integer 'integer for looping through worksheets
        For x = 1 to WorksheetCount 'loop through worksheets           
                Dim lastRow as Double, summaryRow as Integer 'set lastRow to equal number of rows & set starting summary row            
                lastRow = Worksheets(x).Cells(Rows.Count, 1).End(xlUp).Row 'define lastRow               
                summaryRow = 2 'define starting row for summary calcs             
                Dim tickerSymbol as String 'set summary calculation variables
                Dim yearlyChange As Double 
                Dim percentChange As Double 
                Dim totalVolume As Double            
                totalVolume = 0 'define starting point for total volume running total            
                Worksheets(x).Range("I1") = "Ticker" 'print summary headers
                Worksheets(x).Range("J1") = "Yearly Change"
                Worksheets(x).Range("K1") = "Percent Change"
                Worksheets(x).Range("L1") = "Total Stock Volume"
            For i = 2 to lastRow 'set up loop through rows                
                IF ( Worksheets(x).Cells(i+1,1).Value <> Worksheets(x).Cells(i,1).Value ) Then 'detect unique values if consecutive tickers are not equal
                    tickerSymbol = Worksheets(x).Cells(i,1).Value 'set tickerSymbol
                    Worksheets(x).Range("I" & summaryRow).Value = tickerSymbol 'print tickerSymbol to column I
                    totalVolume = totalVolume + Worksheets(x).Cells(i,7).Value 'set totalVolume
                    Worksheets(x).Range("L" & summaryRow).Value = totalVolume 'print totalVolume to column L
                    summaryRow = summaryRow + 1 'add 1 to summary row for next ticker                        
                    totalVolume = 0 'reset volume running total
                    Else totalVolume = totalVolume + Worksheets(x).Cells(i,7).Value 'keep running total of volume until unique ticker is found
                End If
            Next i
        Next x    
End Sub