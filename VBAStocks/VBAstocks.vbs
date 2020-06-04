Sub StockSummary()
    'set lastRow to equal number of rows & set starting summary row
        Dim lastRow as Double, summaryRow as Integer
            lastRow = Cells(Rows.Count, 1).End(xlUp).Row
            summaryRow = 2
    'set summary calculation variables
        Dim tickerSymbol as String 
        Dim yearlyChange As Double 
        Dim percentChange As Double 
        Dim totalVolume As Double
            totalVolume = 0
    'name output headers
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
    'set up loop through rows    
        For i = 2 to lastRow
            'detect unique values if consecutive tickers are not equal
            IF ( Cells(i+1,1).Value <> Cells(i,1).Value ) Then
                'set tickerSymbol and print to column I
                    tickerSymbol = Cells(i,1).Value
                    Range("I" & summaryRow).Value = tickerSymbol
                'set total volume and print to column L
                    totalVolume = totalVolume + Cells(i,7).Value
                    Range("L" & summaryRow).Value = totalVolume
                'add 1 to summary row for next ticker
                    summaryRow = summaryRow + 1
                'reset volume running total
                    totalVolume = 0
            Else
                'keep running total of volume until unique ticker is found
                    totalVolume = totalVolume + Cells(i,7).Value
            End If
        Next i    
End Sub

'Challenge
'1 - return ticker symbol with:
'Greatest % Increase
'Greatest % Decrease
'Greatest Total Volume

'2 add compatability to run on different sheets

