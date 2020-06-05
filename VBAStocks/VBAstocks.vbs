Sub StockSummary()
    Dim WorksheetCount as Integer 'set worksheet variable for worksheet loop
        WorksheetCount = ActiveWorkbook.Worksheets.Count
    Dim x as Integer 'integer for looping through worksheets
        For x = 1 to WorksheetCount 'loop through worksheets
                Dim lastRow as Double  'set lastRow to equal number of rows & set starting summary row            
                    lastRow = Worksheets(x).Cells(Rows.Count, 1).End(xlUp).Row 'define lastRow               
                
                Dim summaryRow as Double 'set variables for summary section
                    summaryRow = 2 'define starting row # for summary calcs
                Dim tickerSymbol as String
                Dim yearlyChange as Double 
                Dim percentChange as Double
                Dim totalVolume as Double                    
                    totalVolume = 0 'define starting point for total volume running total  
                
                Dim openPrice, closePrice, priceReference as Double 'set reference variables                                  
                    priceReference = 2 'define starting point as first row with data

                Worksheets(x).Range("I1") = "Ticker" 'print summary headers
                Worksheets(x).Range("J1") = "Yearly Change"
                Worksheets(x).Range("K1") = "Percent Change"
                Worksheets(x).Range("L1") = "Total Stock Volume"

                Worksheets(x).Range("O2") = "Greatest % Increase" 'print challenge summary labels and headers
                Worksheets(x).Range("O3") = "Greatest % Decrease"
                Worksheets(x).Range("O4") = "Greatest Total Volume"
                Worksheets(x).Range("P1") = "Ticker"
                Worksheets(x).Range("Q1") = "Value"

            For i = 2 to lastRow 'set up loop through rows
                    IF ( Worksheets(x).Cells(i+1,1).Value <> Worksheets(x).Cells(i,1).Value ) Then 'detect unique values if consecutive tickers are not equal
                    tickerSymbol = Worksheets(x).Cells(i,1).Value 'set tickerSymbol
                        Worksheets(x).Range("I" & summaryRow).Value = tickerSymbol 'print tickerSymbol to column I
                    openPrice = Worksheets(x).Cells(priceReference,3).Value 'set open
                    closePrice = Worksheets(x).Cells(i,6).Value 'set close
                    yearlyChange = closePrice - openPrice 'set yearlyChange
                        IF yearlyChange >= 0 Then 'determine positive or negative yearly change
                            Worksheets(x).Range("J" & summaryRow).Interior.ColorIndex = 4 'color cell green
                            Else Worksheets(x).Range("J" & summaryRow).Interior.ColorIndex = 3 'color cell red
                        End If                          
                        Worksheets(x).Range("J" & summaryRow).Value = yearlyChange 'print yearly change to column J
                        IF openPrice = 0 Then
                            percentChange = 0
                            Else percentChange = yearlyChange / openPrice 'set percent change to equal yearly change divided by open price
                        End If
                        Worksheets(x).Range("K" & summaryRow).Value = percentChange 'print percentage change to column K
                        Worksheets(x).Range("K" & summaryRow).Style = "Percent"
                    totalVolume = totalVolume + Worksheets(x).Cells(i,7).Value 'set totalVolume
                        Worksheets(x).Range("L" & summaryRow).Value = totalVolume 'print totalVolume to column L
                    summaryRow = summaryRow + 1 'add 1 to summary row for next ticker  
                    priceReference = i + 1 'set open price reference for next ticker              
                    totalVolume = 0 'reset volume running total
                    Else  
                        totalVolume = totalVolume + Worksheets(x).Cells(i,7).Value 'keep running total of volume until unique ticker is found
                End If
            Next i
        Dim greatestIncrease, greatestDecrease, greatestVolume as Double
        greatestIncrease = WorksheetFunction.Max(Worksheets(x).Range("K2:K" & summaryRow).Value)
        greatestDecrease = WorksheetFunction.Min(Worksheets(x).Range("K2:K" & summaryRow).Value)
        greatestVolume = WorksheetFunction.Max(Worksheets(x).Range("L2:L" & summaryRow).Value)

            Dim y as Double
                For y = 2 to summaryRow
                    IF ( Worksheets(x).Cells(y,11).Value = greatestIncrease ) Then
                        Worksheets(x).Range("P2").Value = Worksheets(x).Cells(y,9).Value
                        Worksheets(x).Range("Q2").Value = Worksheets(x).Cells(y,11).Value
                        Worksheets(x).Range("Q2").Style = "Percent"
                    ElseIF ( Worksheets(x).Cells(y,11).Value = greatestDecrease ) Then
                        Worksheets(x).Range("P3").Value = Worksheets(x).Cells(y,9).Value     
                        Worksheets(x).Range("Q3").Value = Worksheets(x).Cells(y,11).Value
                        Worksheets(x).Range("Q3").Style = "Percent"  
                    ElseIF ( Worksheets(x).Cells(y,12).Value = greatestVolume ) Then
                        Worksheets(x).Range("P4").Value = Worksheets(x).Cells(y,9).Value
                        Worksheets(x).Range("Q4").Value = Worksheets(x).Cells(y,12).Value
                End If 
            Next y
        Worksheets(x).Range("A:Q").Columns.AutoFit       
        Next x    
End Sub