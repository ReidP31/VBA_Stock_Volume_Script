Sub StockVBA()

'Cycle through the workbook
Dim ws as Worksheet

For each ws in ThisWorkbook.Worksheets
    ws.Activate

    'Declare Variable Type
    Dim lrow as Long
    Dim Ticker as String
    Dim Counter as Integer
    Dim Total_Volume as Double 
    Dim Summary_Table_Row as Integer
    Dim OpenPrice as Double
    Dim ClosingPrice as Double
    Dim PercentChange as Variant
    Dim YearlyChange as Double

    'Assign Variable Values
    lrow = cells(Rows.Count, 1).End(xlUp).Row 
    Counter = 0
    Total_Volume = 0
    Summary_Table_Row = 2
    OpenPrice = cells(2,3).Value

    range("I1") = "Ticker"
    range("J1") = "Yearly Change"
    range("K1") = "Percent Change"
    range("L1") = "Total Stock Volume"

    'Loop through the Data and create the Summary Table
    for i = 2 to lrow
        
        if cells(i + 1,1).Value <> cells(i,1).Value then
            Ticker = cells(i, 1).Value 
            Total_Volume = Total_Volume + cells(i,7).Value
            'Counter = Counter + 1
            ClosingPrice = cells(i,6).Value
            YearlyChange = ClosingPrice - OpenPrice
            
            
            if OpenPrice = 0 then
                PercentChange = "#N/A"
            else
                PercentChange = (ClosingPrice - OpenPrice) / OpenPrice
            end if
                        
            
            'Print the Ticker Symbol to the Ticker Column
            Range("I" & Summary_Table_Row).Value = Ticker
            
            'Print the Price Change for the Stock over the year
            Range("J" & Summary_Table_Row).Value = YearlyChange

            'Print the Percent Change for the Stock over the year
            Range("K" & Summary_Table_Row).Value = PercentChange

            'Print the Stock's Total Value in the Total Value Column
            Range("L" & Summary_Table_Row).Value = Total_Volume

            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            

            
            'Reset the Total Volume
            Total_Volume = 0
            'Counter = 0

        else
            Total_Volume = Total_Volume + cells(i,7).Value
            'Counter = Counter + 1

            
        
        
            
        end if

        if YearlyChange > 0 and PercentChange > 0 then
            Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 4 
            Range("K" & Summary_Table_Row - 1).Interior.ColorIndex = 4 
        elseif YearlyChange < 0 and PercentChange < 0 then
            Range("J" & Summary_Table_Row - 1).Interior.ColorIndex = 3
            Range("K" & Summary_Table_Row - 1).Interior.ColorIndex = 3 
        end if

        OpenPrice = cells(i + 1, 3).Value

    next i 

    range("J2","J" & lrow).NumberFormat = "0.0000"
    range("K2","K" & lrow).NumberFormat = "0.0000%"


Next ws

End Sub
