Sub StockVBA()

Dim WS_Count as Integer
Dim j as Integer
Dim lrow as Long
Dim Ticker as String
Dim Counter as Integer
Dim Total_Volume as Double 
Dim Summary_Table_Row as Integer
Dim OpenPrice as Double
Dim ClosingPrice as Double
Dim PercentChange as Double
Dim YearlyChange as Double


lrow = cells(Rows.Count, 1).End(xlUp).Row 
Counter = 0
Total_Volume = 0
Summary_Table_Row = 2

range("I1") = "Ticker"
'range("J1") = "Yearly Change"
'range("K1") = "Percent Change"
range("J1") = "Total Stock Volume"

for i = 2 to lrow
    
    if cells(i + 1,1).Value <> cells(i,1).Value then
    Ticker = cells(i,1).Value 
    Total_Volume = Total_Volume + cells(i,7).Value
    Counter = Counter + 1
    'ClosingPrice = cells(i,6).Value
    'YearlyChange = ClosingPrice - OpenPrice
    'PercentChange = (ClosingPrice# - OpenPrice#) / OpenPrice#


        
        
    'Print the Ticker Symbol to the Ticker Column
    Range("I" & Summary_Table_Row).Value = Ticker
    
    'Print the Price Change for the Stock over the year
    'Range("J" & Summary_Table_Row).Value = YearlyChange

    'Print the Percent Change for the Stock over the year
    'Range("K" & Summary_Table_Row).Value = PercentChange

    'Print the Stock's Total Value in the Total Value Column
    Range("J" & Summary_Table_Row).Value = Total_Volume

    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1

    'Reset the Total Volume
    Total_Volume = 0
    Counter = 0

    else
        Total_Volume = Total_Volume + cells(i,7).Value
        Counter = Counter + 1

        'if Counter = 1 then
        'OpenPrice = cells(i,3).Value
        'End if 
    
    
        
    end if

    
        
next i 


End Sub
