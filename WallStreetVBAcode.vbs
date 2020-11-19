Sub WallStreetChallenge

'Delcaring Variables
Dim ticker as string
Dim ticker_count as integer
Dim opening_price as double 
Dim closing_price as double 
Dim yearly_change as double 
Dim percent_change as double 
Dim stock_volume as double

'Loop over each worksheet 
For each ws in Worksheets
ws.Activate

'Last row of each worksheet
lastRowState=ws.cells(rows.count,"A").End(xlUp).row

'Adding Names to Columns
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"

'Setting variables per worksheet
ticker_count=0
yearly_change=0
opening_price=0
percent_change=0
stock_volume=0

'Loop through ticker 
For i = 2 to lastRowState
  
    ticker=Cells(i,1).Value

 'Opening price per ticker
    If opening_price=0 then
        opening_price=cells(i,3).Value
    Endif

    'Total stock volumes for each ticker
    stock_volume=stock_volume + cells(i,7).value
    
End Sub
