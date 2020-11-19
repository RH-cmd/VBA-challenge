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

'Adding Names to Columns/setting variables
 ws.Range("I1").Value = "Ticker"
 ticker_count=0
 ws.Range("J1").Value = "Yearly Change"
 yearly_change=0
 ws.Range("K1").Value = "Percent Change"
 yearly_change=0
 ws.Range("L1").Value = "Total Stock Volume"
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

    'Once we get to a different ticket
    If Cells(i + 1, 1).value <> ticker then
    ticker_count = ticker_count + 1
    Cells(ticker_count + 1, 9) = ticker

    'Closing price per ticker
    closing_price=cells(i,6).value

    'Yearly change per ticker
    yearly_change = (closing_price - opening_price)
    cells(ticker_count + 1, 10).value = yearly_change

    'Color coding yearly change value(green for positive, red for negative, )
    If yearly_change > 0 then
        cells(ticker_count + 1, 10).Interior.ColorIndex = 4
    Elseif yearly_change < 0 then
        cells(ticker_count + 1, 10).Interior.ColorIndex = 3
    Endif

    'Percent change for ticker
    If opening_price = 0 then
        percent_change = 0
    Else percent_change = (yearly_change / opening_price)
    Endif

    'Format percent change from general to percent 
    Cells(ticker_count + 1, 11).value = format(percent_change, "Percent")

    'Resetting opening price when changing to a different ticker
    opening_price = 0

    'Setting column L to total stock volume
    cells(ticker_count + 1, 12).value = stock_volume

    'Resetting total stock volume when changing to a different ticker
    stock_volume = 0
Endif

Next i

'Bonus - Greatest increase, decrease, etc, 
Dim greatest_percent_inc as double
Dim greatest_percent_inc_ticker as string
Dim greatest_percent_dec as double 
Dim greatest_percent_dec_ticker as string
Dim greatest_total_volume as double 
Dim greatest_volume_ticker as string

Range("O2").value = "Greatest % Percent"
Range("O3").value = "Greatest % Percent"
Range("O4").value = "Greatest Total Volume"
Range("P1").value = "Ticker"
Range("Q1").value = "Value"

lastRowState = ws.cells(rows.count, "I").End(xlUp).ROw

'Intialize and set values of variables
greatest_percent_inc = cells(2,11).value
greatest_percent_inc_ticker = cells(2,9).value
greatest_percent_dec = cells(2,11).value
greatest_percent_dec_ticker = cells(2,9).value
greatest_total_volume = cells(2,12).value
greatest_volume_ticker = cells(2,9).value

'Loop through the list of tickers
For i = 2 to lastRowState

'Ticker with greatest percent increase
    If cells(i,11).value > greatest_percent_inc then
        greatest_percent_inc = cells(i,11).value
        greatest_percent_inc_ticker = cells(i,9).value
    Endif
    
End Sub
