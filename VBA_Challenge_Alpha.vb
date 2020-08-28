Sub stockSummary()

Dim tickerName As String
Dim openPrice As Double
Dim closePrice As Double
Dim tickerVolume As Variant

'Create new summary column headers'
lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

Cells(1, lastColumn + 2).Value = "Ticker"
Cells(1, lastColumn + 3).Value = "Yearly Change"
Cells(1, lastColumn + 4).Value = "Percent Change"
Cells(1, lastColumn + 5).Value = "Total Stock Volume"

'Find Last Rows of Data Set'
lastRowData = Cells(Rows.Count, 1).End(xlUp).Row

'Get opening price for first day of first stock on spreadsheet'
openPrice = Cells(2,3).Value

'Loop through data'
For i = 2 To lastRowData

lastRowTicker = Cells(Rows.Count, lastColumn + 2).End(xlUp).Row + 1

  If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    'set ticker name'
    tickerName = Cells(i, 1).Value
    
    'add stock volume to total'
    tickerVolume = tickerVolume + Cells(i, "G").Value

    'set closing price'
    closePrice = Cells(i,6).Value
    
    'Create ticker summary list'
    Cells(lastRowTicker, lastColumn + 2).Value = tickerName
    Cells(lastRowTicker, lastColumn + 5).Value = tickerVolume
    Cells(lastRowTicker, lastColumn + 6).Value = closePrice - openPrice
    Cells(lastRowTicker, lastColumn + 7).Value = (closePrice - openPrice)/openPrice
    
    'Reset ticker volume'
    tickerVolume = 0

    'Reset the opening price for the next stock before moving on'
    openPrice = Cells(i + 1,3).Value
    
   Else
    tickerVolume = tickerVolume + Cells(i, "G").Value
   
  End If
Next i

End Sub
