Sub stockSummary()
for each ws in worksheets

  Dim tickerName As String
  Dim openPrice As Double
  Dim closePrice As Double
  Dim tickerVolume As Variant

  'Create new summary column headers'

  ws.Cells(1, "I").Value = "Ticker"
  ws.Cells(1, "J").Value = "Yearly Change"
  ws.Cells(1, "K").Value = "Percent Change"
  ws.Cells(1, "L").Value = "Total Stock Volume"

  'Find Last Rows of Data Set'
  lastRowData = ws.Cells(Rows.Count, 1).End(xlUp).Row

  'Get opening price for first day of first stock on spreadsheet'
  openPrice = ws.Cells(2,3).Value

  'Loop through data'
  For i = 2 To lastRowData

  lastRowTicker = ws.Cells(Rows.Count, "I").End(xlUp).Row + 1

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
      
      'set ticker name'
      tickerName = ws.Cells(i, 1).Value
      
      'add stock volume to total'
      tickerVolume = tickerVolume + ws.Cells(i, "G").Value

      'set closing price'
      closePrice = ws.Cells(i,6).Value
      
      'Create ticker summary list'
      ws.Cells(lastRowTicker, "I").Value = tickerName
      ws.Cells(lastRowTicker, "J").Value = closePrice - openPrice
      If openPrice <> 0 Then
        ws.Cells(lastRowTicker, "K").Value = (closePrice - openPrice)/openPrice
        Else ws.Cells(lastRowTicker, "K").Value = 0
        End if
      ws.Cells(lastRowTicker, "L").Value = tickerVolume
      
      'Reset ticker volume'
      tickerVolume = 0

      'Reset the opening price for the next stock before moving on'
      openPrice = ws.Cells(i + 1,3).Value
      
    Else
      tickerVolume = tickerVolume + ws.Cells(i, "G").Value
    
    End If
  Next i

  'Formatting the summary table'
  finalTickerRow = ws.Cells(Rows.Count, "I").End(xlUp).Row

  for i = 2 to finalTickerRow
    If ws.Cells(i,"J").Value >= 0 Then
      ws.Cells(i,"J").Interior.ColorIndex = 4

    Else ws.Cells(i,"J").Interior.ColorIndex = 3

    End if

    ws.Cells(i,"K").NumberFormat = "0.00%"
  next i
next ws  
End Sub
