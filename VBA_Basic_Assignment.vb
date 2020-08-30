Sub stockSummary()


  Dim tickerName As String
  Dim openPrice As Double
  Dim closePrice As Double
  Dim tickerVolume As Variant
  Dim maxPercentInc As Double
  Dim maxPercentDec As Double
  Dim maxVolume As Variant
  Dim maxPercentIncTicker As String
  Dim maxPercentDecTicker As String
  Dim maxVolumeTicker As String

  'Create new summary column headers'

  Cells(1, "I").Value = "Ticker"
  Cells(1, "J").Value = "Yearly Change"
  Cells(1, "K").Value = "Percent Change"
  Cells(1, "L").Value = "Total Stock Volume"

  'Find Last Rows of Data Set'
  lastRowData = Cells(Rows.Count, 1).End(xlUp).Row

  'Get opening price for first day of first stock on spreadsheet'
  openPrice = Cells(2,3).Value

  'Loop through data'
  For i = 2 To lastRowData

  lastRowTicker = Cells(Rows.Count, "I").End(xlUp).Row + 1

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
      
      'set ticker name'
      tickerName = Cells(i, 1).Value
      
      'add stock volume to total'
      tickerVolume = tickerVolume + Cells(i, "G").Value

      'set closing price'
      closePrice = Cells(i,6).Value
      
      'Create ticker summary list'
      Cells(lastRowTicker, "I").Value = tickerName
      Cells(lastRowTicker, "J").Value = closePrice - openPrice
      If openPrice <> 0 Then
        Cells(lastRowTicker, "K").Value = (closePrice - openPrice)/openPrice
        Else Cells(lastRowTicker, "K").Value = 0
        End if
      Cells(lastRowTicker, "L").Value = tickerVolume
      
      'Reset ticker volume'
      tickerVolume = 0

      'Reset the opening price for the next stock before moving on'
      openPrice = Cells(i + 1,3).Value
      
    Else
      tickerVolume = tickerVolume + Cells(i, "G").Value
    
    End If
  Next i

  'Formatting the summary table'
  finalTickerRow = Cells(Rows.Count, "I").End(xlUp).Row

  for i = 2 to finalTickerRow
    If Cells(i,"J").Value >= 0 Then
      Cells(i,"J").Interior.ColorIndex = 4

    Else Cells(i,"J").Interior.ColorIndex = 3

    End if

    Cells(i,"K").NumberFormat = "0.00%"
  next i
 
End Sub
