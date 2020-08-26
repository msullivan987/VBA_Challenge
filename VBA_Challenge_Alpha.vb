Sub stockSummary()

Dim ticker As String
Dim day As Integer
Dim openPrice As Double
Dim highPrice As Double
Dim lowPrice As Double
Dim closePrice As Double
Dim volume As Integer

'Create new summary column headers'
lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column

Cells(1, lastColumn + 2).Value = "Ticker"
Cells(1, lastColumn + 3).Value = "Yearly Change"
Cells(1, lastColumn + 4).Value = "Percent Change"
Cells(1, lastColumn + 5).Value = "Total Stock Volume"

'Find Last Rows of Data Set and Summary Table'
lastRowData = Cells(Rows.Count, 1).End(xlUp).Row

'Creating Ticker Summary List'
For i = 2 To lastRowData

lastRowTicker = Cells(Rows.Count, lastColumn + 2).End(xlUp).Row + 1

  If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Cells(lastRowTicker, lastColumn + 2).Value = Cells(i, 1).Value
    
  End If
Next i


End Sub
