Sub Challenge()
Range("I:L").Clear

'define variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Opening_Price As Double
Dim Ending_Price As Double
Dim Percent_Change As Double
Dim Summary_Table_Row As Integer
Dim Total_Stock_Volume As Double
Dim lastrow As Double
Dim ws As Worksheet

'Ticker = 0
Yearly_Change = 0
'Opening_Price = 0
'Ending_Price = 0
'Percent_Change = 0
Summary_Table_Row = 2
'define last row
lastrow = Cells(Rows.Count, 1).End(xlUp).row

'populate all worksheets
For Each ws In ThisWorkbook.Sheets
    ws.Select

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percent_Change"
Cells(1, 12).Value = "Total_Stock_Volume"

' Loop through all Tickers
For I = 2 To lastrow

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Ticker, Display
      Ticker = Cells(I, 1).Value
      Range("I" & Summary_Table_Row).Value = Ticker
      Summary_Table_Row = Summary_Table_Row + 1
      'Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(I, 7).Value
      
     
    
    'Else

End If
Next I
Summary_Table_Row = 2
Next ws


End Sub
