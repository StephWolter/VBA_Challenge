Sub Challenge()
'Clear formatting and data
Range("I:L").Clear
Range("J").Interior.ColorIndex = 0

'define variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Opening_Price As Double
Dim Ending_Price As Double
Dim Percent_Change As Variant
Dim Summary_Table_Row As Integer
Dim Total_Stock_Volume As Double
Dim lastrow As Double
Dim ws As Worksheet

'populate all worksheets
For Each ws In ThisWorkbook.Sheets
    ws.Select
    
'set initial values
Ticker = 0
Yearly_Change = 0
Opening_Price = 0
Ending_Price = 0
Percent_Change = 0
Summary_Table_Row = 2
Total_Stock_Volume = 0


'define last row
lastrow = Cells(Rows.Count, 1).End(xlUp).row

    
'fill out header
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percent_Change"
Cells(1, 12).Value = "Total_Stock_Volume"

' Loop through all Tickers
For I = 2 To lastrow

'set the Opening for yearly change calculation
If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    Opening_Price = Cells(I, 3).Value
End If
    

' Check if we are still within the same ticker symbol, if it is not...
If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

    'Calculate Total by first accounting for all amounts per ticker symbol
    Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
    Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    'reset total stock volume
    Total_Stock_Volume = 0

    ' Set the Ticker, Display
    Ticker = Cells(I, 1).Value
    Range("I" & Summary_Table_Row).Value = Ticker
    'Set the last closing price
    Closing_Price = Cells(I, 6).Value
      
    'Calculate yearly change
    Yearly_Change = Closing_Price - Opening_Price
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    
    'Calculate Percent Change column
    Percent_Change = FormatPercent((Closing_Price / Opening_Price) - 1)
    Range("K" & Summary_Table_Row).Value = Percent_Change
   

    'add color conditional for yearly change - red for <0 and green for >0
    If Yearly_Change < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    Else
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    End If
    
    'Move to next row
    Summary_Table_Row = Summary_Table_Row + 1

Else
' Add to the Total Stock Volume
Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value



End If
Next I


Next ws


End Sub

