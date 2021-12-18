Sub StockSummary():

'Set Primary Variables
Dim Stock_Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Volume As Double

'Set Secondary Variables
Dim Beginning_Price As Double
Dim Closing_Price As Double
Dim Opening_Price As Double

'Set Bonus Variables
Dim Range_Bonus As Range
Dim Max_Stock, Min_Stock, Max_Volume_Stock As String
Dim Max_Value, Min_Value, Max_Volume_Value As Double

For Each ws In Worksheets

'label summary table headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'label bonus table headers
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Define Variables
Yearly_Change = Closing_Price - Beginning_Price

'Initial total volume is 0
Total_Volume = 0

'Initial Row to be input at Summary Table (Row 2)
Summary_Table_Row = 2

'Determine the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'First Beginning Price
Beginning_Price = ws.Cells(2, 3).Value

'loop conditions
For i = 2 To LastRow

    'If Stock info changes...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Stock_Ticker = ws.Cells(i, 1).Value
        Closing_Price = ws.Cells(i, 6).Value
        Yearly_Change = Closing_Price - Beginning_Price
        
        'Opening price 0 condition
        If Beginning_Price <> 0 Then
            Percent_Change = (Yearly_Change / Beginning_Price) * 100
        End If
        
        'Add Total Volume for previous stock
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
        'Print previous stock info into summary table
        ws.Range("I" & Summary_Table_Row).Value = Stock_Ticker
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change & "%"
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
        
        'Conditional Formatting (Green/Red)
        If Yearly_Change > 0 Then
            'Green
            ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
        Else
            'Red
            ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
        End If
        
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        Yearly_Change = 0
        
        Closing_Price = 0
        
        Beginning_Price = ws.Cells(i + 1, 3).Value
        
        Total_Volume = 0
        
    Else
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
    End If
        
Next i

'Bonus Calculation

'Determine the last row
LastRow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To LastRow_Summary_Table
    
    'Greatest % Increase
    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Columns("K")) Then
        ws.Range("O2") = ws.Cells(i, 9).Value
        ws.Range("P2") = ws.Cells(i, 11).Value
        ws.Range("P2").NumberFormat = "0.00%"
        
    'Greatest % Decrease
    ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Columns("K")) Then
        ws.Range("O3") = ws.Cells(i, 9).Value
        ws.Range("P3") = ws.Cells(i, 11).Value
        ws.Range("P3").NumberFormat = "0.00%"
    
    'Greatest Total Volume
    ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Columns("L")) Then
        ws.Range("O4") = ws.Cells(i, 9).Value
        ws.Range("P4") = ws.Cells(i, 12).Value
    
    End If
    
Next i

Next ws

End Sub
