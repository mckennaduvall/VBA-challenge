Sub stockloop()

Dim ticker_symbol As String

Dim total_volume As Double
total_volume = 0

Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim summary_table_row As Integer

'summary table starts at row 2
summary_table_row = 2

'summary table
Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

'first opening price for ticker A
year_open = Cells(2, 3).Value

last_row = Cells(Rows.Count, 1).End(xlUp).Row


'iterate through ticker values
For i = 2 To last_row

    ticker_symbol = Cells(i, 1).Value
    next_ticker_symbol = Cells(i + 1, 1).Value


'conditional
If ticker_symbol <> next_ticker_symbol Then
    
    ticker_symbol = Cells(i, 1)
    total_volume = total_volume + Cells(i, 7)
    
    year_close = Cells(i, 6).Value

    year_change = Round(year_close - year_open, 2)
    
    percent_change = Round((year_change) / year_open, 4) * 100
   
   
    

    'output
    Range("I" & summary_table_row).Value = ticker_symbol
    Range("L" & summary_table_row).Value = total_volume
    Range("J" & summary_table_row).Value = year_change
    Range("K" & summary_table_row).Value = percent_change
    

    
    'reset for next ticker symbol
    summary_table_row = summary_table_row + 1
    
    total_volume = 0
    
        
    year_open = Cells(i + 1, 3).Value
        
Else
    
    'adding volume when symbols are equal
    total_volume = total_volume + Cells(i, 7)
    
End If


If Cells(i, 10).Value < 0 Then

Cells(i, 10).Interior.ColorIndex = 3

ElseIf Cells(i, 10).Value > 0 Then

Cells(i, 10).Interior.ColorIndex = 4

Else


End If


Next i


End Sub


