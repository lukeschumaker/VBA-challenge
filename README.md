Sub stock_data()

Dim column As Integer
column = 1
Dim summary_table_row As LongLong
summary_table_row = 2
Dim total_stock_volume As LongLong
total_stock_volume = 0
Dim opening_price As Double
Dim Closing_price As Double
Dim Yearly_change As Double
Dim Percent_change As Double
Dim greatest_total_volume As LongLong
Dim greatest_total_volume_ticker As String
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String


Dim condition1 As FormatCondition, condition2 As FormatCondition
Dim rng As Range

Set rng = Range("K2", "K2000")

'Add headings to a summary table
Cells(1, 9).Value = "Summary Table"
Cells(1, 10).Value = "Ticker Symbols"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

greatest_total_volume = 0
greatest_percent_increase = 0
greatest_percent_decrease = 999999999

For i = 2 To lastrow


    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
       opening_price = Cells(i, 3).Value
    
    End If
    

    'Running total of stock volume
    total_stock_volume = total_stock_volume + (Cells(i, 7).Value)
    
    If total_stock_volume > greatest_total_volume Then
        greatest_total_volume = total_stock_volume
        greatest_total_volume_ticker = Cells(i, 1).Value
    End If
    
    If Cells(i, column).Value <> Cells(i + 1, column).Value Then
    
        Closing_price = Cells(i, 6).Value
        Yearly_change = Closing_price - opening_price
        
        Percent_change = (Yearly_change / opening_price)
         
         
    If Percent_change > greatest_percent_increase Then
        greatest_percent_increase = Percent_change
        greatest_percent_increase_ticker = Cells(i, 1).Value
    End If
    
    If Percent_change < greatest_percent_decrease Then
        greatest_percent_decrease = Percent_change
        greatest_percent_decrease_ticker = Cells(i, 1).Value
    End If
    
        Cells(summary_table_row, 10).Value = Cells(i, 1).Value
        
        Cells(summary_table_row, 11).Value = Yearly_change
        
        Cells(summary_table_row, 12).Value = Percent_change
     
        Cells(summary_table_row, 13).Value = total_stock_volume
        
        summary_table_row = summary_table_row + 1
        total_stock_volume = 0
    
      End If
      
     'Conditional formatting for Yearly Change
      If Cells(i, 11).Value >= 0 Then
        Cells(i, 11).Interior.ColorIndex = 10
        Else: Cells(i, 11).Interior.ColorIndex = 3
      End If
      
      'Conditional formatting for Percent Change
      If Cells(i, 12).Value >= 0 Then
      Cells(i, 12).Interior.ColorIndex = 10
      Else: Cells(i, 12).Interior.ColorIndex = 3
      End If
      
Next i

Cells(4, 18).Value = greatest_total_volume

Cells(4, 17).Value = greatest_total_volume_ticker

Cells(2, 18).Value = greatest_percent_increase

Cells(2, 17).Value = greatest_percent_increase_ticker

Cells(3, 18).Value = greatest_percent_decrease

Cells(3, 17).Value = greatest_percent_decrease_ticker

Cells(2, 18).NumberFormat = "0.00%"
Cells(3, 18).NumberFormat = "0.00%"
Range("L2:L3000").NumberFormat = "0.00%"
    

End Sub
