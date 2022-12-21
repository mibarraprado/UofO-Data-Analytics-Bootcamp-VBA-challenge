Attribute VB_Name = "Stock_data_summary"
'This Vba code runs the same macro to summarize the stock data on everysheet of the workbook
Sub Multiple_year_stock_data()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stock_data_summary
    Next
    Application.ScreenUpdating = True
End Sub

'The macro below summarizes the yearly stock data

Sub Stock_data_summary()
 Dim yearly_change As Double
 Dim percent_change As Double
 Dim stock_volume As Double
 Dim first_day_year_open_value, last_day_year_close_value As Double
 Dim lastrow, last_column, lastrow_summary As Integer
 
 Dim greatest_percent_increase As Double
 Dim greatest_percent_decrease As Double
 Dim greatest_stock_volume As Double
 Dim greatest_percent_increase_Ticker As String
 Dim greatest_percent_decrease_Ticker As String
 Dim greatest_stock_volume_Ticker As String
 
 'Determine the last row and last column of the dataset
 lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 last_column = Cells(1, Columns.Count).End(xlToLeft).Column
 lastrow_summary = 1
 
 ' Create summary tables
 
 Cells(1, last_column + 2).Value = "Ticker"
 Cells(1, last_column + 3).Value = "Yearly Change"
 Cells(1, last_column + 4).Value = "Percent Change"
 Cells(1, last_column + 5).Value = "Total Stock Volume"
 
 Cells(1, last_column + 9).Value = "Ticker"
 Cells(1, last_column + 10).Value = "Value"
 
 Cells(2, last_column + 8).Value = "Greatest % increase"
 Cells(3, last_column + 8).Value = "Greatest % decrease"
 Cells(4, last_column + 8).Value = "Greatest total volume"
 
  ' initial values of variables for first summary table
  
 stock_volume = Cells(2, 7).Value
 first_day_year_open_value = Cells(2, 3).Value
 last_day_year_close_value = Cells(2, 6).Value
 
 
 ' Cycle to read values of dataset and calculate stock volume and changes
 
 For i = 2 To lastrow
 'Compare if ticker name of row i is equel to the name of the row below
    'if so, add stock volume of the row below to the total
    If (Cells(i, 1).Value = Cells(i + 1, 1).Value) Then
    stock_volume = stock_volume + Cells(i + 1, 7).Value
    ' determine if row i is the last date of year and then capture close value
     If (CLng(Cells(i, 2).Value) < CLng(Cells(i + 1, 2).Value)) Then
     last_day_year_close_value = Cells(i + 1, 6).Value
     End If
    
    Else
' If ticker name of row i is not the same as the one of the row below,
' then we are done reading all the data of the ticker and we can calculate
' and populate summary table values

    'Calculate yearly change
    yearly_change = last_day_year_close_value - first_day_year_open_value
    
    'Calculate percent change
    percent_change = last_day_year_close_value / first_day_year_open_value - 1

    'Populate Ticker Name in the summary table
    Cells(lastrow_summary + 1, last_column + 2).Value = Cells(i, 1).Value
    
    'Populate yearly change in the summary table
    Cells(lastrow_summary + 1, last_column + 3).Value = yearly_change
        'Color format for yearly change cells
        If (yearly_change < 0) Then
        Cells(lastrow_summary + 1, last_column + 3).Interior.Color = RGB(255, 0, 0)
        
        Else
        Cells(lastrow_summary + 1, last_column + 3).Interior.Color = RGB(0, 255, 0)
        End If
    
     'Populate percent change in the summary table
    Cells(lastrow_summary + 1, last_column + 4).Value = FormatPercent(percent_change, 2)
    
    'Output of stock volume in the summary table
    
    Cells(lastrow_summary + 1, last_column + 5).Value = stock_volume
    
    'restart values for the next ticker
    
     first_day_year_open_value = Cells(i + 1, 3).Value
     last_day_year_close_value = Cells(i + 1, 6).Value
     stock_volume = Cells(i + 1, 7).Value
     lastrow_summary = lastrow_summary + 1
    
    End If
 Next i

' --- Bonus-----

'Second summary table calculations
' set initial values
greatest_percent_increase = Cells(2, last_column + 4).Value
greatest_percent_decrease = Cells(2, last_column + 4).Value
greatest_stock_volume = Cells(2, last_column + 5).Value

For j = 2 To lastrow_summary

    If (greatest_percent_increase < Cells(j, last_column + 4).Value) Then
    greatest_percent_increase = Cells(j, last_column + 4).Value
    greatest_percent_increase_Ticker = Cells(j, last_column + 2).Value
    End If
    
    If (greatest_percent_decrease > Cells(j, last_column + 4).Value) Then
    greatest_percent_decrease = Cells(j, last_column + 4).Value
    greatest_percent_decrease_Ticker = Cells(j, last_column + 2).Value
    End If
    
    If (greatest_stock_volume < Cells(j, last_column + 5).Value) Then
    greatest_stock_volume = Cells(j, last_column + 5).Value
    greatest_stock_volume_Ticker = Cells(j, last_column + 2).Value
    End If

Next j

 Cells(2, last_column + 9).Value = greatest_percent_increase_Ticker
 Cells(2, last_column + 10).Value = FormatPercent(greatest_percent_increase, 2)
 
 Cells(3, last_column + 9).Value = greatest_percent_decrease_Ticker
 Cells(3, last_column + 10).Value = FormatPercent(greatest_percent_decrease, 2)
 
 Cells(4, last_column + 9).Value = greatest_stock_volume_Ticker
 Cells(4, last_column + 10).Value = greatest_stock_volume
End Sub
