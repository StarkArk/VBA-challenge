Sub stock_results():
    
For Each ws In Worksheets
    
    ws.Select
    
    ' Assign Variables
        
        ' Provides the last row # for the ticker collumn
        Dim ticker_entries As Long
        ticker_entries = Application.WorksheetFunction.CountA(Range("A:A"))
        Dim stock_ticker As String
        stcok_ticker = ""
        
        ' Counter to track the current row for the ticker in the Summary Table
        Dim counter As Integer
        counter = 2
        
        ' For the Summary Table
        Dim row_firstopen As Long
        row_firstopen = 2
        Dim row_lastclose As Long
        row_last_close = 0
        Dim open_value As Single
        open_value = 0
        Dim closing_value As Single
        closing_value = 0
        Dim yearly_change As Single
        yearly_change = 0
        Dim yearly_percent_change As Double
        percent_change = 0
        Dim total_stock_volume As Double
        total_stock_volume = 0
        
        ' Summary Table for Max/Min values
        Dim greatest_percent_increase As Single
        greatest_percent_increase = -1
        Dim greatest_percent_decrease As Single
        greatest_percent_decrease = 0
        Dim greatest_total_volume As Double
        greatest_total_volume = 0
        Dim greatest_percent_increase_ticker As String
        greatest_percent_increase = -1
        Dim greatest_percent_decrease_ticker As String
        greatest_percent_decrease = 0
        Dim greatest_total_volume_ticker As String
        greatest_total_volume = 0
        
        ' Provides the last row # for the summary table ticker collumn
        Dim summary_ticker_entries As Integer
        summary_ticker_entries = 0

    
    ' Format first column as string
        Range("A:A").NumberFormat = "@"
        
    ' Assign Column headings and Cell labels
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change ($)"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
    
    
    ' Populate Summary Table Values
    For i = 1 To ticker_entries
        
        If Cells(i + 1, 1).Value <> stock_ticker Then
            stock_ticker = Cells(i + 1, 1)
            Cells(counter, 9).Value = stock_ticker                                          ' Ticker
            row_lastclose = i
            
            If counter > 2 Then
                
                closing_value = Cells(i, 6).Value
                yearly_change = closing_value - open_value
                yearly_percent_change = yearly_change / open_value
                'total_stock_volume = Application.WorksheetFunction.Sum(Range(Cells(row_firstopen, 7), Cells(row_lastclose, 7)))
                total_stock_volume = Application.WorksheetFunction.Sum(Range("G" & row_firstopen, "G" & row_lastclose))
                'total_stock_volume = 0
                'For j = row_firstopen To row_lastclose
                '    volume
                Cells(counter - 1, 10).Value = yearly_change                                ' Yearly Change
                Cells(counter - 1, 11).Value = yearly_percent_change                        ' Percent Change
                Cells(counter - 1, 12).Value = total_stock_volume                           ' Total Stock Volume
               
                
               
            End If
            
            row_firstopen = i + 1
            counter = counter + 1
            open_value = Cells(i + 1, 3).Value
            
        End If
      
    Next i
    
    
    ' Populate Min/Max Summary Table Values
    
    summary_ticker_entries = Application.WorksheetFunction.CountA(Range("L:L"))
    
    For j = 2 To summary_ticker_entries
        
        If Range("K" & j).Value > greatest_percent_increase Then
            greatest_percent_increase = Range("K" & j).Value
            greatest_percent_increase_ticker = Range("I" & j).Value
        End If
        
        If Range("K" & j).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Range("K" & j).Value
            greatest_percent_decrease_ticker = Range("I" & j).Value
        End If
        
        If Range("L" & j).Value > greatest_total_volume Then
            greatest_total_volume = Range("L" & j).Value
            greatest_total_volume_ticker = Range("I" & j).Value
        End If
    
    Next j
    
    Range("P" & 2).Value = greatest_percent_increase_ticker                                 ' Greatest % Increase Value
    Range("P" & 3).Value = greatest_percent_decrease_ticker                                 ' Greatest % Increase Ticker
    Range("P" & 4).Value = greatest_total_volume_ticker                                     ' Greatest % Decrease Value
    Range("Q" & 2).Value = greatest_percent_increase                                        ' Greatest % Decrease Ticker
    Range("Q" & 3).Value = greatest_percent_decrease                                        ' Greatest Total Volume Value
    Range("Q" & 4).Value = greatest_total_volume                                            ' Greatest Total Volume Ticker
    
    
    ' Format Columns/Cells Summary Tables
        Columns("J:L").EntireColumn.AutoFit
        Columns("O:O").EntireColumn.AutoFit
        
        ' Yearly Change
        Range("J2").Select
            Range(Selection, Selection.End(xlDown)).Select
                Selection.NumberFormat = "0.00"
                'Selection.
                Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:="=0"
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    Selection.FormatConditions(1).Interior.ColorIndex = 10
                Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:="=0"
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    Selection.FormatConditions(1).Interior.ColorIndex = 46
        
        ' Percent Change
        Range("K2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.NumberFormat = "0.00%"
        
        ' Total Stock Volume
        Range("L2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.NumberFormat = "0"
        Columns("L:L").EntireColumn.AutoFit
        
        ' Greatest(Increase, Decrease, Total Volume)
        Columns("O:O").EntireColumn.AutoFit
        Range("Q2:Q3").Select
            Selection.NumberFormat = "0.00%"
        Range("Q4").Select
            Selection.NumberFormat = "0.00E+00"
            
Next ws
            
End Sub
