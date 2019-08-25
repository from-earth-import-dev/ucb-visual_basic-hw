Sub moderate():
    For Each ws In Worksheets
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        Dim row As Long
        Dim summary_table_row As Long
        Dim per_stock_volume As Double
        Dim lastrow As Long
        Dim opening_price As Double
        Dim closing_price As Double
        
        summary_table_row = 2
        per_stock_volume = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        
        For row = 2 To lastrow
            
            'If this is the first row of new ticker symbol, grab opening price
            If ws.Cells(row, 1).Value <> ws.Cells(row - 1, 1).Value Then
                opening_price = ws.Cells(row, 3).Value
            
            End If

            'If current row's ticker symbol and next row ticker symbol match, then add to the per_stock_volume
            If ws.Cells(row, 1).Value = ws.Cells(row + 1, 1).Value Then
                per_stock_volume = per_stock_volume + ws.Cells(row, 7).Value

            Else
                'Add current row to per_stock_volume
                per_stock_volume = per_stock_volume + ws.Cells(row, 7).Value
                
                'Grab closing price
                closing_price = ws.Cells(row, 6).Value
                
                'Add ticker symbol to summary table
                ws.Cells(summary_table_row, 9).Value = ws.Cells(row, 1).Value
                
                'Add yearly_change to summary table
                ws.Cells(summary_table_row, 10).Value = closing_price - opening_price
                
                'Check for divide by 0
                If opening_price = 0 Then
                    ws.Cells(summary_table_row, 10).Value = closing_price
                    ws.Cells(summary_table_row, 11).Value = 0
                    ws.Cells(summary_table_row, 12).Value = per_stock_volume
                    per_stock_volume = 0
                    summary_table_row = summary_table_row + 1
                    GoTo NextIteration
                End If
                
                'Add percent_change to summary table
                ws.Cells(summary_table_row, 11).Value = (closing_price - opening_price) / opening_price
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                
                'Add stock volume to summary table
                ws.Cells(summary_table_row, 12).Value = per_stock_volume

                'Conditional formatting for yearly change cells
                If ws.Cells(summary_table_row, 10).Value > 0# Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
                
                'Move summary table row down 1, reset per_stock_volume
                summary_table_row = summary_table_row + 1
                
                'Reset counters for next company
                per_stock_volume = 0
                opening_price = 0
                closing_price = 0
NextIteration:
            End If
        Next row
        Columns("I:L").AutoFit

    Next ws
End Sub