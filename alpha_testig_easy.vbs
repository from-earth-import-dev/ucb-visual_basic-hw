Sub stocks_analysis():
    For Each ws In Worksheets
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Total Stock Volume"
        
        Dim row As Long
        Dim summary_table_row As Long
        Dim per_stock_volume As Double
        Dim lastrow As Long
        
        summary_table_row = 2
        per_stock_volume = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        
        For row = 2 To lastrow
            'If current row's ticker symbol and next row ticker symbol match, then add to the per_stock_volume
            If ws.Cells(row, 1).Value = ws.Cells(row + 1, 1).Value Then
                per_stock_volume = per_stock_volume + ws.Cells(row, 7).Value
            Else
                'Add current row to per_stock_volume
                per_stock_volume = per_stock_volume + ws.Cells(row, 7).Value
                'Add ticker symbol to summary table
                ws.Cells(summary_table_row, 9).Value = ws.Cells(row, 1).Value
                'Add stock volume to summary table
                ws.Cells(summary_table_row, 10).Value = per_stock_volume
                'Move summary table row down 1, reset per_stock_volume
                summary_table_row = summary_table_row + 1
                'MsgBox (summary_table_row)
                'MsgBox (per_stock_volume)
                per_stock_volume = 0
                'MsgBox ("Stock volume reset : " & per_stock_volume)
            End If
        
        Next row
    Next ws
End Sub