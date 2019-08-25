Sub moderate():
    For Each ws In Worksheets
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        Dim row As Long
        Dim summary_table_row As Long
        Dim per_stock_volume As Double
        Dim lastrow As Long
        Dim opening_price As Double
        Dim closing_price As Double
        Dim max As Double
        Dim min As Double
        Dim volumeMax As Double
        
        
        summary_table_row = 2
        per_stock_volume = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
        max = 0
        min = 0
        volumeMax = 0
        
        
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
    Dim newlastrow As Integer
    
    newlastrow = ws.Cells(Rows.Count, 11).End(xlUp).row
    
    'Find the max change
    For row = 2 To newlastrow:
            If ws.Cells(row, 11).Value > max Then
                max = ws.Cells(row, 11).Value
                ws.Range("P2").Value = ws.Cells(row, 9)
                ws.Range("Q2").Value = ws.Cells(row, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
    'Find the min change
            If ws.Cells(row, 11).Value < min Then
                min = ws.Cells(row, 11).Value
                ws.Range("P3").Value = ws.Cells(row, 9)
                ws.Range("Q3").Value = ws.Cells(row, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
    'Find the largest volume
            If ws.Cells(row, 12).Value > volumeMax Then
                volumeMax = ws.Cells(row, 12).Value
                ws.Range("P4").Value = ws.Cells(row, 9)
                ws.Range("Q4").Value = ws.Cells(row, 12).Value
            End If
            
        Next row
        max = 0
        min = 0
        volumeMax = 0
        ws.Columns("I:Q").AutoFit
        
    Next ws
End Sub