Sub Multi_Year_Stock()

    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
   
        Dim Ticker_Name As String
        Dim Ticker_Table_Row As Long
        
        Dim open_value As Double
        Dim last_closed_value As Double
        Dim percent_change As Double
        Dim current_row As Long
        Dim starter_row As Long
       
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        change_row = 2
        Ticker_Table_Row = 2
        
        For current_row = 2 To lastRow
           'ws.Select
            If ws.Cells(current_row + 1, 1).Value <> ws.Cells(current_row, 1).Value Then
                    'ticker
                    ws.Range("I" & Ticker_Table_Row).Value = ws.Cells(current_row, 1).Value
                    
                    last_closed_value = ws.Cells(current_row, 6).Value
                    open_value = ws.Cells(change_row, 3).Value
                    yearly_change = last_closed_value - open_value
                    ws.Range("J" & Ticker_Table_Row).Value = yearly_change
                    
                    'yearly_change = last_closed_value - open_value
                    If ws.Range("J" & Ticker_Table_Row).Value < 0 Then
                            ws.Range("J" & Ticker_Table_Row).Interior.Color = vbRed
                    Else
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbGreen
                    End If
                    
                    If ws.Cells(change_row, 3).Value = 0 Then
                        percent_change = 0
                    Else
                        percent_change = yearly_change / open_value
                        'percent_change = last_closed_value / open_value
                    End If
                        ws.Range("K" & Ticker_Table_Row).Value = percent_change
                        ws.Range("K" & Ticker_Table_Row).NumberFormat = "00.00%"
                    change_row = current_row + 1
                    Ticker_Table_Row = Ticker_Table_Row + 1
                    
                      
                   End If
                   
                'calc total = sum of  I and J WorksheetFinction.Sum
                ws.Range("L" & Ticker_Table_Row).Value = WorksheetFunction.Sum(Range(ws.Cells(change_row, 7), ws.Cells(current_row, 7)))
                    
        
        Next current_row
        
            'checking for greatest and least increase by reassigning the values to the variables during the loop.
            Dim new_last_row As Long
            new_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
            change_volume_row = 2
            NEWTicker_Table_Row = 2
            Greatest_Percent_Increase = ws.Cells(2, 11).Value
            Greatest_Percent_Decrease = ws.Cells(2, 11).Value
            Greatest_Total_Stock = ws.Cells(2, 12).Value
            
            
            For new_ticker_row = 2 To lastRow:
                If ws.Cells(new_ticker_row, 11) > Greatest_Percent_Increase Then
                    Greatest_Percent_Increase = ws.Cells(new_ticker_row, 11).Value
                    
                ElseIf ws.Cells(new_ticker_row + 1, 11) < Greatest_Percent_Decrease Then
                    Greatest_Percent_Decrease = ws.Cells(new_ticker_row, 11).Value
                    
                ElseIf ws.Cells(new_ticker_row + 1, 12) > Greatest_Total_Stock Then
                    Greatest_Total_Stock = ws.Cells(new_ticker_row, 12).Value

                
                End If
                Next new_ticker_row
                'place in new table
                ws.Cells(NEWTicker_Table_Row, 17).Value = Greatest_Percent_Increase
                ws.Cells(NEWTicker_Table_Row + 1, 17).Value = Greatest_Percent_Decrease
                ws.Range("Q2:Q3").NumberFormat = "0.00%"
                ws.Cells(NEWTicker_Table_Row + 2, 17).Value = Greatest_Total_Stock
                'ws.Cells(NEWTicker_Table_Row, 16).Value = Greatest_Increase_Ticker
                'ws.Range(NEWTicker_Table_Row + 1, 16).Value = Greatest_Decrease_Ticker
                'ws.Range(NEWTicker_Table_Row + 2, 16).Value = Greatest_Total_Ticker
        Next ws
    End Sub

