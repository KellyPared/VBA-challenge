# VBA-challenge
bootcamp Module 2 Assignment

In this repo, I created a script that loops through various sheets and reads all the stocks for one year and outputs information. It was a long path in a short amount of time to develop the final solution. This is part of my journey and thought process. 

Before the Total Stock Volume, Excel crashed and could not handle my code, which led me to believe I had a looping problem somewhere, making Excel time out even on small Excel files.

### Final Code

Sub Multi_Year_Stock()

    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
   
        Dim Ticker_Name As String
        Dim Ticker_Table_Row As Long
        
        Dim open_value As Double
        Dim last_closed_value As Long
        Dim percent_change As Long
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
           ws.Select
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
                        'percent_change = yearly_change / open_value
                        percent_change = yearly_change / open_value
                    End If
                        ws.Range("K" & Ticker_Table_Row).Value = percent_change
                        ws.Range("K" & Ticker_Table_Row).Style = "Percent"
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
           
        Next ws
    End Sub





### Learning Codes

#### Headers
    Sub Generate_Headers()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
End Sub

#### Sub Move_to_Sheets():
    ' loop through sheets https://www.youtube.com/watch?v=bUMS_BCF08g
    Dim Wksht As Worksheet
    For Each Wksht In ThisWorkbook.Worksheets
    Next Wksht
    End Sub
#### Sub Tcker_Analysis()
    Sub Tickr_Analysis()
    ' https://www.youtube.com/watch?v=nV_oDWJccu8
    ' Find the Ticker value in Range
    Dim Ticker As Range
    Dim count As Integer
    Dim open_value As Double
    Dim closed_value As Double
    Dim yearly_change As Double
    Set Ticker = Range("A2").Find(what:=Range("A2"), LookIn:=xlValues, lookat:=xlWhole)
    count = 1
    ' Copy Ticker Value into I
    Range("I2").Value = Ticker
    ' Assign variables to the Offset Values
    open_value = Ticker.Offset(, 2).Value
    closed_value = Ticker.Offset(count, 5).Value
    yearly_change = closed_value - open_value
    Range("J2").Value = yearly_change
    MsgBox (Ticker)
    MsgBox (open_value)
    MsgBox (closed_value)
    End Sub
    
#### Make Headers for the Columns  

 ## Version 1
 
 '____________________________
' Looping Code for Worksheets
'____________________________

Sub Loop_over_worksheets()
    ' THis will loop through all the worksheets.
    For Each ws In Worksheets
    Dim WorksheetName As String
    WorksheetName = ws.Name
    Next
End Sub
'____________________________
' Making Headers for the worksheets
'____________________________
Sub Generate_Headers()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
End Sub
'____________________________
' Looping Through the Tickers
'____________________________
Sub Ticker_Analysis()
    ' Set an initial variable for holding the Ticker name
    Dim Ticker_Name As String
    ' Set an initial variables for holding the total yearly change
    Dim yearly_change As Double
    yearly_change = 0
    ' Set a variable for holding
    Dim open_value As Double
    Dim closed_value As Double
    open_value = 0
    closed_value = 0
    ' Keep track of the location for each credit card brand in the summary table
    Dim Ticker_Table_Row As Integer
    Ticker_Table_Row = 2
    ' Loop through all tickers in column A find the last row
    lastrow = Cells(Rows.count, 1).End(xlUp).Row   
    For i = 2 To lastrow
        ' If values are not the same keep filtering to the the end.
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Set the Single Ticker Name
            Ticker_Name = Cells(i, 1).Value 
            'Print the TIcker name in Column I
            Range("I" & Ticker_Table_Row).Value = Ticker_Name
            Ticker_Table_Row = Ticker_Table_Row +1
            
        End If
            
    Next i
End Sub


  ###Version 1 without calculations
  Sub Test()
  For Each ws In Worksheets
    'calls name of sheet
    Dim WorksheetName As String
    WorksheetName = ws.Name  
   ' Set an initial variable for holding the Ticker name
    Dim Ticker_Name As String
    ' Set an initial variables for holding the total yearly change
    Dim yearly_change As Double
    yearly_change = 0   
    ' Set an initial variable for Percent change
    Dim percent_change As Double
    percent_change = 0
    ' Set a variable for holding
    Dim open_value As Double
    Dim closed_value As Double
    open_value = 0
    closed_value = 0   
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        Dim Ticker_Table_Row As Integer
        Ticker_Table_Row = 2
        ' Loop through all tickers in column A find the last row
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        For i = 2 To lastrow
            ' If values are not the same keep filtering to the the end.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                If ws.Name Like "2018" Then              
                    ws.Select
                    ' Set the Single Ticker Name
                    Ticker_Name = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
                    open_value = ws.Cells((i - 250), 3).Value
                    'ws.Range("P" & Ticker_Table_Row).Value = open_value
                    closed_value = Cells(i, 6).Value
                    'calculate the yearly change
                    yearly_change = closed_value - open_value
                    'ws.Range("Q" & Ticker_Table_Row).Value = closed_value
                    ws.Range("J" & Ticker_Table_Row).Value = yearly_change
                    If yearly_change >= 0 Then
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbRed
                    Else
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbGreen
                    Ticker_Table_Row = Ticker_Table_Row + 1
                    End If                 
                ElseIf ws.Name Like "2019" Then
                    ws.Select
                    ' Set the Single Ticker Name
                    Ticker_Name = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
                    open_value = ws.Cells((i - 251), 3).Value
                    ws.Range("P" & Ticker_Table_Row).Value = open_value
                    closed_value = Cells(i, 6).Value
                    ws.Range("Q" & Ticker_Table_Row).Value = closed_value
                    Ticker_Table_Row = Ticker_Table_Row + 1\
                Else
                     'Set the Single Ticker Name
                     ws.Select
                    Ticker_Name = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
                    open_value = ws.Cells((i - 252), 3).Value
                    ws.Range("P" & Ticker_Table_Row).Value = open_value
                    closed_value = Cells(i, 6).Value
                    ws.Range("Q" & Ticker_Table_Row).Value = closed_value
                    Ticker_Table_Row = Ticker_Table_Row + 1
                End If
           End If
        Next i Next ws
    End Sub

### Version 2 - Cleaned up Comments and organized variables

Sub Test()

    ' Dim variables and initiate
    For Each ws In Worksheets
    Dim WorksheetName As String
    WorksheetName = ws.Name
    Dim Ticker_Name As String
    Dim yearly_change As Double
    yearly_change = 0
    Dim percent_change As Double
    percent_change = 0
    Dim open_value As Double
    Dim closed_value As Double
    open_value = 0
    closed_value = 0
    Dim Ticker_Table_Row As Integer
    Ticker_Table_Row = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ' Loop through all tickers in column A find the last row
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        For i = 2 To lastrow
            ' If values are not the same keep filtering to the the end.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                If ws.Name Like "2018" Then
                    ws.Select
                    Ticker_Name = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
                    open_value = ws.Cells((i - 250), 3).Value
                    closed_value = Cells(i, 6).Value
                    yearly_change = closed_value - open_value
                    ws.Range("J" & Ticker_Table_Row).Value = yearly_change
                    percent_change = yearly_change / open_value
                    ws.Range("K" & Ticker_Table_Row).Value = percent_change
                    ws.Range("K" & Ticker_Table_Row).NumberFormat = "0.00%"
                    If yearly_change <= 0 Then
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbRed
                    Else
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbGreen
                    End If
                    Ticker_Table_Row = Ticker_Table_Row + 1
                ElseIf ws.Name Like "2019" Then
                    ws.Select
                    Ticker_Name = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
                    open_value = ws.Cells((i - 251), 3).Value
                    closed_value = Cells(i, 6).Value
                    yearly_change = closed_value - open_value
                    ws.Range("J" & Ticker_Table_Row).Value = yearly_change
                    percent_change = yearly_change / open_value
                    ws.Range("K" & Ticker_Table_Row).Value = percent_change
                    ws.Range("K" & Ticker_Table_Row).NumberFormat = "0.00%"
                    If yearly_change <= 0 Then
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbRed
                    Else
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbGreen
                    End If
                    Ticker_Table_Row = Ticker_Table_Row + 1
                Else
                     'Set the Single Ticker Name
                     ws.Select
                    Ticker_Name = ws.Cells(i, 1).Value
                    ws.Range("I" & Ticker_Table_Row).Value = Ticker_Name
                    
                    open_value = ws.Cells((i - 252), 3).Value
                    closed_value = Cells(i, 6).Value
                    yearly_change = closed_value - open_value
                    ws.Range("J" & Ticker_Table_Row).Value = yearly_change
                    percent_change = yearly_change / open_value
                    ws.Range("K" & Ticker_Table_Row).Value = percent_change
                    ws.Range("K" & Ticker_Table_Row).NumberFormat = "0.00%"
                    If yearly_change <= 0 Then
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbRed
                    Else
                        ws.Range("J" & Ticker_Table_Row).Interior.Color = vbGreen
                    End If
                    Ticker_Table_Row = Ticker_Table_Row + 1
                End If
           End If
        Next i
    Next ws
    End Sub

###Practice Data with an error in Percent calculation

<img width="1324" alt="Screen Shot 2022-12-19 at 9 57 36 AM" src="https://user-images.githubusercontent.com/40581033/208455136-d1f74a8a-47c6-4473-93bb-13d31b69fc34.png">

