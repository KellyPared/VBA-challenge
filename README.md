# VBA-challenge
bootcamp Module 2 Assignment

In this repo, I created a script that loops through various sheets and reads all the stocks for one year and outputs information. It was a long path in a short amount of time to develop the final solution. This is part of my journey and thought process. 

Before the Total Stock Volume, Excel crashed and could not handle my code, which led me to believe I had a looping problem somewhere, making Excel time out even on small Excel files.

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
                    
                    
                    
                    Ticker_Table_Row = Ticker_Table_Row + 1
                
                    
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
        Next i
           
        Next ws
        
        
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


