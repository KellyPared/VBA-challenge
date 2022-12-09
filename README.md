# VBA-challenge
bootcamp Module 2 Assignment

In this repo, I created a script that loops through various sheets and reads all the stocks for one year and outputs information.

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


    


