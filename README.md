# VBA-challenge
bootcamp Module 2 Assignment

In this repo, I created a script that loops through various sheets and reads all the stocks for one year and outputs information.

### Learning Codes

#### Sub Move_to_Sheets():
    ' loop through sheets https://www.youtube.com/watch?v=bUMS_BCF08g
    Dim Wksht As Worksheet
    For Each Wksht In ThisWorkbook.Worksheets

    Next Wksht
    End Sub

#### Sub Tcker_Analysis()
    ' https://www.youtube.com/watch?v=nV_oDWJccu8
    
    ' Find the Ticker value in Range
    Dim Ticker As Range
    Dim count As Integer
    Set Ticker = Range("A2").Find(what:=Range("A2"), LookIn:=xlValues, lookat:=xlWhole)
    count = 1
    ' Copy Ticker Value into I
    Range("I2").Value = Ticker
    
    ' Assign variables to the Offset Values
    open_value = Ticker.Offset(, 2).Value
    closed_value = Ticker.Offset(count, 5).Value
    
    MsgBox (Ticker)
    MsgBox (open_value)
    MsgBox (closed_value)
    

#### Make Headers for the Columns

    
    
    
    


