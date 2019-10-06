Attribute VB_Name = "Module1"

Sub Stock_stuff()

' designate the file as a worksheet

Dim ws As Worksheet

' Couldn't successfully debug my code so I found this command online that doesn't fix the error, but allows the script to continue running.
' Originally ran a "On Error GoTo 0" so that i could open the debugger on the line on which the error occurred.
' I think the error is in regard to dividing by zero on the stocks that don't have a year_open value.
' Made a note lower in the code designating the code in question.

On Error Resume Next


' Column headers and their location

For Each ws In ThisWorkbook.Worksheets
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
Summary_Table_Row = 2

' Tells  program to start at row 2 and move through to the last row
For i = 2 To ws.UsedRange.Rows.Count

'Begin worksheet loop

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'designates the value of specific variables

ticker = ws.Cells(i, 1).Value

vol = vol + ws.Cells(i, 7).Value

year_open = ws.Cells(i, 3).Value

year_close = ws.Cells(i, 6).Value

yearly_change = year_close - year_open

' This is the line that causes the error that I have been unable to reconcile.
percent_change = year_close / year_open

'This part designates the location that the variables designated immediately above are going to be placed within the spreadsheet

ws.Cells(Summary_Table_Row, 9).Value = ticker
ws.Cells(Summary_Table_Row, 10).Value = yearly_change
ws.Cells(Summary_Table_Row, 11).Value = percent_change
ws.Cells(Summary_Table_Row, 12).Value = vol
Summary_Table_Row = Summary_Table_Row + 1
vol = 0
End If

Next i
ws.Columns("K").NumberFormat = "0.00%"


' I used the Macro recorder in the dev bar to record this conditional formatting code, but I couldn't get it to loop to each page so I just did it in Excel.

Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.499984740745262
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("N6").Select
    ActiveWindow.SmallScroll Down:=-18
    Application.Goto Reference:="ColorMacro"
    ActiveWorkbook.Save
Next
End Sub
