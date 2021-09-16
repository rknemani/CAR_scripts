Sub CopyRows()
'
' New_CAR_Meeting Macro
'
' Keyboard Shortcut: Ctrl+m

' This selects column CZ and formats it to a number with no decimals
' and autofits the column
    Columns("CZ:CZ").Select
    Selection.NumberFormat = "0"
    Columns("CZ:CZ").EntireColumn.AutoFit
    
' The Summary sheet is selected and the number of open ADQ CARs is determined from CZ6 which is determined in the autosort macro
' This value is then added to two (the Summary and Hidden Template are 1 and 2) and stored in SummaryLastRow
    Sheets("Summary").Select
    Range("CZ6").Select
    SummaryLastRow = Range("CZ6").Value + 2
        
' This loop goes through every CAR data sheet and selects it and will move to the next copy operation
' if the last meeting date listed is today's current date
    For i = 3 To SummaryLastRow
    
    Sheets(i).Select
    Dim TempLastRow As Long
    With ActiveSheet
        TempLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    If Cells(TempLastRow, 1) = Date And i = SummaryLastRow Then
    MsgBox "Today's meeting date has already been entered"
    Sheets("Summary").Select
    Range("V1").Select
    
    ElseIf Cells(TempLastRow, 1) = Date Then
    Sheets(i + 1).Select
    
    Else
' This with statement determines the last row number that is filled and saved into LastRow
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
' Fills the next empty meeting date cell with the current date
    Cells(LastRow + 1, 1).Value = Date
    
' This loop goes through every CAR data sheet and copies the previous row into
' the current meeting date row
    Dim j As Integer
    For j = 2 To 27
    Cells(LastRow + 1, j).Value = Cells(LastRow, j).Value
    Next j
    End If
    
    Next i
End Sub
