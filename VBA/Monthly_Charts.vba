Sub Monthly_Charts()

' This with statement determines the last filled row and column
' number on the PE Log sheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Sheets("PE Log").Select
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        LastColumn = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    
' This updates Product Exams
Sheets("Product Exams").Select
Range("N2").Select
ActiveCell.Formula = "=SUM('PE Log'!AD1:AD" & LastRow - 1 & ")"
Range("N3").Select
ActiveCell.Formula = "=COUNTIF('PE Log'!AB1:AB" & LastRow - 1 & ","">0"")"

' This sorts the Observations By Program numerically
Sheets("Observations by Program").Select
Range("R4:T8").Select
ActiveWorkbook.Worksheets("Observations by Program").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Observations by Program").Sort.SortFields.Add Key _
        :=Range("S4:S8"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Observations by Program").Sort
        .SetRange Range("R3:T8")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
' This sorts the Observations By Category alphabetically
Sheets("Observations by Category").Select
Range("F20:F33").Select
Selection.NumberFormat = "General"
Range("E20:F31").Select
    ActiveWorkbook.Worksheets("Observations by Category").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Observations by Category").Sort.SortFields.Add Key _
        :=Range("E20"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Observations by Category").Sort
        .SetRange Range("E20:F31")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

' This calculates the number of Observations By Category alphabetically
Range("F20").Select
ActiveCell.Formula = "=SUM('PE Log'!P" & LastRow & ":P" & LastRow & ")"
Range("F21").Select
ActiveCell.Formula = "=SUM('PE Log'!O" & LastRow & ":O" & LastRow & ")"
Range("F22").Select
ActiveCell.Formula = "=SUM('PE Log'!S" & LastRow & ":S" & LastRow & ")"
Range("F23").Select
ActiveCell.Formula = "=SUM('PE Log'!W" & LastRow & ":W" & LastRow & ")"
Range("F24").Select
ActiveCell.Formula = "=SUM('PE Log'!X" & LastRow & ":X" & LastRow & ")"
Range("F25").Select
ActiveCell.Formula = "=SUM('PE Log'!Q" & LastRow & ":Q" & LastRow & ")"
Range("F26").Select
ActiveCell.Formula = "=SUM('PE Log'!U" & LastRow & ":U" & LastRow & ")"
Range("F27").Select
ActiveCell.Formula = "=SUM('PE Log'!V" & LastRow & ":V" & LastRow & ")"
Range("F28").Select
ActiveCell.Formula = "=SUM('PE Log'!R" & LastRow & ":R" & LastRow & ")"
Range("F29").Select
ActiveCell.Formula = "=SUM('PE Log'!Z" & LastRow & ":Z" & LastRow & ")"
Range("F30").Select
ActiveCell.Formula = "=SUM('PE Log'!T" & LastRow & ":T" & LastRow & ")"
Range("F31").Select
ActiveCell.Formula = "=SUM('PE Log'!Y" & LastRow & ":Y" & LastRow & ")"

' This sorts the Observations By Category numerically
Range("E20:F31").Select
    ActiveWorkbook.Worksheets("Observations by Category").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Observations by Category").Sort.SortFields.Add Key _
        :=Range("F20:F31"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Observations by Category").Sort
        .SetRange Range("E20:F31")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Range("G33").Select
End Sub
