Sub Format_Date()
'
' Format_Date Macro
'

'
    Columns("K:K").Select
    Selection.NumberFormat = "m/d/yyyy"
    
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "K").End(xlUp).row
    End With
    
    For i = 2 To LastRow
        Range("K" & i & "").Formula = Range("K" & i & "").Formula
    Next i

End Sub
