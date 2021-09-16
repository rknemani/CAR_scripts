Sub CAR_Sort()
'
' CAR_Sort Macro
'
' Keyboard Shortcut: Ctrl+g

    Dim LastRow As Long
    LastRow = Range("CZ4").Value
    LastRowIndex = Range("CZ4").Value + 1
    
    
' This finds the number of Open CARs
    Range("CZ5").Select
    ActiveCell.Formula = "=COUNTIF(S2:S" & LastRow & ", ""Open"")"
    Dim OpenNum As Long
    OpenNum = Range("CZ5").Value + 1
    
' This finds the number of Open ADQ CARs
    Range("CZ6").Select
    ActiveCell.Formula = "=COUNTIFS(S2:S" & LastRow & ", ""Open"",U2:U" & LastRow & ", ""ADQ"")"
    Dim OpenADQNum As Long
    OpenADQNum = Range("CZ6").Value + 1
    
' This sorts all the CARs by Closure and puts closed CARs at the bottom
    Range("A1:U" & LastRow & "").Select
    ActiveWorkbook.Worksheets("Summary").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Summary").Sort.SortFields.Add Key:=Range("S2:S" & LastRow & "") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Summary").Sort
        .SetRange Range("A1:U" & LastRow & "")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
' This sorts all the CARs by Group and puts ADE CARs at the next lowest level
    Range("A1:U" & OpenNum & "").Select
    ActiveWorkbook.Worksheets("Summary").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Summary").Sort.SortFields.Add Key:=Range("U2:U" & OpenNum & "") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Summary").Sort
        .SetRange Range("A1:U" & OpenNum & "")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
' This sorts all the ADQ CARs by Issue Date
    ActiveWindow.SmallScroll Down:=-15
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "CAR #"
    Range("A1:U" & OpenADQNum & "").Select
    ActiveWorkbook.Worksheets("Summary").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Summary").Sort.SortFields.Add Key:=Range("B2:B" & OpenADQNum & "") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Summary").Sort
        .SetRange Range("A1:U" & OpenADQNum & "")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A" & LastRow + 1 & "").Select
    
' This finds closed CARs that have not been hidden and moves them to the end
' and hides them
    For i = 3 To LastRowIndex - 1
    If Cells(i, 19) <> "Open" Then
        SheetName = Cells(i, 1).Text
        
        On Error GoTo ErrorHandler
        Sheets("" & SheetName & "").Select
        Sheets("" & SheetName & "").Move After:=Sheets(LastRowIndex)
        Sheets("" & SheetName & "").Select
        ActiveWindow.SelectedSheets.Visible = False
        Sheets("Summary").Select
    End If
    Next i
    Exit Sub
ErrorHandler:
i = i + 1
Resume Next
    
End Sub

