Sub PN_Column()
' This with statement determines the last filled row number on the Tags sheet
    Sheets("Tags_April-June 2015").Select
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With

    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Value = "PN"

' This establishes the exception list for trim 11 instead of 9
    Dim row As Integer
    Dim VList
    Set VList = CreateObject("Scripting.Dictionary")
    VList.Add "Plane_1", "1"
    VList.Add "Plane_2", "2"
    VList.Add "Plane_3", "3"
    VList.Add "Plane_4", "4"
    
' This loop goes and enters the simplified PN from the Part Number
    For row = 2 To LastRow
        searchString = CStr(Cells(row, 5).Text)
        If VList.Exists(searchString) Then
            Cells(row, 6).Select
            ActiveCell.Formula = "=LEFT(G" & row & ",11)"
        Else
            Cells(row, 6).Select
            ActiveCell.Formula = "=LEFT(G" & row & ",9)"
        End If
    Next row
    
End Sub