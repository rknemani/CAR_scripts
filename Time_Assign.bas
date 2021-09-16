Sub Assign_Time()
'
' Time assignment Macro
'
'
Dim LastRow As Long
Dim searchString As String
Dim Col As Integer

With ActiveSheet
    LastRow = .Cells(.Rows.Count, "G").End(xlUp).Row
End With

Col = 8 ' This tells you which colum to output in by number
' Column G is 7 so if this changes you will need to change the number

For i = 2 To LastRow
    searchString = CStr(Cells(i, 7).Value)
    If searchString = "PR" Then
        Cells(i, Col).Value = 40
        
    ElseIf searchString = "MRB_INLINE" Or searchString = "PE MGI" _
    Or searchString = "WAWF" Then
        Cells(i, Col).Value = 1
    
    ElseIf searchString = "CRR/CTR" Or searchString = "DCA" _
    Or searchString = "PE SOF" Then
        Cells(i, Col).Value = 0.5
        
    ElseIf searchString = "MRB_PR" Then
        Cells(i, Col).Value = 0.25
    
    Else
        Cells(i, Col).Value = 0
    End If
Next i

End Sub
