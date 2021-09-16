Dim Last_Non_0_Row As Integer
Dim j As Integer
j = 37
    Do While True
        j = j - 1
        
        If Cells(j, 4).Value <> 0 Or Cells(j, 3).Value <> 0 Or Cells(j, 2).Value <> 0 Then
            Exit Do
        End If
    Loop
 Last_Non_0_Row = j

ActiveSheet.ChartObjects("Chart 2").Activate
ActiveChart.SetSourceData Source:=Range("'Product Exams'!$A$27:$D$" & Last_0_Row & "")