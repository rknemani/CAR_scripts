Sub Delegator()
'
' Delegator Macro
'
' Keyboard Shortcut: Ctrl+d
'
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    Dim Delegator As Long
    Dim Cage_Code As Long
    Dim d As Long
    Dim cc As Long

    d = Range("A1:AZ1").Find(What:="Delegator", LookIn:=xlValues, LookAt:=xlWhole, _
    MatchCase:=False, SearchFormat:=False).Column

    cc = Range("A1:AZ1").Find(What:="Cage Code", LookIn:=xlValues, LookAt:=xlWhole, _
    MatchCase:=False, SearchFormat:=False).Column
    
' Person_1's List of CAGE Codes
    Dim Person_1_List
    Set Person_1_List = CreateObject("Scripting.Dictionary")
    Person_1_List.Add "05722", "1"
    Person_1_List.Add "07754", "2"
    Person_1_List.Add "09790", "3"
    Person_1_List.Add "12813", "4"
    Person_1_List.Add "16472", "5"
    Person_1_List.Add "23974", "6"
    Person_1_List.Add "34087", "7"
    Person_1_List.Add "39443", "8"
    Person_1_List.Add "58635", "9"
    Person_1_List.Add "62659", "10"
    Person_1_List.Add "66449", "11"
    Person_1_List.Add "73760", "12"
    Person_1_List.Add "81929", "13"
    Person_1_List.Add "78943", "14"
    Person_1_List.Add "81873", "15"
    Person_1_List.Add "81982", "16"
    Person_1_List.Add "82106", "17"
    Person_1_List.Add "86090", "18"
    Person_1_List.Add "94821", "19"
    Person_1_List.Add "97415", "20"
    Person_1_List.Add "25583", "21"
    Person_1_List.Add "99931", "22"
'**************************************************************

' Person_2's List of CAGE Codes
    Dim Person_2_List
    Set Person_2_List = CreateObject("Scripting.Dictionary")
    Person_2_List.Add "04939", "1"
    Person_2_List.Add "05088", "2"
    Person_2_List.Add "42827", "3"
    Person_2_List.Add "11673", "4"
    Person_2_List.Add "12837", "5"
    Person_2_List.Add "13002", "6"
    Person_2_List.Add "19904", "7"
    Person_2_List.Add "54112", "8"
'**************************************************************

' Person_3's List of CAGE Codes
    Dim Person_3_List
    Set Person_3_List = CreateObject("Scripting.Dictionary")
    Person_3_List.Add "25598", "1"
    Person_3_List.Add "29183", "2"
    Person_3_List.Add "14248", "3"
    Person_3_List.Add "78710", "4"
    Person_3_List.Add "90073", "5"
    Person_3_List.Add "31218", "6"
    Person_3_List.Add "99789", "7"
    Person_3_List.Add "63123", "8"
    Person_3_List.Add "09245", "9"
'**************************************************************

' Person_4's List of CAGE Codes
    Dim Person_4_List
    Set Person_4_List = CreateObject("Scripting.Dictionary")
    Person_4_List.Add "03104", "1"
    Person_4_List.Add "08748", "2"
    Person_4_List.Add "12536", "3"
    Person_4_List.Add "17610", "4"
    Person_4_List.Add "21215", "5"
    Person_4_List.Add "26838", "6"
    Person_4_List.Add "30974", "7"
    Person_4_List.Add "50502", "8"
'**************************************************************

' Person_5's List of CAGE Codes
    Dim Person_5_List
    Set Person_5_List = CreateObject("Scripting.Dictionary")
    Person_5_List.Add "25167", "1"
    Person_5_List.Add "17472", "2"
    Person_5_List.Add "23518", "3"
    Person_5_List.Add "77745", "4"
    Person_5_List.Add "81860", "5"
    Person_5_List.Add "72314", "6"
    Person_5_List.Add "73030", "7"
    Person_5_List.Add "33068", "8"
    Person_5_List.Add "81205", "9"
    Person_5_List.Add "86985", "10"
    Person_5_List.Add "96124", "11"
    Person_5_List.Add "96487", "12"
    Person_5_List.Add "99251", "13"
'**************************************************************

' Person_6's List of CAGE Codes
    Dim Person_6_List
    Set Person_6_List = CreateObject("Scripting.Dictionary")
	Person_6_List.Add "91417", "1"
    Person_6_List.Add "07639", "2"
    Person_6_List.Add "10933", "3"
    Person_6_List.Add "12779", "4"
    Person_6_List.Add "18118", "5"
    Person_6_List.Add "6Q425", "6"
    Person_6_List.Add "29616", "7"
    Person_6_List.Add "53449", "8"
    Person_6_List.Add "59384", "9"
    Person_6_List.Add "62228", "10"
    Person_6_List.Add "62319", "11"
    Person_6_List.Add "64547", "12"
    Person_6_List.Add "64658", "13"
    Person_6_List.Add "79318", "14"
    Person_6_List.Add "81833", "15"
    Person_6_List.Add "82877", "16"
    Person_6_List.Add "83326", "17"
'**************************************************************

' Person_7's List of CAGE Codes
    Dim Person_7_List
    Set Person_7_List = CreateObject("Scripting.Dictionary")
    Person_7_List.Add "07187", "1"
    Person_7_List.Add "26512", "2"
    Person_7_List.Add "97942", "3"
    Person_7_List.Add "00816", "4"
    Person_7_List.Add "59211", "5"
    Person_7_List.Add "35351", "6"
    Person_7_List.Add "66126", "7"
'**************************************************************

' Person_8's List of CAGE Codes
    Dim Person_8_List
    Set Person_8_List = CreateObject("Scripting.Dictionary")
    Person_8_List.Add "19710", "1"
    Person_8_List.Add "64896", "2"
    Person_8_List.Add "64117", "3"
'**************************************************************

    For i = 2 To LastRow
        searchString = CStr(Cells(i, cc).Value)
        
        If Person_1_List.Exists(searchString) Then
        Cells(i, d) = "Person_1"
        
        ElseIf Person_2_List.Exists(searchString) Then
        Cells(i, d) = "Person_2"
            
        ElseIf Person_3_List.Exists(searchString) Then
        Cells(i, d) = "Person_3"
            
        ElseIf Person_4_List.Exists(searchString) Then
        Cells(i, d) = "Person_4"
            
        ElseIf Person_5_List.Exists(searchString) Then
        Cells(i, d) = "Person_5"
        
        ElseIf Person_6_List.Exists(searchString) Then
        Cells(i, d) = "Person_6"
        
        ElseIf Person_7_List.Exists(searchString) Then
        Cells(i, d) = "Person_7"
        
        ElseIf Person_8_List.Exists(searchString) Then
        Cells(i, d) = "Person_8"

        Else
        Cells(i, d) = ""
        
    End If
    Next i
End Sub
