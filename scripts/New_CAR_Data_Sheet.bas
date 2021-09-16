Sub New_CAR_Data_Sheet()
'
' New_CAR_Data_Sheet Macro
'
' Keyboard Shortcut: Ctrl+n

' This selects column CZ and formats it to a number with no decimals
' and autofits the column
    Columns("CZ:CZ").Select
    Selection.NumberFormat = "0"
    Columns("CZ:CZ").EntireColumn.AutoFit

' This selects a cell that is away from normal visible range and determines
' the last filled cell in column A. This value is then pasted in the cell below it
' so an actual value can be used in the rest of the macro.
    Range("CZ2").Select
    ActiveCell.Formula = "=INDEX(A:A,COUNTA(A:A),1)"
    Range("CZ2").Select
    Selection.Copy
    Range("CZ3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
' The hidden Template sheet is revealed copied and pasted into the workbook and
' then hidden again
    Sheets("Template").Visible = True
    Sheets("Template").Select
    Application.CutCopyMode = False
    Sheets("Template").Copy After:=Sheets(2)
    Sheets("Template").Visible = False

' The name of the most recent sheet is copied from the Summary sheet and CAR # is filled
    Sheets("Summary").Select
    Range("CZ3").Select
    Selection.Copy
    Sheets("Summary").Select
    Range("A1").Select
    Sheets("Template (2)").Select
    Range("C1:AA1").Select
    ActiveSheet.Paste
' The current date is placed into the meeting date
    Range("AD6").Select
    ActiveCell.Formula = "=TODAY()"
    Range("AD6").Select
    Selection.Copy
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AD6").Select
    Application.CutCopyMode = False
    Selection.ClearContents
' The sheet is renamed to the CAR#
    Sheets("Template (2)").Select
    Sheets("Template (2)").Name = Range("C1").Value
    Range("B2").Select
    
' This with statement determines the last filled row number on the Summary sheet and
' selects the current sheet name from CZ3
    Dim LastRow As Long
    Sheets("Summary").Select
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    Range("CZ4").Value = LastRow
    SheetName = Range("CZ3").Text
    
' The last value in the respective columns on the summary page are found from the respective
' CAR data sheet
    Cells(LastRow, 2).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!B6:B50,1)"
    Cells(LastRow, 3).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!G6:G50,1)"
    Cells(LastRow, 4).Select
    ActiveCell.Formula = "=IF(MIN(" & SheetName & "!G6:G50)=0,""Not Received"",MIN(" & SheetName & "!G6:G50))"
    Cells(LastRow, 5).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!G6:G50,1)"
    Cells(LastRow, 6).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!H6:H50,1)"
    Cells(LastRow, 7).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!I6:I50,1)"
    Cells(LastRow, 8).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!J6:J50,1)"
    Cells(LastRow, 9).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!K6:K50,1)"
    Cells(LastRow, 10).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!L6:L50,1)"
    Cells(LastRow, 11).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!Q6:Q50,1)"
    Cells(LastRow, 12).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!R6:R50,1)"
    Cells(LastRow, 13).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!S6:S50,1)"
    Cells(LastRow, 14).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!T6:T50,1)"
    Cells(LastRow, 15).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!W6:W50,1)"
    Cells(LastRow, 16).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!X6:X50,1)"
    Cells(LastRow, 17).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!Y6:Y50,1)"
    Cells(LastRow, 18).Select
    ActiveCell.Formula = "=VLOOKUP(9.99999999999999E+307," & SheetName & "!Z6:Z50,1)"
    Cells(LastRow, 19).Select
    ActiveCell.Formula = "=IFERROR(VLOOKUP(9.99999999999999E+307," & SheetName & "!AA6:AA50,1),""Open"")"
    
' The last name of the POC is added
    Cells(LastRow, 20).Select
    ActiveCell.Formula = "=TRIM(RIGHT(SUBSTITUTE(" & SheetName & "!B2,"" "",REPT("" "",255)),255))"
    
' The group for the tag is assigned
    Cells(LastRow, 21).Select
    ActiveCell.Formula = "=IFERROR(IF(FIND(""ADE""," & SheetName & "!B3),""ADE"", ""ADQ""),""ADQ"")"
    
' This calculates the data required for the CAR Charts
    Cells(LastRow, 27).Select
    ActiveCell.Formula = "=IF(S" & LastRow & "= ""Open"",TODAY()-B" & LastRow & "," & 0 & ")"
    Cells(LastRow, 37).Select
    ActiveCell.Formula = "=IF(S" & LastRow & "=""Open"",""0"",S" & LastRow & "-B" & LastRow & ")"
    Cells(LastRow, 47).Select
    ActiveCell.Formula = "=IFERROR(E" & LastRow & "-B" & LastRow & ",""Not Received"")"
    Cells(LastRow, 48).Select
    ActiveCell.Formula = "=IFERROR(K" & LastRow & "-E" & LastRow & ",""Not Received"")"
    Cells(LastRow, 49).Select
    ActiveCell.Formula = "=IFERROR(S" & LastRow & "-B" & LastRow & ",""Not Received"")"
    Cells(LastRow, 50).Select
    ActiveCell.Formula = "=IF(S" & LastRow & "=""Open"",TODAY()-B" & LastRow & ",""Closed"")"
    
' The CAR # on the summary page is turned into a hyperlink for the appropriate CAR data sheet
    SheetName = Range("CZ3").Text
    LastRow = Range("CZ4").Value
    Cells(LastRow, 1).Select
    ActiveCell.FormulaR1C1 = "=HYPERLINK(""[CAR Tracker Meeting.xlsm]" & SheetName & "!B4"",""" & SheetName & """)"
    
' This sets up conditional color formatting once the CAR has a Closure Date
    Range("A" & LastRow & ":U" & LastRow & "").Select
    Range("S" & LastRow & "").Activate
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$S" & LastRow & "<>""Open"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
' Selects active CAR Issue Date
    Sheets("" & SheetName & "").Select
    Range("B6").Select
End Sub
