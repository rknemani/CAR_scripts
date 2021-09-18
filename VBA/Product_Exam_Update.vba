Sub Product_Exam_Update()
'
' Product_Exam_Update Macro
'
' Keyboard Shortcut: Ctrl+p
    
' This selects the previous month deletes the column and creates the new
' 12 month average
    Range("B1:B4").Select
    Selection.ClearContents
    Range("C1:N4").Select
    Selection.Copy
    Range("B1").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
        IconFileName:=False
    Range("B1").Select
    Selection.Copy
    Range("N1").Select
    ActiveSheet.Paste
    Range("O11").Select

' This selects the previous month deletes the column and creates the new
' month's PE by Program
    Range("B27:B37").Select
    Selection.ClearContents
    Range("C27:D37").Select
    Selection.Copy
    Range("B27").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
        IconFileName:=False
    Range("B1").Select
    Selection.Copy
    Range("D27").Select
    ActiveSheet.Paste
    Range("E27").Select

' Clears the previous month's data
    Sheets("PE Log").Select
    Range("Table_owssvr").Select
    Selection.ClearContents

' Turns off the paste warnings
    Application.DisplayAlerts = False
    
'This copies the PE Log into the Monthly Report file
    Workbooks.Open ("C:\Users\Documents\PE Monthly Report\Monthly Charts 2015\March 2015 Data.xlsx")
    Windows("March 2015 Data.xlsx").Activate
    Sheets(1).Select
    Range("Table_owssvr").Select
    Selection.Copy
    Windows("Product Exam Template(2).xlsm").Activate
    Sheets("PE Log").Select
    Range("Table_owssvr").Select
    ActiveSheet.Paste
    Workbooks.Open("C:\Users\Documents\PE Monthly Report\Monthly Charts 2015\March 2015 Data.xlsx").Close

' This fixes the the Total Observations Made column
    Range("AA2").Select
    ActiveCell.Formula = _
        "=COUNTA(Table_owssvr[@[Bonding]:[Protective Finish Coverage]])"
    Range("AA2").Select
    Selection.AutoFill Destination:=Range("Table_owssvr[Total Observations Made]" _
        ), Type:=xlFillDefault
    
' Selects a cell in the Summary Sheet
    Sheets("Product Exams").Select
    Range("P1").Select
    Application.DisplayAlerts = True

End Sub
