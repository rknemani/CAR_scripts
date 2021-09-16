Sub PE_MonthlySheetCopy()
'
' PE_MonthlySheetCopy Macro
'
' Keyboard Shortcut: Ctrl+p
'
' This macro opens up each Q Data sheet and copies the PE Log to the PE Monthly Excel Sheet

' Turns off the paste warnings
    Application.DisplayAlerts = False
    
    Workbooks.Open ("C:\Users\Documents\PE Monthly Report\PE Log.xlsx")
    Windows("PE Log.xlsx").Activate
    Sheets(1).Select
    Cells.Select
    Selection.Copy
    Windows("Formulas 2014 PE totals.xlsm").Activate
    Sheets("PE Log").Select
    Cells.Select
    ActiveSheet.Paste
    Workbooks.Open("C:\Users\Documents\PE Monthly Report\PE Log.xlsx").Close
    
' Selects a cell in the Summary Sheet
    Sheets("Sheet1").Select
    Range("Z62").Select
    
    Application.DisplayAlerts = True
    
End Sub
