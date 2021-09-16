@echo off

:: These lines copy all the relevant QDATA sheets to my personal folder where I keep the XLSX file

xcopy "P:\file_location\Q-Data_1.xlsx" ^
"C:\Users\Name\Documents\PE_Monthly_Report" /Y

xcopy "P:\file_location\Q-Data_2.xlsx" ^
"C:\Users\Name\Documents\PE_Monthly_Report" /Y

xcopy "P:\file_location\Q-Data_3.xlsx" ^
"C:\Users\Name\Documents\PE_Monthly_Report" /Y

