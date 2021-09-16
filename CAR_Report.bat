@echo off

:: start /d "" IEXPLORE.EXE https://url/applications.form
:: start /d "" IEXPLORE.EXE https://url/app/workload_browse.cfm
:: timeout /t 10

start /d "" IEXPLORE.EXE https://url/login.do?BusinessArea=CAR
timeout /t 10
start /d "" IEXPLORE.EXE javascript: openFullWindow('../home/launchReport.do?id=11111&reportUrl=https://url/cgi-bin/cognos.cgi&paramUrl=b_action=cognosViewer^ui.action=run^p_EmployeeID=XXXXXX^p_District=P7^p_Dodaac=S2606A^p_Cust=null^p_Dodaac2=DUMMY^run.prompt=true^ui.object=/content/package[@name=\'CAR\']/report[@name=\'Quality Assurance CAR Report\']^run.outputFormat=PDF', 'XXXXX')