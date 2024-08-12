'Description: This VBScript takes the RAW Data Un-formated (Tab Delimited) xls files exported from FBMS and BASIS+ in to Formated TAB Delimited .txt files.
'   This script loops through an array of report names and converts the RAW Files to txt file
'   for each file in the array the script also runs sql scripts that create a SQL table to accept the data
'   after the table is created this script runs a corresponding SQL script to load the data from the converted txt in the the table for that report
'   results from the scripts are stored in a log files and are overwritten every try the script is run.
'   this script has commented code for CSV however the RAW data files have columns with ',' in them causing SQL bulk load issues
'PRE: The machines where this script is run is required to have MS EXCEL Installed
'Some custimization must be completed in the "Modify" section prior to running script
'   SQL scripts will need to be modified to point to the source report txt file
'rptList - below can be modified by adding/removing the file names of the RAW Data files
'rptSrc - is the Source location of the RAW Data files
'rptDest - is the location where the converted txt files will be saved
'   tblPath - is the location of the SQL scripts that create the tables for the data to be loaded into
'   dataPath - is the location of the SQL scripts to load the data from the converted txt files into the tables created
'   logPath - is the location where the logs for creating tables and loading data will be saved
'   SQLSVR - is the server name and SQL instance runnging SQL
'   DB - is name of the Database where tables are created
'POST: The resulting file from the source Tab Deliimited file is a correctly formated txt file that can be imported into SQL and named the same as Source File with the .txt extention
'   Table for each report will have been created on SQL Server
'   Each table will have data from converted txt Files loaded into tables
'   log files for Table creation will be empty on a clean run
'   log files for data loading will display number of rows loaded on a clean runs
'Contact: Ross Wickman - Ross@tactful.cloud
'Modify Date: 23Nov15

'******************** CONFIGURE VARIABLES : NO NEED TO MODIFY ********************
'On Error Resume Next
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Dim rptList, rpt, rptSrc, rptDest
Dim tblPath, dataPath, logPath, SQLSVR
Set objShell = CreateObject("WScript.Shell")



'******************** !! MODIFY !! - CHANGE THESE VARIABLES TO ADD/REMOVE REPORTs OR CHANGE FILE PATHS ********************

rptList = Array("530B","565A")

rptSrc = "\\servername\Reports\"
rptDest = "\\servername\Reports Test\"

SQLSVR = "servername\instance"
tblPath = "\\path"
dataPath = "\\path"
logPath = "\\path"


'******************** END MODIFICATION AREA ********************


'*******************************************************
'***************** DO NOT MODIFY BELOW *****************
'*******************************************************


'* UPDATE DATABASE WITH RUNTIME INFORMATION
connstr = "Provider=SQLOLEDB.1;Data Source=" & SQLSVR & ";Initial Catalog=" & DB & ";Trusted_Connection=Yes;"
Set myConn = CreateObject("ADODB.Connection")
Set myCommand = CreateObject("ADODB.Command" )
Set rs = CreateObject("ADODB.Recordset")
myConn.Open connstr
Set myCommand.ActiveConnection = myConn
dim lastMod, nextRun, reqTime, fmtDT

rs.Open "SELECT ModifiedBy as modBy FROM Settings WHERE Name = 'RunScheduled'", myConn
lastMod = rs("modBy")
rs.close
rs.Open "SELECT Value as time FROM Settings WHERE Name = 'RunScheduled'", myConn
reqTime = rs("time")
rs.close

IF Time < TimeValue("09:55:59 AM") THEN
    nextRun = DateAdd("d",0,Date()) & " 09:55:59 AM"
ELSE 
    nextRun = DateAdd("d",1,Date()) & " 09:55:59 AM"
END IF

fmtDT = FormatDateTime(reqTime)

IF DateDiff("s",Now, fmtDT) < 0 THEN
'IF TRUE THEN

    myCommand.CommandText = "UPDATE Settings SET Value = '" & Now & "', ModifiedBy = '" & lastMod & "' WHERE Name = 'ReportRunTime'"
    myCommand.Execute
    myCommand.CommandText = "UPDATE Settings SET Value = '" & nextRun & "', ModifiedBy = 'System' WHERE Name = 'RunScheduled'"
    myCommand.Execute

    '* CONVERT REPORTS
    FOR EACH rpt IN rptList

    oExcel.DisplayAlerts = False
    Set oBook = oExcel.Workbooks.Open("" & rptSrc & rpt & ".xls")
    'oBook.SaveAs "" & rptDest & rpt & ".csv", 6 'To COMMA Delimited CSV
        oBook.SaveAs "" & rptDest & rpt & ".txt", -4158 'To TAB Delimited TXT

    oBook.Close False
        oExcel.Quit

    Set objScriptExec = objShell.Exec("SQLCMD -S " & SQLSVR & " -E -i " & tblPath & rpt & "_Table.sql -o " & logPath & rpt & "_Create.log")
    WScript.Sleep 1000
        Set objScriptExec = objShell.Exec("SQLCMD -S " & SQLSVR & " -E -i " & dataPath & rpt & "_Data.sql -o " & logPath & rpt & "_Load.log")

    NEXT

END IF

'* CLEAN UP
myConn.Close
Set myConn = nothing
Set myCommand = nothing
Set rs = nothing