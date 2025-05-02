'-------------------------------------------------------------------------
' Start of MAIN code (checkservers.vbs)
'-------------------------------------------------------------------------
'
' VBScript: checkservers.vbs
' Author: Jeff Mason aka TNJMAN aka bitdoctor
' 09/06/2013
'
' Assumptions/Notes:
' 1) Create Excel document (c:\scripts\servers.xlsx) with ONE worksheet, 
'    containing only "server name" in Column 1
' 2) Save this code as c:\scripts\checkservers.vbs
' 3) Change directory to your scripts folder, then run like this: 
'    cscript checkservers.vbs
' 4) Assume initial sheet has NO description in Column 2
' 5) IF the value in Column 2 (Description) is NOT empty, skip to next row
' 6) Assume user will customize/choose which 'strRoot' (FULL domain or just 
'    specific OU) that he/she wants to run against
' 7) Assume user will change excelPath to meet his/her environment
' 8) To enhance further, I would add an "updatecount" to display how many
'    Excel rows actually got updated; but it may not be needed.
'
' CREDIT goes to Mr. Greg Hatcher for the base code for reading an Excel sheet
' http://www.gregthatcher.com/Papers/VBScript/ExcelExtractScript.aspx
' Assumptions:
' You must Set excelPath = "C:\scripts\servers.xlsx" (wherever your xlsx file is)
' You must have at least "read" permissions to AD/LDAP
'
Option Explicit
Dim objExcel
Dim excelPath
Dim worksheetCount
Dim counter				' To count rows and/or columns
Dim currentWorkSheet
Dim usedColumnsCount, ColumnsCount	' Variable for the # of columns
Dim usedRowsCount			' Variables for the # of rows
Dim row
Dim column 
Dim top					' Variables for cells
Dim left
Dim Cells
Dim curCol				' Current row & column of current worksheet
Dim curRow
Dim server				' Values in the current row and column
Dim descr

Dim svr, svrcmp, svrflag, objRootDSE, objCn, objCmd, objRes
Dim strRoot, strfilter, strAttributes, strScope, strTmp, strDescription

excelPath = "C:\scripts\servers.xlsx"

WScript.Echo "Reading Data from Path/File: " & excelPath

Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth

' open read/write 
objExcel.Workbooks.open excelPath

' One worksheet, with Server Name in Col 1 and Description (ultimately) in Col 2
workSheetCount = 1
ColumnsCount = 1 'Uses only the leftmost Column to "read" the Sheet, then plugs, Column 2 with data

WScript.Echo "-------------------------------------------------------"
WScript.Echo "Reading data from worksheet " & workSheetCount
WScript.Echo "-------------------------------------------------------"  & vbCRLF

Set currentWorkSheet = objExcel.ActiveWorkbook.Worksheets(workSheetCount)
' how many columns are used in the current worksheet
usedColumnsCount = ColumnsCount
' how many rows are used in the current worksheet
usedRowsCount = currentWorkSheet.UsedRange.Rows.Count
' What is the topmost row in the spreadsheet that has data in it
top = currentWorksheet.UsedRange.Row
' What is the leftmost column in the spreadsheet that has data in it
left = currentWorksheet.UsedRange.Column
Set Cells = currentWorksheet.Cells

'-----------------------------------------------------------------------------
' Row Loop - Loop through each row in the worksheet
' 
' Only deal with Cols 1 & 2 of Sheet1, since SERVER=Col1 and DESCRIPTION=Col2
' Column 2 is built by "checksvr" subroutine, based on Column 1)
'
  For row = 0 to (usedRowsCount-1)
  ' only look at rows/cols in the "used" range
    curRow = row+top
'   curCol = column+left
    curCol = 1
'   get the (server name) that is in each row of Column 1 
    server = Cells(curRow,curCol).Value
    strDescription = Cells(curRow,curCol+1).Value
    If IsEmpty(strDescription) Then ' If Col 2 already populated, skip to next row in sheet
      If Not (IsEmpty(server)) Then
       'Uncomment for debug: WScript.Echo (server)
       'Call the checksvr subroutine (server) to find if server is in AD
        Checksvr(server)
       'Uncomment for debug: Wscript.Echo strDescription
        Cells(curRow,curCol+1).Value = strDescription
       'Populate excel sheet description column , save below when done
      End If
    End If
  Next
'
' End Row loop
'-----------------------------------------------------------------------------
' Done with the current worksheet, release the memory
Set currentWorkSheet = Nothing

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit
WScript.Echo "Finished."

Set currentWorkSheet = Nothing
' Finished with Excel object, release it from memory & get out !!!
Set objExcel = Nothing
WScript.Quit(0)
'-------------------------------------------------------------------------
' End of MAIN code 
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
' Subroutine (checksvr) to check for the sever name in Active Directory
'-------------------------------------------------------------------------
'
' CREDIT to Mr. Gregory Shiro for the TechNet script I turned into a subroutine
' http://tinyurl.com/ljwjfwe

Sub checksvr(svr)

On Error Resume Next

' Point to the domain/ldap root
Set objRootDSE = GetObject("LDAP://RootDSE")

' Query all Active Directory
' strRoot = objRootDSE.Get("DefaultNamingContext") 'Uncomment this if you want to search ENTIRE AD TREE

' Query a specific Organizational Unit
strRoot = "OU=Servers,DC=YOUR-DOMAIN,DC=com" ' Comment this out, if searching ALL OF AD
strfilter = "(&(objectCategory=Computer)(objectClass=Computer))"
strAttributes = "sAMAccountName,description"
strScope = "subtree"
Set objCn = CreateObject("ADODB.Connection")
Set objCmd = CreateObject("ADODB.Command")

objCn.Provider = "ADsDSOObject"
objCn.Open "Active Directory Provider"
objCmd.ActiveConnection = objCn
objCmd.Properties("Page Size") = 1000
' Filter the query for only sAMAccountName,description of any computers in AD
objCmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & strAttributes & ";" & strScope
Set objRes = objCmd.Execute

svrcmp = UCase(svr) & "$" 'Upper-case the Server entry from the spreadsheet for consistent compare
svrflag = "" 'Clear out the "found-server" flag

Do While Not objRes.EOF
' If description is blank/null, set the value to the word "BLANK"
    strDescription = ""
    If Not (IsNUll(objRes.Fields("description").Value)) Then
        For Each strTmp in objRes.Fields("description").Value
            strDescription = strTmp
        Next
    Else
       strDescription = "BLANK"
    End If

    ' We want to check ALL descriptions, including null descriptions
    ' But only for the server passed into this script as an argument

    If svrcmp = objRes.Fields("sAMAccountName").Value Then

     'If Excel server name found in AD, set svrflag = "TRUE" & end the subroutine
      svrflag = "TRUE"
     'Write this to the Excel spreadsheet / exit the subroutine
      Exit Sub

    End If 
   'Move to / read the next AD resource record
    objRes.MoveNext
Loop
   'If flag never set to "TRUE" then fall out through here - server not found in AD
    strDescription = "NOT FOUND IN AD"
objRes.close
ObjCn.close
'
'-------------------------------------------------------------------------
 End Sub
'-------------------------------------------------------------------------
