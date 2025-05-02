Dim obArgs,cArgs,iSize,ndate,tmailbox


Set obArgs = WScript.Arguments
tmailbox = obArgs.Item(0)

Main

Sub Main()
   Dim sConnString,domainname

   rem On Error Resume Next
   domainname = "yourdomain.com" ' Use the Primary SMTP Exchange Domain name
  ' Set up connection string to mailbox.
   sConnString = "file://./backofficestorage/" & domainname
   sConnString = sConnString & "/mbx/" & obArgs.Item(0) & ""


   iSize = 0

   RecurseFolder(sConnString)
End Sub

Public Sub RecurseFolder(sConnString)

   Dim oConn
   Dim oRecSet
   Dim sSQL

   ' Set up SQL SELECT statement.
   sSQL = "SELECT ""http://schemas.microsoft.com/mapi/proptag/x0e080003"", "
   sSQL = sSQL & """DAV:href"",""DAV:hassubs"",""DAV:displayname"" "
   sSQL = sSQL & "FROM SCOPE ('SHALLOW TRAVERSAL OF """ & sConnString
   sSQL = sSQL & """') WHERE ""DAV:isfolder"" = true"


   ' Create Connection object.
   Set oConn = CreateObject("ADODB.Connection")
   if Err.Number <> 0 then
      WScript.Echo "Error creating ADO Connection object: " & Err.Number & " " & Err.Description
   end if

   ' Create RecordSet object.
   Set oRecSet = CreateObject("ADODB.Recordset")
   if Err.Number <> 0 then
      WScript.Echo "Error creating ADO RecordSet object: " & Err.Number & " " & Err.Description
      Set oConn = Nothing
      Exit Sub
   end if

   ' Set provider to EXOLEDB.
   oConn.Provider = "Exoledb.DataSource"

   ' Open connection to folder.
 wscript.echo sConnString
   oConn.Open sConnString
   if Err.Number <> 0 then
      WScript.Echo "Error opening connection: " & Err.Number & " " & Err.Description
      Set oRecSet = Nothing
      Set oConn = Nothing
      Exit Sub
   end if

   ' Open Recordset of all subfolders in folder.
   oRecSet.CursorLocation = 3
   oRecSet.Open sSQL, oConn.ConnectionString
   if Err.Number <> 0 then
      WScript.Echo "Error opening recordset: " & Err.Number & " " & Err.Description
      oRecSet.Close
      oConn.Close
      Set oRecSet = Nothing
      Set oConn = Nothing
      Exit Sub
   end if

   if oRecSet.RecordCount = 0 then
      oRecSet.Close
      oConn.Close
      Set oRecSet = Nothing
      Set oConn = Nothing
      Exit Sub
   end if

   ' Move to first record.
   oRecSet.MoveFirst
   if Err.Number <> 0 then
      WScript.Echo "Error moving to first record: " & Err.Number & " " & Err.Description
      oRecSet.Close
      oConn.Close
      Set oRecSet = Nothing
      Set oConn = Nothing
      Exit Sub
   end if

   ' Loop through all of the records, and then add the size of the 
   ' subfolders to obtain the total size.
   While oRecSet.EOF <> True
      ' Increment size.
      iSize = iSize + oRecSet.Fields.Item("http://schemas.microsoft.com/mapi/proptag/x0e080003")
      workfolderfp = oRecSet.Fields("DAV:href").value
      wscript.echo workfolderfp
      archivemail(workfolderfp)      
      ' If the folder has subfolders, recursively call RecurseFolder to process them.

      If oRecSet.Fields.Item("DAV:hassubs") = True then
         RecurseFolder oRecSet.Fields.Item("DAV:href")
      End If
      ' Move to next record.
      oRecSet.MoveNext
      if Err.Number <> 0 then
         WScript.Echo "Error moving to next record: " & Err.Number & " " & Err.Description
         Set oRecSet = Nothing
         Set oConn = Nothing
         Exit Sub
      end if
   wend

   ' Close Recordset and Connection.
   oRecSet.Close
   if Err.Number <> 0 then
      WScript.Echo "Error closing recordset: " & Err.Number & " " & Err.Description
      Set oRecSet = Nothing
      Set oConn = Nothing
      Exit Sub
   end if

   oConn.Close
   if Err.Number <> 0 then
      WScript.Echo "Error closing connection: " & Err.Number & " " & Err.Description
      Set oRecSet = Nothing
      Set oConn = Nothing
      Exit Sub
   end if

   ' Clean up memory.
   Set oRecSet = Nothing
   Set oConn = Nothing
   wscript.echo sConnString
End Sub


Sub archivemail(workfolderfp)  

mailboxurl = workfolderfp
set Rec = CreateObject("ADODB.Record")
set Rs = CreateObject("ADODB.Recordset")
Set Conn = CreateObject("ADODB.Connection")
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open mailboxurl, ,3
SSql = "SELECT ""DAV:href"", ""DAV:contentclass"" FROM scope('shallow traversal of """ & mailboxurl & """') " 
SSql = SSql & " WHERE (""urn:schemas:httpmail:datereceived"" < CAST(""" & isodateit(now()-31) & """ as 'dateTime')) AND ""DAV:isfolder"" = false"                 
SSql = SSql & " AND ""DAV:contentclass"" = 'urn:content-classes:message'"
Rs.CursorLocation = 2 'adUseServer = 2, adUseClient = 3
rs.open SSql, rec.ActiveConnection, 3
while not rs.eof
	rs.delete 1
	rs.movenext
wend
rs.close

end sub


function isodateit(datetocon)
	strDateTime = year(datetocon) & "-"
	if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Month(datetocon) & "-"
	if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Day(datetocon) & "T" & formatdatetime(datetocon,4) & ":00Z"
	isodateit = strDateTime
end function