Dim obArgs 
Dim cArgs 
Dim iSize 


Set obArgs = WScript.Arguments 
cArgs = obArgs.Count 


Main 


Sub Main() 
   Dim sConnString 

   If cArgs <> 1 Then 
      WScript.Echo "Usage: cscript  DOMAINNAME" 
      Exit Sub 
   End If 


   ' Set up connection string to mailbox. 
   sConnString = "file://./backofficestorage/" & obArgs.Item(0) 
   sConnString = sConnString & "/public folders" 
   iSize = 0 
   RecurseFolder(sConnString) 
End Sub 


Public Sub RecurseFolder(sConnString) 
   Dim oConn 
   Dim oRecSet 
   Dim sSQL 


   ' Set up SQL SELECT statement. 
   sSQL = "SELECT ""http://schemas.microsoft.com/mapi/proptag/x0e080003"", " 
   sSQL = sSQL & """DAV:href"", ""DAV:contentclass"", " 
   sSQL = sSQL & """DAV:hassubs"", " 
   sSQL = sSQL & """http://schemas.microsoft.com/exchange/publicfolderemailaddress"" " 
   sSQL = sSQL & "FROM SCOPE ('SHALLOW TRAVERSAL OF """ & sConnString 
   sSQL = sSQL & """') WHERE ""DAV:isfolder"" = true" 


   ' Create Connection object. 
   Set oConn = CreateObject("ADODB.Connection") 
   if Err.Number <> 0 then 
      WScript.Echo "Error creating ADO Connection object: " & Err.Number & "" & Err.Description 
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
      if isnull(oRecSet.Fields.Item("http://schemas.microsoft.com/exchange/publicfolderemailaddress")) then 
	    wscript.echo "Folder Not Mail Enabled" & oRecSet.Fields.Item("DAV:href") & " " &  oRecSet.Fields.Item("DAV:contentclass")
            if oRecSet.Fields.Item("DAV:contentclass") = "urn:content-classes:mailfolder" then call Mailenable(oRecSet.Fields.Item("DAV:href")) 
      else
	    wscript.echo oRecSet.Fields.Item("http://schemas.microsoft.com/exchange/publicfolderemailaddress") 
     end if 

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
End Sub 

sub Mailenable(FolderURL)

set objFolder = createobject("CDO.Folder")

objFolder.DataSource.Open FolderURL, , 3, -1
Set objRecip = objFolder.Getinterface("IMailRecipient")
objRecip.MailEnable
objFolder.DataSource.Save
set objFolder = nothing
set objRecip = nothing
wscript.echo "Mail Enabled Folder - " & FolderURL

end sub
