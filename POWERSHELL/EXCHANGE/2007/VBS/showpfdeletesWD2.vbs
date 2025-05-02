days = wscript.arguments(1)
servername = wscript.arguments(0)
SB = 0
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\PfDeletes.csv",2,true)
wfile.writeline("DateDeleted,FolderName,FolderID,DeletedBy-MailboxName,DeletedBy-UserName,FolderSize,NumberofItems")
dtmStartDate = CDate(Date) - days
dtmStartDate = Year(dtmStartDate) & Right( "00" & Month(dtmStartDate), 2) & Right( "00" & Day(dtmStartDate), 2)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '9682' and TimeWritten >= '" & dtmStartDate & "' ",,48)
wscript.echo "Date Deleted		FolderName	FolderID	Deleter-Mbox		Deleter-UserName"
For Each objEvent in colLoggedEvents
    SB = 1
    Time_Written = cdate(DateSerial(Left(objEvent.TimeWritten, 4), Mid(objEvent.TimeWritten, 5, 2), Mid(objEvent.TimeWritten, 7, 2)) & " " & timeserial(Mid(objEvent.TimeWritten, 9, 2),Mid(objEvent.TimeWritten, 11, 2),Mid(objEvent.TimeWritten,13, 2)))
    FolderName = Mid(objEvent.Message,8,instr(objEvent.Message,"with folder")-9)
    FolderID = Mid(objEvent.Message,instr(objEvent.Message,"with folder ID")+15,(instr(objEvent.Message,"was deleted by")-(instr(objEvent.Message,"with folder ID")+16)))
    MailboxName = Mid(objEvent.Message,instr(objEvent.Message,"was deleted by")+15,instr(objEvent.Message,", user account")-(instr(objEvent.Message,"was deleted by")+15))
    UserName = Mid(objEvent.Message,instr(objEvent.Message,", user account")+15,(instr(objEvent.Message,chr(13))-(instr(objEvent.Message,", user account")+16)))  
    retarry =  showdeletails(servername,FolderName)
    wscript.echo Time_Written & "	" & FolderName & "	" & FolderID & "		" & MailboxName & "		" & UserName & "	" & retarry(0) &  "	" & retarry(1)
    wfile.writeline(Time_Written & "," & FolderName & "," & FolderID & "," & MailboxName & "," & UserName & "," & retarry(0) & "," & retarry(1))
next
wfile.close
set wfile = nothing
if SB = 0 then queryeventlog = "No Public Folder Deletes recorded in the last " & days & "Days"

function showdeletails(servername,pfpath)
dim retarry(2)
retarry(0) = " "
retarry(1) = " "
if instr(2,pfpath,"/") then
	lastslash = ""
	for i = 2 to len(pfpath)
		if mid(pfpath,i,1) = "/" then
			lastslash = i
		end if
		
	next
	rfolder = "http://" & servername &  "/public" & left(pfpath,int(lastslash))
else
	rfolder = "http://" & servername &  "/public/"
end if

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"", ""http://schemas.microsoft.com/mapi/proptag/x669B0014"", "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/x66400003"", ""http://schemas.microsoft.com/exchange/permanenturl"" FROM scope('SOFTDELETED traversal of """
strQuery = strQuery & rfolder & """') Where ""DAV:isfolder"" = True and "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/x6707001E"" = '" & pfpath & "'</D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", rfolder, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("d:x669B0014")						
   set oNodeList1 = oResponseDoc.getElementsByTagName("d:x66400003")
   set oNodeList2 = oResponseDoc.getElementsByTagName("a:href")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	set oNode1 = oNodeList1.nextNode
	set oNode2 = oNodeList2.nextNode
	permurl = oNode2.text
	if instr(2,pfpath,"/") then
		rfold = left(pfpath,int(lastslash)-1)
	else
		rfold = ""
	end if
	wscript.echo permurl
	if oNode.Text <> 0 then
		retarry(0) = formatnumber(oNode.text / 1024 /1024)
		retarry(1) = oNode1.text
	else
		retarry(0) = oNode.text
		retarry(1) = oNode1.text
	end if
   Next	
Else

End If
showdeletails = retarry
end function

