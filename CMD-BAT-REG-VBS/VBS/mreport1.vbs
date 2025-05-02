Dim obArgs,cArgs,iSize,ndate,tmailbox

servername = "yourservername"

Set Cnxn1 = CreateObject("ADODB.Connection")
strCnxn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\mbsize.mdb;"
Cnxn1.Open strCnxn1
ndate = getdate1()

Set obArgs = WScript.Arguments
tmailbox = obArgs.Item(0)

set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
polQuery = "<LDAP://" & strNameingContext &  ">;(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy));distinguishedName,gatewayProxy;subtree"
Com.ActiveConnection = Conn
Com.CommandText = polQuery
Set plRs = Com.Execute
while not plRs.eof
	for each adrobj in plrs.fields("gatewayProxy").value
		if instr(adrobj,"SMTP:") then dpDefaultpolicy = right(adrobj,(len(adrobj)-instr(adrobj,"@")))
	next
	plrs.movenext
wend



adminvroot = "http://" & servername & "/exadmin/admin/" & dpDefaultpolicy 

Main

Sub Main()
   Dim sConnString,domainname

   rem On Error Resume Next
   sConnString = adminvroot & "/mbx/" & obArgs.Item(0) & "/NON_IPM_SUBTREE"
   ' wscript.echo sConnString

   iSize = 0

   RecurseFolder(sConnString)
   sqlstate1 =  "insert into dbo_mailboxusage ([date],Mailbox_Name,folder_path,folder_name,folder_size) values ('" & ndate & "','" & tmailbox & "','Totalsize','" & replace(tmailbox,"'","") & "','" & replace(formatnumber((iSize/1024/1024),2),",","") & "')"
   Cnxn1.Execute(sqlstate1)
   ' wscript.echo "Mailbox Size: " & iSize
End Sub

Public Sub RecurseFolder(sUrl)
   Dim oXMLHttp
   Dim oXMLDoc
   Dim oXMLSizeNodes
   Dim oXMLHREFNodes
   Dim oXMLHasSubsNodes
   Dim sQuery

   Set oXMLHttp = CreateObject("Microsoft.xmlhttp")
   If Err.Number <> 0 Then
      ' wscript.echo "Error Creating XML object"
      ' wscript.echo Err.Number & ": " & Err.Description
      Set oXMLHttp = Nothing
   End If

   ' Open DAV connection.
   oXMLHttp.open "SEARCH", sUrl, False, "",""
   If Err.Number <> 0 Then
      ' wscript.echo "Error opening DAV connection"
      ' wscript.echo Err.Number & ": " & Err.Description
      Set oXMLHttp = Nothing
   End If

   ' Set up query.
   sQuery = "<?xml version=""1.0""?>"
   sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
   sQuery = sQuery & "<g:sql>SELECT ""http://schemas.microsoft.com/"
   sQuery = sQuery & "mapi/proptag/x0e080003"", ""DAV:hassubs"", ""DAV:displayname"" FROM SCOPE "
   sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
   sQuery = sQuery & "WHERE ""DAV:isfolder"" = true"
   sQuery = sQuery & "</g:sql>"
   sQuery = sQuery & "</g:searchrequest>"

   ' Set request headers.
   oXMLHttp.setRequestHeader "Content-Type", "text/xml"
   oXMLHttp.setRequestHeader "Translate", "f"
   oXMLHttp.setRequestHeader "Depth", "0"
   oXMLHttp.setRequestHeader "Content-Length", "" & Len(sQuery)

   ' Send request.
   oXMLHttp.send sQuery
   If Err.Number <> 0 Then
      ' wscript.echo "Error Sending Query"
      ' wscript.echo Err.Number & ": " & Err.Description
      Set oXMLHttp = Nothing
   End If 

   ' Load XML.
   Set oXMLDoc = oXMLHttp.responseXML

   ' Get the XML nodes that contain the individual sizes.
   Set oXMLSizeNodes = oXMLDoc.getElementsByTagName("d:x0e080003")

   ' Get the XML nodes that contain the individual HREFs.
   Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")

   ' Get the XML nodes that contain the individual HasSubs.
   Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")

   Set oXMLDispNodes = oXMLDoc.getElementsByTagName("a:displayname")

   ' Loop through the nodes, and then add all of the sizes.
   For i = 0 to oXMLSizeNodes.length - 1
      ' wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
      ' wscript.echo "Size: " & oXMLSizeNodes.Item(i).nodeTypedValue
      iSum = iSum + oXMLSizeNodes.Item(i).nodeTypedValue
      iSize = iSum 
      foldersize = oXMLSizeNodes.Item(i).nodeTypedValue
      workfolderfp = oXMLHREFNodes.Item(i).nodeTypedValue 
      workfolder = oXMLDispNodes.Item(i).Text
      ' wscript.echo workfolder 
      sqlstate1 = "insert into dbo_mailboxusage ([date],Mailbox_Name,folder_path,folder_name,folder_size) values ('" & ndate & "','" & tmailbox & "','" & replace(workfolderfp,"'","") &"','" & replace(workfolder,"'","") & "','" & replace(formatnumber((foldersize/1024/1024),2),",","") & "')"
      Cnxn1.Execute(sqlstate1)	
      ' If the folder has subfolders, call your recursive function to 
      ' process subfolders.
      If oXMLHasSubsNodes.Item(i).nodeTypedValue = True Then
         RecurseFolder oXMLHREFNodes.Item(i).nodeTypedValue
      End If
   Next

   ' Clean up.
   Set oXMLSizeNodes = Nothing
   Set oXMLDoc = Nothing
   Set oXMLHttp = Nothing
End Sub

function getdate1()
cdate1 = now()
if month(cdate1) < 10 then 
	if day(cdate1) < 10 then
		qdat = year(cdate1) & "0" & month(cdate1) & "0" & day(cdate1)
	else
		qdat = year(cdate1) & "0" & month(cdate1) & day(cdate1)
	end if 
else
	if day(cdate1) < 10 then
		qdat = year(cdate1) & month(cdate1) & "0" & day(cdate1)
	else
		qdat = year(cdate1) & month(cdate1) & day(cdate1)
	end if 
end if
getdate1 = qdat
end function