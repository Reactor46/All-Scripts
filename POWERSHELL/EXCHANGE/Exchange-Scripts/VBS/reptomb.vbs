snServername = "server"
mnMailboxname = "mailbox"
domain = "domain"
strpassword = "password"
strusername =  domain & "\" & mnMailboxname

DestinURL = "https://" & snServername & "/exchange/" & mnMailboxname & "/contacts/"
SourceURL = "https://" & snServername & "/public/pubcontacts/"

szXml = "destination=https://" & snServername & "/exchange/&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
set req = createobject("microsoft.xmlhttp")
req.Open "post", "https://" & snServername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
	if instr(lcase(reqhedrarry(c)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
next

set objSession = CreateObject("MAPI.Session")
strProfile = snServername  & vbLf & mnMailboxname
objSession.Logon "",,, False,, True, strProfile
Set objInfoStore = objSession.GetInfoStore(objSession.Inbox.StoreID)
set CntFolder = objSession.GetDefaultFolder(5)
Set objpubstore = objSession.InfoStores("Public Folders")

pfPublicFolderID = getpfid()
colbblob = Collabblobget()

QueryMailbox(colbblob)

wscript.echo "Done"

sub QueryMailbox(colbblob)

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" xmlns:R=""http://schemas.microsoft.com/repl/""><R:repl><R:collblob>" & colbblob & "</R:collblob></R:repl>"
strQuery = strQuery & "<D:sql>SELECT ""DAV:href"", ""urn:schemas:httpmail:subject"", ""http://schemas.microsoft.com/mapi/proptag/x0fff0102"",""http://schemas.microsoft.com/repl/repl-uid"" "
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & SourceURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False "
strQuery = strQuery & "</D:sql></D:searchrequest>"
req.open "SEARCH", SourceURL, false
req.setrequestheader "Content-Type", "text/xml"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: An error occurred on the server."
ElseIf req.status = 207 Then
   wscript.echo "Status: " & req.status
   wscript.echo "Status text:  " & req.statustext
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("d:collblob")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
        colblob =  oNode.Text
	Collabblobset(colblob)
   Next
   set idNodeList = oResponseDoc.getElementsByTagName("f:x0fff0102")
   set replidNodeList = oResponseDoc.getElementsByTagName("d:repl-uid")
   set replchangeType = oResponseDoc.getElementsByTagName("d:changetype")
   for id = 0 To (idNodeList.length -1)
	set oNode1 = idNodeList.nextNode
	set oNode2 = replidNodeList.nextNode
	set oNode3 = replchangeType.nextNode
	select case oNode3.text
		case "new" call CopyContact(Octenttohex(oNode1.nodeTypedValue),oNode2.text)
		case "delete" wscript.echo oNode3.text
			      wscript.echo oNode2.text
			      DeleteContact(oNode2.text)
		case "change" Wscript.echo "Change"
			      call DeleteContact(oNode2.text)
			      call CopyContact(Octenttohex(oNode1.nodeTypedValue),oNode2.text)
	end select
   next
Else
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: " & req.statustext
   wscript.echo "Response text: " & req.responsetext
End If

End Sub

function Collabblobget()

xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:cp='" & SourceURL & "'><a:prop><cp:collblob/></a:prop></a:propfind>"
req.open "PROPFIND", DestinURL, false, "", ""
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Depth", "0"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate", "f"
req.send xmlreqtxt
set oResponseDoc = req.responseXML
set oCobNode = oResponseDoc.getElementsByTagName("d:collblob")
For i1 = 0 To (oCobNode.length -1)
   set oNode = oCobNode.nextNode
   Collabblobget = oNode.Text   
Next

End function

Sub Collabblobset(colblob)
xmlstr = "<?xml version=""1.0""?>" _
& "<g:propertyupdate " _
& "    xmlns:g=""DAV:"" xmlns:e=""http://schemas.microsoft.com/exchange/""" _ 
& "    xmlns:dt=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"" " _
& "    xmlns:cp=""" & SourceURL & """ " _
& "    xmlns:header=""urn:schemas:mailheader:"" " _
& "    xmlns:mail=""urn:schemas:httpmail:"">  " _
& "    <g:set>  " _
& "        <g:prop>  " _
& "            <cp:collblob>" & colblob & "</cp:collblob>  " _
& "        </g:prop>  " _
& "    </g:set>  " _ 
& "</g:propertyupdate>" 

req.open "PROPPATCH", DestinURL, False
req.setRequestHeader "Content-Type", "text/xml;"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate", "f"
req.setRequestHeader "Content-Length:", Len(xmlstr)
req.send(xmlstr)
wscript.echo req.responsetext

end sub

Sub CopyContact(messageEntryID,ReplID)
set objcontact = objSession.getmessage(messageEntryID,objpubstore.ID)
set objCopyContact = objcontact.copyto(CntFolder.id,objSession.Inbox.StoreID)
objCopyContact.Unread = false
objCopyContact.Fields.Add &H3A2D001E, ReplID
objCopyContact.Update
Set objCopyContact = Nothing
wscript.echo objcontact.subject

end Sub

Sub DeleteContact(replUID)

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"">"
strQuery = strQuery & "<D:sql>SELECT ""DAV:Displayname"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & DestinURL & """') Where ""urn:schemas:contacts:internationalisdnnumber"" = '" & replUID & "' AND ""DAV:isfolder"" = False "
strQuery = strQuery & "</D:sql></D:searchrequest>"
req.open "SEARCH", DestinURL, false
req.setrequestheader "Content-Type", "text/xml"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate","f"
req.send strQuery
wscript.echo strQuery
wscript.echo req.responsetext
If req.status >= 500 Then
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: An error occurred on the server."
ElseIf req.status = 207 Then
   wscript.echo "Status: " & req.status
   wscript.echo "Status text:  " & req.statustext
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	wscript.echo oNode.text
	req.open "DELETE", oNode.text, false
	req.SetRequestHeader "cookie", reqsessionID
	req.SetRequestHeader "cookie", reqCadata
	req.send 
	wscript.echo "Status: " & req.status
   Next
Else
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: " & req.statustext
   wscript.echo "Response text: " & req.responsetext
End If

end Sub

function getpfid()

xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:e='http://schemas.microsoft.com/mapi/proptag/'><a:prop><e:x0FFF0102/></a:prop></a:propfind>"
req.open "PROPFIND", SourceURL, false, "", ""
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Depth", "0"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate", "f"
req.send xmlreqtxt
set oResponseDoc = req.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("d:x0FFF0102")
For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	getpfid = Octenttohex(oNode.nodeTypedValue)
Next

end function

Function Octenttohex(OctenArry)  
  ReDim aOut(UBound(OctenArry)) 
  For i = 1 to UBound(OctenArry) + 1 
    if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
    	aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
    else
	aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
    end if
  Next 
  Octenttohex = join(aOUt,"")
End Function 