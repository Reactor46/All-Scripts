Servername = wscript.arguments(0)
feedfile = ""c:\temp\feedpubnew.xml"
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
dtListFrom = DateAdd("n", minTimeOffset, now())
gmttime = dateadd("h",-toffset,now())
dateto = isodateit(gmttime)
datefrom = isodateit(DateAdd("d",-1,gmttime))
set objdom = CreateObject("MICROSOFT.XMLDOM")
set req = createobject("microsoft.xmlhttp")
rem Create Root RSS feed
Set objField = objDom.createElement("rss")
Set objattID = objDom.createAttribute("version")
objattID.Text = "2.0"
objField.setAttributeNode objattID
objDom.appendChild objField
Set objField1 = objDom.createElement("channel")
objfield.appendChild objField1
Set objField3 = objDom.createElement("link")
objfield3.text = "http://" & Servername & "/public"
objfield1.appendChild objField3
Set objField4 = objDom.createElement("title")
objfield4.text = "Public Folder Feed"
objfield1.appendChild objField4
Set objField5 = objDom.createElement("description")
objfield5.text = "New Public Folder items in the last 24 Hours"
objfield1.appendChild objField5
Set objField6 = objDom.createElement("language")
objfield6.text = "en-us"
objfield1.appendChild objField6
Set objField7 = objDom.createElement("lastBuildDate")
objfield7.text = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00 GMT"
objfield1.appendChild objField7

set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
polQuery = "<LDAP://" & strNameingContext &  ">;(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy));distinguishedName,gatewayProxy;subtree"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = polQuery
Set plRs = Com.Execute
while not plRs.eof
	for each adrobj in plrs.fields("gatewayProxy").value
		if instr(adrobj,"SMTP:") then dpDefaultpolicy = right(adrobj,(len(adrobj)-instr(adrobj,"@")))
	next
	plrs.movenext
wend
wscript.echo dpDefaultpolicy 
falias = "http://" & servername & "/exadmin/admin/" & dpDefaultpolicy & "/Public Folders/"
RecurseFolder(falias)
wscript.echo falias
set conn = nothing
set com = nothing
set wfile = nothing
set fso = Nothing
Set objPI = objDom.createProcessingInstruction("xml", "version='1.0'")
objDom.insertBefore objPI, objDom.childNodes(0)
objdom.save(feedfile)

Public Sub RecurseFolder(sUrl)
  
   req.open "SEARCH", sUrl, False, "", ""
   sQuery = "<?xml version=""1.0""?>"
   sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
   sQuery = sQuery & "<g:sql>SELECT ""http://schemas.microsoft.com/"
   sQuery = sQuery & "mapi/proptag/x0e080003"", ""DAV:hassubs"" FROM SCOPE "
   sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
   sQuery = sQuery & "WHERE ""DAV:isfolder"" = true and ""DAV:ishidden"" = false and ""http://schemas.microsoft.com/mapi/proptag/x36010003"" = 1"
   sQuery = sQuery & "</g:sql>"
   sQuery = sQuery & "</g:searchrequest>"
   req.setRequestHeader "Content-Type", "text/xml"
   req.setRequestHeader "Translate", "f"
   req.setRequestHeader "Depth", "0"
   req.setRequestHeader "Content-Length", "" & Len(sQuery)
   req.send sQuery
   Set oXMLDoc = req.responseXML
   Set oXMLSizeNodes = oXMLDoc.getElementsByTagName("d:x0e080003")
   Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")
   Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")
   For i = 0 to oXMLSizeNodes.length - 1
      call procfolder(oXMLHREFNodes.Item(i).nodeTypedValue,sUrl)
      wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
      If oXMLHasSubsNodes.Item(i).nodeTypedValue = True Then
         call RecurseFolder(oXMLHREFNodes.Item(i).nodeTypedValue)
      End If
   Next
End Sub

sub procfolder(strURL,pfname)
wscript.echo strURL
ReDim resarray(1,6)
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", ""DAV:getcontentlength"", "
strQuery = strQuery & """urn:schemas:httpmail:fromemail"",  ""urn:schemas:httpmail:to"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False AND " 
'strQuery = strQuery & """urn:schemas:httpmail:datereceived"" &lt; CAST(""" & dateto & """ as 'dateTime') AND "
strQuery = strQuery & """urn:schemas:httpmail:datereceived"" &gt; CAST(""" & datefrom & """ as 'dateTime')</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   set oSize = oResponseDoc.getElementsByTagName("a:getcontentlength")
   set odatereceived = oResponseDoc.getElementsByTagName("a:creationdate")
   set fEmail = oResponseDoc.getElementsByTagName("d:fromemail")
   set TEmail = oResponseDoc.getElementsByTagName("d:to")
   For i = 0 To (oNodeList.length -1)
		set oNode = oNodeList.nextNode
		set oNode1 = oNodeList1.nextNode
		set oNode2 = oSize.nextNode
		set oNode3 = odatereceived.nextNode
		set oNode4 = fEmail.nextNode
		set oNode5 = TEmail.nextNode
		wscript.echo oNode3.text
		export = 0
		If InStr(LCase(oNode4.text),LCase(domaintosearch))Then
			export = 1
		End If
		if InStr(LCase(oNode5.text),LCase(domaintosearch))Then
			export = 1
		End If
		If export = 1 Then
			Call AddtoFeed(oNode1.text,oNode.text)
		End if
   Next
Else
End If

end sub

sub AddtoFeed(exporthref,subject)

xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:m='urn:schemas:httpmail:' xmlns:mapi='http://schemas.microsoft.com/mapi/proptag/'>" _
& "<a:prop><mapi:x6707001E/></a:prop><a:prop><a:displayname/></a:prop><a:prop><m:subject/></a:prop><a:prop><m:fromemail/>"_
& "</a:prop><a:prop><m:htmldescription/></a:prop><a:prop><m:datereceived/></a:prop></a:propfind>"
req.open "PROPFIND", exporthref, false, "", ""
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Depth", "0"
req.setRequestHeader "Translate", "f"
req.send xmlreqtxt
set oResponseDoc1 = req.responseXML
set pfParentFolder = oResponseDoc1.getElementsByTagName("d:x6707001E")
set feFromEmail = oResponseDoc1.getElementsByTagName("e:fromemail")
set sjSubject = oResponseDoc1.getElementsByTagName("e:subject")
set drDateRecieved = oResponseDoc1.getElementsByTagName("e:datereceived")
set bdHtmlBody = oResponseDoc1.getElementsByTagName("e:htmldescription")
set dnDisplayName = oResponseDoc1.getElementsByTagName("a:displayname")

For i = 0 To (sjSubject.length -1)
	set pfnode = sjSubject.nextNode
	set pfnode1 = feFromEmail.nextNode
	set pfnode2 = drDateRecieved.nextNode
	set pfnode3 = bdHtmlBody.nextNode
	Set pfnode4 = pfParentFolder.nextNode
	Set pfnode5 = dnDisplayName.nextNode
	wscript.echo pfnode.text
	wscript.echo pfnode1.text
	wscript.echo pfnode2.text
	rem wscript.echo pfnode3.text
	wscript.echo pfnode4.text
	wscript.echo pfnode5.text
	wscript.echo left(Replace(pfnode2.text,"T"," "),19)
	Set objField2 = objDom.createElement("item")
	objfield1.appendChild objField2
	Set objField8 = objDom.createElement("guid")
	Set objattID8 = objDom.createAttribute("isPermaLink")
	objattID8.Text = "false"
	objField8.setAttributeNode objattID8
	objfield8.text = exporthref
	objfield2.appendChild objField8
	Set objField9 = objDom.createElement("title")
	objfield9.text = pfnode.text
	if objfield9.text = "" then objfield9.text = "Blank"
	objfield2.appendChild objField9
	Set objField10 = objDom.createElement("link")
	objfield10.text = "http://" & Servername & "/public" & pfnode4.text
	objfield2.appendChild objField10
	Set objField11 = objDom.createElement("description")
	objfield11.text = pfnode3.text
    	if objfield11.text = "" then objfield11.text = "Blank"
	objfield2.appendChild objField11
   	Set objField12 = objDom.createElement("author")
	objfield12.text = pfnode1.text
	objfield2.appendChild objField12
	Set objField13 = objDom.createElement("pubDate")
	objfield13.text = WeekdayName(weekday(left(Replace(pfnode2.text,"T"," "),19)),3) & ", " & day(left(Replace(pfnode2.text,"T"," "),19)) & " " & Monthname(month(left(Replace(pfnode2.text,"T"," "),19)),3) & " " & year(left(Replace(pfnode2.text,"T"," "),19)) & " " & formatdatetime(left(Replace(pfnode2.text,"T"," "),19),4) & ":00 GMT"
	objfield2.appendChild objField13
	Set objField14 = objDom.createElement("category")
	objfield14.text = unescape(Replace(LCase(pfnode4.text),LCase(pfnode5.text),""))
	objfield2.appendChild objField14
	set objfield2 = nothing
	set objfield8 = nothing
	set objfield9 = nothing
	set objfield10 = nothing
	set objfield11 = nothing
	set objfield12 = nothing
	set objfield13 = nothing
next

End Sub


function isodateit(datetocon)
strDateTime = year(datetocon) & "-"
if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
strDateTime = strDateTime & Month(datetocon) & "-"
if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
strDateTime = strDateTime & Day(datetocon) & "T" & formatdatetime(datetocon,4) &":00Z"
isodateit = strDateTime
end function 



