snServername = "webmail.domain.com"
mnMailboxname = "mailbox"
domain = "domain"
strpassword = "password"

FeedTitle = "Username - Calendar"
scriptAgentName = "PortCalScriptAgent-user@domain.com"
strusername =  domain & "\" & mnMailboxname
SourceURL = "https://" & snServername & "/exchange/" & mnMailboxname & "/calendar"
xfFeedPath = "c:\calfeedvc3.xml"

set shell = createobject("wscript.shell")
set xdXmlDocument = CreateObject("MICROSOFT.XMLDOM")
set fso = createobject("Scripting.FileSystemObject")
set req = createobject("microsoft.xmlhttp")
set vcreq = createobject("microsoft.xmlhttp")

strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
rdate =  dateadd("h",-toffset,now())

if fso.fileexists(xfFeedPath) then
	xdXmlDocument.async="false"
	xdXmlDocument.load(xfFeedPath)		
else
	BuildHeader(FeedTitle)
	xdXmlDocument.async="false"
	xdXmlDocument.load(xfFeedPath)	
end if

szXml = "destination=https://" & snServername & "/exchange/&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & snServername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
	if instr(lcase(reqhedrarry(c)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
rsscollblob = ""
rsscollblob = Collabblobget()
QueryCalendar(SourceURL)
xdXmlDocument.save(xfFeedPath)


function QueryCalendar(mbMailboxURL)

xsXmlString = "<a:searchrequest xmlns:a=""DAV:"" xmlns:R=""http://schemas.microsoft.com/repl/"">" _  
& "<R:repl><R:collblob>" & rsscollblob & "</R:collblob></R:repl>" _
& "<a:sql>" _ 
& "Select ""urn:schemas:calendar:location"" AS location,""urn:schemas:httpmail:subject"" AS subject," _  
& " ""urn:schemas:calendar:dtstart"" AS dtstart,""urn:schemas:calendar:dtend"" AS dtend," _ 
& " ""urn:schemas:calendar:busystatus"" AS busystatus,""urn:schemas:calendar:instancetype"" AS instancetype," _
& " ""urn:schemas:calendar:alldayevent"" AS alldayevent,""urn:schemas:calendar:remindernexttime"" AS remindernexttime," _
& " ""http://schemas.microsoft.com/mapi/proptag/x0fff0102"", ""http://schemas.microsoft.com/repl/repl-uid""," _ 
& " ""http://schemas.microsoft.com/mapi/proptag/x10c70003"" AS pr_cdormid," _ 
& " ""http://schemas.microsoft.com/exchange/sensitivity"" AS sensitivity," _  
& " ""http://schemas.microsoft.com/mapi/id/{00020329-0000-0000-C000-000000000046}/0x8543"" As SxID, " _
& " ""http://schemas.microsoft.com/mapi/apptstateflags"" AS apptstateflags," _  
& " ""urn:schemas:calendar:created"" As CreatedDate, " _
& " ""urn:schemas:calendar:uid"" AS uid    FROM Scope('SHALLOW TRAVERSAL OF """"')" _  
& " WHERE  NOT (""urn:schemas:calendar:instancetype"" = 2 oR ""urn:schemas:calendar:instancetype"" = 3)   AND (""DAV:ishidden"" is Null  OR ""DAV:contentclass"" = 'urn:content-classes:appointment')" _
& " ORDER BY ""urn:schemas:calendar:created"" ASC </a:sql>" _
& "</a:searchrequest>" 

req.open "SEARCH", mbMailboxURL, false
req.setrequestheader "Content-Type", "text/xml"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate","f"
req.send xsXmlString

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
   set idNodeList = oResponseDoc.getElementsByTagName("e:x0fff0102")
   set replidNodeList = oResponseDoc.getElementsByTagName("d:repl-uid")
   set replchangeType = oResponseDoc.getElementsByTagName("d:changetype")
   set slSubjectList = oResponseDoc.getElementsByTagName("subject")
   set ctCreatedTime = oResponseDoc.getElementsByTagName("CreatedDate")
   set irItemHref = oResponseDoc.getElementsByTagName("a:href")
   set sxidList = oResponseDoc.getElementsByTagName("SxID")
   for id = 0 To (idNodeList.length -1)
	set oNode1 = idNodeList.nextNode
	set oNode2 = replidNodeList.nextNode
	set oNode3 = replchangeType.nextNode
	set oNode4 = slSubjectList.nextNode
	set oNode5 = ctCreatedTime.nextNode
	set oNode6 = irItemHref.nextNode
	set oNode7 = sxidList.nextNode
	select case oNode3.text
		case "new" call additem(oNode2.text,oNode4.text,oNode5.text,oNode6.text,Octenttohex(oNode1.nodeTypedValue),oNode7.text)
		case "delete" call deleteitem(oNode2.text)      
		case "change" call changeitem(oNode2.text,oNode5.text,oNode6.text)
	end select
   next
Else
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: " & req.statustext
   wscript.echo "Response text: " & req.responsetext
End If


end function


Function BuildHeader(ctCalendarTitle)

Set objPI = xdXmlDocument.createProcessingInstruction("xml", "version='1.0'")
xdXmlDocument.insertBefore objPI, xdXmlDocument.childNodes(0)
Set elElement1 = xdXmlDocument.CreateElement("rss")
Set elElement2 = xdXmlDocument.CreateElement("channel")
Set elElement3 = xdXmlDocument.CreateElement("title")
Set elElement3a = xdXmlDocument.CreateElement("link")
Set elElement4 = xdXmlDocument.CreateElement("description")
Set elElement5 = xdXmlDocument.CreateElement("language")
Set elElement6 = xdXmlDocument.CreateElement("lastBuildDate")
Set elElement8 = xdXmlDocument.CreateElement("portCalendar:collblob")
Set elElement9 = xdXmlDocument.CreateElement("sx:sharing")

elElement3.text = ctCalendarTitle
elElement3a.text = SourceURL
elElement4.text = "Portable Calendar Feed"
elElement5.text = "en-us"
elElement6.text = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00 GMT"
elElement8.text = ""

Set objattID = xdXmlDocument.createAttribute("version")
objattID.Text = "2.0"
xdXmlDocument.appendChild elElement1
set objattID1 = xdXmlDocument.createAttribute("xmlns:portCalendar")
objattID1.Text = "http://msgdev.mvps.org/portCalendar/"
set objattID2 = xdXmlDocument.createAttribute("xmlns:sx")
objattID2.Text = "http://www.microsoft.com/schemas/rss/sse"
set objattID3 = xdXmlDocument.createAttribute("version")
objattID3.Text = "0.91"
elElement1.setAttributeNode objattID
elElement1.setAttributeNode objattID1
elElement1.setAttributeNode objattID2
elElement9.setAttributeNode objattID3
elElement1.appendChild elElement9
elElement1.appendChild elElement2
elElement2.appendChild elElement3
elElement2.appendChild elElement3a
elElement2.appendChild elElement4
elElement2.appendChild elElement5
elElement2.appendChild elElement6
elElement2.appendChild elElement8


xdXmlDocument.save(xfFeedPath)

end function

function Collabblobset(collblob)

Set colNode = xdXmlDocument.documentElement.selectSingleNode("channel/portCalendar:collblob")
colNode.text = collblob
xdXmlDocument.save(xfFeedPath)

end function

function Collabblobget()

Set colNode = xdXmlDocument.documentElement.selectSingleNode("channel/portCalendar:collblob")
Collabblobget = colNode.text

end function

function additem(riReplID,csCalenderSubject,csCreatedTime,itHref,eiEntryID,SxiD)
wscript.echo itHref
wscript.echo SxiD
csCreatedTime = cdate(mid(csCreatedTime,1,10) & " " & mid(csCreatedTime,12,8))
Set chnlNode = xdXmlDocument.documentElement.selectSingleNode("channel")
set oNodeList2 = xdXmlDocument.getElementsByTagName("item")
Set SelElement1 = xdXmlDocument.CreateElement("item")
Set SelElement2 = xdXmlDocument.CreateElement("guid")
Set SelElement3 = xdXmlDocument.CreateElement("title")
Set SelElement4 = xdXmlDocument.CreateElement("description")
Set SelElement5 = xdXmlDocument.CreateElement("author")
Set SelElement6 = xdXmlDocument.CreateElement("pubDate")
Set SelElement7 = xdXmlDocument.CreateElement("portCalendar:entryID")
Set SelElement8 = xdXmlDocument.CreateElement("sx:sync")
Set SelElement9 = xdXmlDocument.CreateElement("sx:history")


Set SelobjattID = xdXmlDocument.createAttribute("id")
Set SelobjattID1 = xdXmlDocument.createAttribute("version")
Set SelobjattID2 = xdXmlDocument.createAttribute("deleted")
Set SelobjattID3 = xdXmlDocument.createAttribute("conflict")
Set SelobjattID4 = xdXmlDocument.createAttribute("when")
Set SelobjattID5 = xdXmlDocument.createAttribute("by")
Set SelobjattID6 = xdXmlDocument.createAttribute("isPermaLink")


if oNodeList2.length = 0 then
	chnlNode.appendChild SelElement1
else
	Set fitemNode = xdXmlDocument.documentElement.selectSingleNode("channel/item")
	chnlNode.insertBefore SelElement1, fitemNode
end if
selElement2.text = riReplID
selElement3.text = csCalenderSubject
set cdel = xdXmlDocument.createCDATASection(getVcard(itHref))
selElement5.text = scriptAgentName
selElement6.text =  WeekdayName(weekday(csCreatedTime),3) & ", " & day(csCreatedTime) & " " & Monthname(month(csCreatedTime),3) & " " & year(csCreatedTime) & " " & formatdatetime(csCreatedTime,4) & ":00 GMT"
SelElement7.text = eiEntryID

if SxiD = "" then
	SelobjattID.Text =  eiEntryID
else
	SelobjattID.Text = SxiD
end if
SelobjattID1.Text = "1"
SelobjattID2.Text = "false"
SelobjattID3.Text = "false"
SelobjattID4.Text = WeekdayName(weekday(rdate),3) & ", " & day(rdate) & " " & Monthname(month(rdate),3) & " " & year(rdate) & " " & formatdatetime(rdate,4) & ":00 GMT"
SelobjattID5.Text = scriptAgentName
SelobjattID6.Text = "false"


selElement8.setAttributeNode SelobjattID
selElement8.setAttributeNode SelobjattID1
selElement8.setAttributeNode SelobjattID2
selElement8.setAttributeNode SelobjattID3
selElement9.setAttributeNode SelobjattID4
selElement9.setAttributeNode SelobjattID5
selElement2.setAttributeNode SelobjattID6

selElement1.appendChild selElement2
selElement1.appendChild selElement3
selElement1.appendChild selElement4
selElement4.appendChild cdel
selElement1.appendChild selElement5
selElement1.appendChild selElement6
selElement1.appendChild SelElement7
selElement1.appendChild SelElement8
selElement8.appendChild SelElement9

xdXmlDocument.save(xfFeedPath)

end function

function changeitem(riReplID,csCreatedTime,itHref)

Set mnModeNode = xdXmlDocument.selectNodes("//*[text() = '" & riReplID & "']")
Wscript.echo "Nodes Found: " & mnModeNode.length
for each mnMode in mnModeNode
	set inItemnode = mnMode.parentnode
	for each dsnode in inItemnode.childnodes
		select case dsnode.nodename 
			case "description" set cdel = xdXmlDocument.createCDATASection(getVcard(itHref))
			     dsnode.text = ""
			     dsnode.appendchild cdel
			     xdXmlDocument.save(xfFeedPath)
			case "sx:sync" call updatesxnode(dsnode,"update")
		end select
	next
next

end function

function updatesxnode(dsnode,upmode)
	for each csnode in dsnode.childnodes 
		if csnode.nodename = "sx:history" then
			wscript.echo "num nodes" & csnode.childnodes.length
			Set sxSelElement1 = xdXmlDocument.CreateElement("sx:update")
			Set sxselobjattID = xdXmlDocument.createAttribute("when")
			Set sxSelobjattID1 = xdXmlDocument.createAttribute("by")
			sxselobjattID.Text = csnode.attributes.getNamedItem("when").nodeValue
			sxselobjattID1.Text = csnode.attributes.getNamedItem("by").nodeValue
			if upmode = "delete" then
				dsnode.attributes.getNamedItem("deleted").nodeValue = "true" 
			end if
			sxSelElement1.setAttributeNode sxselobjattID
			sxSelElement1.setAttributeNode sxselobjattID1
			csnode.attributes.getNamedItem("when").nodeValue = WeekdayName(weekday(rdate),3) & ", " & day(rdate) & " " & Monthname(month(rdate),3) & " " & year(rdate) & " " & formatdatetime(rdate,4) & ":00 GMT"
			csnode.attributes.getNamedItem("by").nodeValue = scriptAgentName
			if csnode.childnodes.length = 0 then
				csnode.appendChild sxSelElement1
			else
				csnode.insertBefore sxSelElement1, csnode.childnodes(0)
				wscript.echo "Inserting"
			end if
			dsnode.attributes.getNamedItem("version").nodeValue = cint(dsnode.attributes.getNamedItem("version").nodeValue) + 1
			xdXmlDocument.save(xfFeedPath)
		end if
	next
end function

function deleteitem(riReplID)

Set mnModeNode = xdXmlDocument.selectNodes("//*[text() = '" & riReplID & "']")
Wscript.echo "Nodes Found: " & mnModeNode.length
for each mnMode in mnModeNode
	set inItemnode = mnMode.parentnode
	for each dsnode in inItemnode.childnodes
		select case dsnode.nodename 
			case "description" dsnode.text = ""
				  xdXmlDocument.save(xfFeedPath)
			case "sx:sync" call updatesxnode(dsnode,"delete")
		end select
	next
next

end function

function getVcard(href)

vcreq.open "GET", href, false,"",""
vcreq.setrequestheader "Content-Type", "text/xml"
vcreq.SetRequestHeader "cookie", reqsessionID
vcreq.SetRequestHeader "cookie", reqCadata
vcreq.setRequestHeader "Translate","f"
vcreq.send 
mstream = vcreq.responsetext
vstream = mid(mstream,instr(mstream,"BEGIN:VCALENDAR"),(instr(mstream,"END:VCALENDAR")+13)-instr(mstream,"BEGIN:VCALENDAR"))
getVcard = vstream

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