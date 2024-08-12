<SCRIPT LANGUAGE="VBScript">

Sub ExStoreEvents_OnSave(pEventInfo, bstrURLItem, lFlags)

on error resume next
servername = "yourserver"
set DispEvtInfo = pEventInfo
set ADODBRec = DispEvtInfo.EventRecord
set objdom = CreateObject("MICROSOFT.XMLDOM")
Set objField = objDom.createElement("rss")
Set objattID = objDom.createAttribute("version")
objattID.Text = "2.0"
objField.setAttributeNode objattID
objDom.appendChild objField
Set objField1 = objDom.createElement("channel")
objfield.appendChild objField1
Set objField3 = objDom.createElement("title")
objfield3.text = "Public Folder Feed"
objfield1.appendChild objField3
Set objField4 = objDom.createElement("link")
objfield4.text = "http://" & servername & "/exchange" & right(ADODBRec.fields("Dav:parentname"),(len(ADODBRec.fields("Dav:parentname"))-instr(ADODBRec.fields("Dav:parentname"),"/MBX/"))-3)
objfield1.appendChild objField4
Set objField5 = objDom.createElement("description")
objfield5.text = "Public Folder Feed For Path"
objfield1.appendChild objField5
Set objField6 = objDom.createElement("language")
objfield6.text = "en-us"
objfield1.appendChild objField6
Set objField7 = objDom.createElement("lastBuildDate")
objfield7.text = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00 GMT"
objfield1.appendChild objField7
Set Rs = CreateObject("ADODB.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject")
Set msgobj = CreateObject("CDO.Message")
tyear = year(now()-7)
tmonth = month(now()-7)
if tmonth < 10 then tmonth = 0 & tmonth
stday = day(now()-7)
if stday < 10 then stday = 0 & stday
sttime = formatdatetime(now1,4)
qdatest = tyear & "-" & tmonth & "-" & stday & "T"
qdatest1 = qdatest & sttime & ":" & "00Z"
set Rec = CreateObject("ADODB.Record")
set Rec1 = CreateObject("ADODB.Record")
Set Conn = CreateObject("ADODB.Connection")
mailboxurl = ADODBRec.fields("Dav:parentname")
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open mailboxurl, ,3
SSql = "SELECT ""DAV:href"", ""DAV:contentclass"", ""urn:schemas:httpmail:textdescription"", ""urn:schemas:httpmail:datereceived"", "  
SSql = SSql & """urn:schemas:httpmail:fromemail"", ""urn:schemas:httpmail:subject"", ""DAV:ishidden"" " 
Ssql = SSql & " FROM scope('shallow traversal of """ & mailboxurl & """') " 
SSql = SSql & " WHERE (""urn:schemas:httpmail:datereceived"" > CAST(""" & qdatest1 & """ as 'dateTime')) AND ""DAV:isfolder"" = false"                 
Rs.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Rs.CursorType = 3
rs.open SSql, rec.ActiveConnection, 3
if Rs.recordcount <> 0 then 
Rs.movefirst
while not rs.eof
        if rs.fields("DAV:ishidden") = 0 then
		Set objField2 = objDom.createElement("item")
		objfield1.appendChild objField2
		Set objField8 = objDom.createElement("title")
		objfield8.text = rs.fields("urn:schemas:httpmail:subject")
		objfield2.appendChild objField8
		Set objField9 = objDom.createElement("link")
		objfield9.text = "http://" & servername & "/public" & right(Rs.fields("Dav:href"),(len(Rs.fields("Dav:href"))-instr(Rs.fields("Dav:href"),"/Public Folders/"))-14)
		objfield2.appendChild objField9
		Set objField10 = objDom.createElement("description")
		objfield10.text = Rs.fields("urn:schemas:httpmail:textdescription")
		objfield2.appendChild objField10
     	        Set objField11 = objDom.createElement("author")
		objfield11.text = rs.fields("urn:schemas:httpmail:fromemail")
		objfield2.appendChild objField11
		Set objField12 = objDom.createElement("pubDate")
		objfield12.text = WeekdayName(weekday(dtstartd),3) & ", " & day(dtstartd) & " " & Monthname(month(dtstartd),3) & " " & year(dtstartd) & " " & formatdatetime(rs.fields("urn:schemas:calendar:dtstart"),4) & ":00 GMT"
		objfield2.appendChild objField12
 		set objfield2 = nothing
		set objfield8 = nothing
		set objfield9 = nothing
		set objfield10 = nothing
		set objfield11 = nothing
	end if
	rs.movenext
wend
end if
rs.close
Set objPI = objDom.createProcessingInstruction("xml", "version='1.0'")
objDom.insertBefore objPI, objDom.childNodes(0)
objdom.save("\\webserver\rsspub\testpub.xml")

End Sub



</SCRIPT>