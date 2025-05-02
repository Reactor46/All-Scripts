<SCRIPT LANGUAGE="VBScript">

Sub ExStoreEvents_OnSave(pEventInfo, bstrURLItem, lFlags)

on error resume next
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
objfield3.text = "Calendar Folder Feed"
objfield1.appendChild objField3
Set objField4 = objDom.createElement("link")
objfield4.text = "http://servername/exchange" & right(ADODBRec.fields("Dav:parentname"),(len(ADODBRec.fields("Dav:parentname"))-instr(ADODBRec.fields("Dav:parentname"),"/MBX/"))-3)
objfield1.appendChild objField4
Set objField5 = objDom.createElement("description")
objfield5.text = "Calendar Feed For Path"
objfield1.appendChild objField5
Set objField6 = objDom.createElement("language")
objfield6.text = "en-us"
objfield1.appendChild objField6
Set objField7 = objDom.createElement("lastBuildDate")
objfield7.text = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00 GMT"
objfield1.appendChild objField7
Set Rs = CreateObject("ADODB.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject")
set shell = CreateObject("WScript.Shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
dtListFrom = DateAdd("n", minTimeOffset, now())
dtListTo = isodateit(DateAdd("d",7,dtListFrom))
dtListFrom = isodateit(dtListFrom)
set Rec = CreateObject("ADODB.Record")
set Rec1 = CreateObject("ADODB.Record")
Set Conn = CreateObject("ADODB.Connection")
mailboxurl = ADODBRec.fields("Dav:parentname")
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open mailboxurl, ,3
SSql = "Select ""DAV:href"", ""urn:schemas:httpmail:subject"", ""urn:schemas:calendar:dtstart"", ""urn:schemas:calendar:dtend"", "
SSql = SSql &  " ""urn:schemas:calendar:organizer"", ""urn:schemas:calendar:location"", ""DAV:contentclass"", "
SSql = SSql &  " ""urn:schemas:httpmail:textdescription"", ""urn:schemas:httpmail:fromemail"", ""DAV:ishidden"" "
SSql = SSql &  " FROM scope('shallow traversal of """ & mailboxurl & """') " 
SSql = SSql &  "WHERE (""urn:schemas:calendar:dtstart"" >= CAST(""" & dtListFrom & """ as 'dateTime')) "
SSql = SSql &  "AND (""urn:schemas:calendar:dtstart"" < CAST(""" & dtListTo & """ as 'dateTime'))" 
SSql = SSql &  " AND ""DAV:contentclass"" = 'urn:content-classes:appointment' ORDER BY ""urn:schemas:calendar:dtstart"" ASC"
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
		objfield9.text = "http://servername/exchange" & right(Rs.fields("Dav:href"),(len(Rs.fields("Dav:href"))-instr(Rs.fields("Dav:href"),"/MBX/"))-3)
		objfield2.appendChild objField9
		Set objField10 = objDom.createElement("description")
		if isnull(Rs.fields("urn:schemas:httpmail:textdescription")) then
			Etext = "Starts : " & dateadd("h",toffset,rs.fields("urn:schemas:calendar:dtstart")) & "<BR>"
			Etext = Etext & "Ends : " & dateadd("h",toffset,rs.fields("urn:schemas:calendar:dtend")) & "<BR>"
			Etext = Etext & "Location : " & rs.fields("urn:schemas:calendar:location")
			objfield10.text = Etext
		else
			Etext = "Starts : " & dateadd("h",toffset,rs.fields("urn:schemas:calendar:dtstart")) & "<BR>"
			Etext = Etext & "Ends : " & dateadd("h",toffset,rs.fields("urn:schemas:calendar:dtend")) & "<BR>"
			Etext = Etext & "Location : " & rs.fields("urn:schemas:calendar:location") & "<BR><BR>"
			Etext = Etext & Rs.fields("urn:schemas:httpmail:textdescription")
			objfield10.text = Etext
		end if
		objfield2.appendChild objField10
     	        Set objField11 = objDom.createElement("author")
		objfield11.text = rs.fields("urn:schemas:httpmail:fromemail")
		objfield2.appendChild objField11
		Set objField12 = objDom.createElement("pubDate")
		dtstartd = rs.fields("urn:schemas:calendar:dtstart")
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
objdom.save("\\servname\wwwroot\calpub4.xml")

End Sub

function isodateit(datetocon)
	strDateTime = year(datetocon) & "-"
	if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Month(datetocon) & "-"
	if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Day(datetocon) & "T" & formatdatetime(datetocon,4) & ":00Z"
	isodateit = strDateTime
end function


</SCRIPT>