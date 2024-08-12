set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
dtListFrom = DateAdd("n", minTimeOffset, now())
dtListFrom = DateAdd("h",-wscript.arguments(1),dtListFrom)
strStartDateTime = year(dtListFrom)
if (Month(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Month(dtListFrom)
if (Day(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Day(dtListFrom)
if (Hour(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Hour(dtListFrom)
if (Minute(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Minute(dtListFrom)
if (Second(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Second(dtListFrom) & ".000000+000"

wscript.echo strStartDateTime

strComputer = wscript.arguments(0)
csCurrentdbFileName = "c:\temp\fiddb.xml"
csCurrentReportFileName = "c:\temp\fidreport.htm"

Set guUsersReport = CreateObject("Scripting.Dictionary")
Set fso = CreateObject("Scripting.FileSystemObject")
set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
xdXmlDocument.async="false"
xdXmlDocument.load(csCurrentdbFileName)

treport = "<table border=""1"" width=""100%"">" & vbcrlf
treport = treport & "  <tr>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Mailbox Name</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Date</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">User Attempting Access</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Folder Attempting Access To</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Folder Path</font></b></td>" & vbcrlf
treport = treport & "</tr>" & vbcrlf

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where EventCode = '1029' and TimeWritten >= '" & strStartDateTime  & "' ",,48)
For Each objEvent in colLoggedEvents
	slen = InStr(objEvent.Message,"The distinguished name of the owning mailbox is")
	elen = InStr(objEvent.Message,". The folder ID")
	struser = GetUser(Mid(objEvent.Message,slen+48,elen-(slen+48)))
    wscript.echo left(objEvent.Message,instr(objEvent.Message," "))
	wscript.echo struser
    farray =  GetFolder(Octenttohex(objEvent.Data))
	wscript.echo unescape(farray(0)) & "	" & unescape(farray(1))
	treport1 = ""
	treport1 = treport1 & "<tr>" & vbcrlf
	treport1 = treport1 & "<td align=""center"">" & struser & "&nbsp;</td>" & vbcrlf
	treport1 = treport1 & "<td align=""center"">" & WMIDateStringToDate(objEvent.TimeWritten) & "&nbsp;</td>" & vbcrlf
	treport1 = treport1 & "<td align=""center"">" & left(objEvent.Message,instr(objEvent.Message," ")) & "&nbsp;</td>" & vbcrlf
	treport1 = treport1 & "<td align=""center"">" & unescape(farray(0)) & "&nbsp;</td>" & vbcrlf
	treport1 = treport1 & "<td align=""center"">" & unescape(farray(1)) & "&nbsp;</td>" & vbcrlf
	treport1 = treport1 & "</tr>" & vbcrlf
	If guUsersReport.exists(struser) then
			guUsersReport(struser) = guUsersReport(struser) & treport1
	Else
			guUsersReport.add struser,treport1
	End if

Next
For Each strkeyUsr In guUsersReport
	treport = treport + guUsersReport(strkeyUsr)
	treport = treport + "<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
next
set wfile = fso.opentextfile(csCurrentReportFileName,2,true) 
wfile.write treport
wfile.close


Function Octenttohex(OctenArry) 
ReDim aOut(UBound(OctenArry)) 
For i = 0 To UBound(OctenArry)
		aOut(i) = Hex(OctenArry(i))
		If Len(aOut(i)) = 1 Then  aOut(i) =  "0" & aOut(i)
Next
For j = 1 To UBound(aOut) 
	If aOut(j) <> "00" Then outtext = outtext & aOut(j) 	
Next
Octenttohex = Replace(OctenArry(0) & "-" & LCase(outtext),"-0","-")
End Function 

Function GetUser(LegDN)
set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Com.ActiveConnection = Conn
Ldapfilter = "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(legacyExchangeDN=" & LegDN & ")))))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & Ldapfilter & ";displayName,distinguishedName;subtree"
com.Properties("Page Size") = 100
Com.CommandText = strQuery
Set Rs1 = Com.Execute
while not Rs1.eof
	getuser = rs1.fields("displayName")
	rs1.movenext
Wend

End Function 

Function GetFolder(fid)
ReDim retarray(1)
Set xnFidNodes = xdXmlDocument.selectNodes("//*[@fid = '" & fid & "']")
If xnFidNodes.length = 1 Then
	retarray(0) = xnFidNodes(0).attributes.getNamedItem("Name").nodeValue
	retarray(1) = xnFidNodes(0).attributes.getNamedItem("Path").nodeValue
else
	retarray(0) = "Not Found " & fid
	retarray(1) = "Not Found " & fid
End if
GetFolder = retarray
End Function 

Function WMIDateStringToDate(wmiConDate)
    WMIDateStringToDate = CDate(Mid(wmiConDate, 5, 2) & "/" & _
        Mid(wmiConDate, 7, 2) & "/" & Left(wmiConDate, 4) _
            & " " & Mid (wmiConDate, 9, 2) & ":" & Mid(wmiConDate, 11, 2) & ":" & Mid(wmiConDate,13, 2))
End Function