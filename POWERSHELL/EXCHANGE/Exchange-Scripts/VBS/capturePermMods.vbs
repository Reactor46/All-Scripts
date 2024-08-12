Const RIGHT_DS_DELETE = &H10000
Const RIGHT_DS_READ = &H20000
Const RIGHT_DS_CHANGE = &H40000
Const RIGHT_DS_TAKE_OWNERSHIP = &H80000
Const RIGHT_DS_MAILBOX_OWNER = &H1
Const RIGHT_DS_SEND_AS = &H2
Const RIGHT_DS_PRIMARY_OWNER = &H4

csCurrentSnapFileName = "c:\temp\currentSnap.xml"
psPreviousSnapFileName = "c:\temp\prevSnap.xml"
adArchieveDirectory = "c:\temp\SnapArchive\"
rfReportFileName = "c:\temp\ACLChangeReport-"
rrRightReport = 0

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(csCurrentSnapFileName) Then
	wscript.echo "Snap Exists"
	If fso.FileExists(psPreviousSnapFileName) Then fso.deletefile(psPreviousSnapFileName)
	fso.movefile csCurrentSnapFileName, psPreviousSnapFileName
	set xdXmlDocument = CreateObject("Microsoft.XMLDOM")
	xdXmlDocument.async="false"
	xdXmlDocument.load(psPreviousSnapFileName)
	Set xnSnaptime = xdXmlDocument.selectNodes("//SnappedACLS")
	For Each exSnap In xnSnaptime
		oldSnap = exSnap.attributes.getNamedItem("SnapDate").nodeValue
		wscript.echo "Snap Taken : " & oldSnap 
		takesnap
		afFileName = adArchieveDirectory & Replace(Replace(Replace(exSnap.attributes.getNamedItem("SnapDate").nodeValue,":",""),",","")," ","") & ".xml"
		wscript.echo "Archiving Old Snap to : " & afFileName
		fso.copyfile psPreviousSnapFileName, afFileName
	Next
	set xdXmlDocument1 = CreateObject("Microsoft.XMLDOM")
	xdXmlDocument1.async="false"
	xdXmlDocument1.load(csCurrentSnapFileName)
	Set ckCurrentPerms = CreateObject("Scripting.Dictionary")
	Set pkPreviousPerms = CreateObject("Scripting.Dictionary")
	Set xnCurrentPermsUsers = xdXmlDocument1.selectNodes("//User")
	For Each xnUserNode In xnCurrentPermsUsers
		unUserName =  xnUserNode.attributes.getNamedItem("SamaccountName").nodeValue
		For Each caACLs In xnUserNode.ChildNodes
			ckCurrentACL = unUserName & "|-|" & caACLs.attributes.getNamedItem("User").nodeValue
		    ckCurrentPerms.add   ckCurrentACL, caACLs.attributes.getNamedItem("Right").nodeValue
		Next
	Next
	Set xnPrevPermsUsers = xdXmlDocument.selectNodes("//User")
	For Each xnUserNode1 In xnPrevPermsUsers
		unUserName1 =  xnUserNode1.attributes.getNamedItem("SamaccountName").nodeValue
		For Each caACLs1 In xnUserNode1.ChildNodes
			pkPrevACL = unUserName1 & "|-|" & caACLs1.attributes.getNamedItem("User").nodeValue
		    pkPreviousPerms.add   pkPrevACL, caACLs1.attributes.getNamedItem("Right").nodeValue
			rem Do a Check for Any Deleted or Changed Permisssions
			If ckCurrentPerms.exists(pkPrevACL) Then
				If 	ckCurrentPerms(pkPrevACL) <> caACLs1.attributes.getNamedItem("Right").nodeValue Then
					rrRightReport = 1 
					wscript.echo "Found Changed ACL " 
					wscript.echo "Old Rights : "  & pkPrevACL & "	" & caACLs1.attributes.getNamedItem("Right").nodeValue
					wscript.echo "New Rights : "  & pkPrevACL & "	" & ckCurrentPerms(pkPrevACL)
					hrmodHtmlReport = hrmodHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & unUserName1 & " </font></td>" & vbcrlf
					hrmodHtmlReport = hrmodHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">Old Rights: " &  caACLs1.attributes.getNamedItem("User").nodeValue _
					& "	" & caACLs1.attributes.getNamedItem("Right").nodeValue &  " </font></td>" & vbcrlf
					hrmodHtmlReport = hrmodHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">New Rights: " _ 
					&  caACLs1.attributes.getNamedItem("User").nodeValue & "	" & ckCurrentPerms(pkPrevACL) & " </font></td></tr>" & vbcrlf
				End if
			Else
				rrRightReport = 1 
				hrDelHtmlReport = hrDelHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & unUserName1 & " </font></td>" & vbcrlf
				hrDelHtmlReport = hrDelHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  caACLs1.attributes.getNamedItem("User").nodeValue _
					& "	" & caACLs1.attributes.getNamedItem("Right").nodeValue & " </font></td></tr>" & vbcrlf
				Wscript.echo "Found Deleted ACL : " &  pkPrevACL & "	" & caACLs1.attributes.getNamedItem("Right").nodeValue
			End if
		Next
	Next
	rem Do forward check of ACL's
	For Each dkCurrenPermKey In ckCurrentPerms.keys
		If Not pkPreviousPerms.exists(dkCurrenPermKey) Then
			rrRightReport = 1 
		    dknewpermarry = Split(dkCurrenPermKey,"|-|")
			hrnewHtmlReport = hrnewHtmlReport & "<tr><td><font face=""Arial"" color=""#000080"" size=""2"">" & dknewpermarry(0) & " </font></td>" & vbcrlf
			hrnewHtmlReport = hrnewHtmlReport & "<td><font face=""Arial"" color=""#000080"" size=""2"">" &  dknewpermarry(1) _
			& "	" & ckCurrentPerms(dkCurrenPermKey) & " </font></td></tr>" & vbcrlf
			Wscript.echo "Found new ACL : "  & dkCurrenPermKey & "	" &  ckCurrentPerms(dkCurrenPermKey)
		End if
	next
Else
	wscript.echo "No current permissions snap exists taking snap"
	Call TakeSnap
End If
If rrRightReport = 1 Then
	wscript.echo "Writing Report"
	hrHtmlReport = "<html><body>" & vbcrlf
	NewSnapDate = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00" 
	hrHtmlReport = hrHtmlReport  & "<p><font size=""4"" face=""Arial Black"" color=""#008000"">Change To Mailbox Rights Report for Snaps Taken Between - </font>" & oldSnap  & " and "_
	&  NewSnapDate & "</font></p>" & vbcrlf
	If hrnewHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">ACL's Added</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & Replace(Replace(hrnewHtmlReport,"-exra-",""),"-exsa-","") & "</table>"
	End If
	If hrmodHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">ACL's Modified</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & Replace(Replace(hrmodHtmlReport,"-exra-",""),"-exsa-","") & "</table>"
	End If
	If hrDelHtmlReport <> "" Then
		hrHtmlReport = hrHtmlReport & "<p><font face=""Arial"" color=""#000080"" size=""2"">ACL's Deleted</font></p>"
		hrHtmlReport = hrHtmlReport & "<table border=""1"" width=""100%"" id=""table1"" cellspacing=""0"" cellpadding=""0"" bordercolor=""#000000"">"
		hrHtmlReport = hrHtmlReport & Replace(Replace(hrDelHtmlReport,"-exra-",""),"-exsa-","") & "</table>"
	End If
	hrHtmlReport = hrHtmlReport  &  "</body></html>" & vbcrlf
	rfReportFileName = rfReportFileName & Replace(Replace(Replace(NewSnapDate,":",""),",","")," ","")  & ".htm"
	wscript.echo rfReportFileName
	set rfile = fso.opentextfile(rfReportFileName,2,true) 
	rfile.writeline(hrHtmlReport)
End If 



Sub TakeSnap

set wfile = fso.opentextfile(csCurrentSnapFileName,2,true) 
wfile.writeline("<?xml version=""1.0""?>")
wfile.writeline("<SnappedACLS SnapDate=""" & WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00" & """>")

Set objSystemInfo = CreateObject("ADSystemInfo") 
strdname = objSystemInfo.DomainShortName
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString	
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Query = "<LDAP://" & strNameingContext & ">;(&(&(& (mailnickname=*) (| (&(objectCategory=person)(objectClass=user)(|(homeMDB=*)(msExchHomeServerName=*))) ))));samaccountname,displayname,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = Query
Com.Properties("Page Size") = 1000
Set Rs = Com.Execute
While Not Rs.EOF
	dn = "LDAP://" & replace(rs.Fields("distinguishedName").Value,"/","\/")
	set objuser = getobject(dn)
	Set oSecurityDescriptor = objuser.Get("msExchMailboxSecurityDescriptor")
	Set oUserSecurityDescriptor = objuser.Get("ntSecurityDescriptor")
	Set oUserdacl = oUserSecurityDescriptor.DiscretionaryAcl
	Set oUserace = CreateObject("AccessControlEntry")
	Set dacl = oSecurityDescriptor.DiscretionaryAcl
	Set ace = CreateObject("AccessControlEntry")
	fwFirstWrite = 0
	For Each ace In dacl
		   if ace.AceFlags <> 18 then
			if ace.Trustee <> "NT AUTHORITY\SELF" Then
				If fwFirstWrite = 0 Then
					wfile.writeline("	<User SamaccountName=""" & rs.fields("samaccountname") & """>")
					fwFirstWrite = 1
				End if
				wfile.writeline("<ACE User=""" & ace.Trustee & """ Right=""" & getRights(ace.AccessMask) & """></ACE>")
			end if
		   end if
	Next
	For Each oUserace In oUserdacl
		   if lcase(oUserace.ObjectType) = "{ab721a54-1e2f-11d0-9819-00aa0040529b}" and oUserace.AceType = 5 Then
				If oUserace.AceFlags <> 26 then
					if oUserace.Trustee <> "NT AUTHORITY\SELF" and oUserace.AceFlags <> 6 Then
						if fwFirstWrite = 0 Then
							wfile.writeline("	<User SamaccountName=""" & rs.fields("samaccountname") & """>")
							fwFirstWrite = 1
						End If
						wfile.writeline("<ACE User=""" & oUserace.Trustee & "-exsa-" & """ Right=""Send As""></ACE>")
					end If
				End if
		   end if
		   if lcase(oUserace.ObjectType) = "{ab721a56-1e2f-11d0-9819-00aa0040529b}" and oUserace.AceType = 5 Then
				If oUserace.AceFlags <> 26 then		
					if oUserace.Trustee <> "NT AUTHORITY\SELF" and oUserace.AceFlags <> 6 then
						If fwFirstWrite = 0 Then
							wfile.writeline("	<User SamaccountName=""" & rs.fields("samaccountname") & """>")
							fwFirstWrite = 1
						End If
						wfile.writeline("<ACE User=""" & oUserace.Trustee & "-exra-" &  """ Right=""Recieve As""></ACE>")
					end If
				End if
		   end if
	Next


	If fwFirstWrite = 1 then
		wfile.writeline("</User>")
	End if
	rs.movenext
Wend
wfile.writeline("</SnappedACLS>")
wscript.echo "New Snap Taken"

End Sub

Function getRights(hvHexValue)
		If (hvHexValue And RIGHT_DS_SEND_AS) Then
			getRights =  "Send As"
		End If
		If (hvHexValue And RIGHT_DS_CHANGE) Then
			getRights =  "Modify user attributes"
		End If
		If (hvHexValue And RIGHT_DS_DELETE) Then
			getRights =  "Delete mailbox store"
		End If
		If (hvHexValue And RIGHT_DS_READ) Then
			getRights =  "Read permissions"
		End If
		If (hvHexValue And RIGHT_DS_TAKE_OWNERSHIP) Then
			getRights =  "Take Ownership"
		End If
		If (hvHexValue And RIGHT_DS_MAILBOX_OWNER) Then
			getRights =  "Mailbox Owner"
		End If
		If (hvHexValue And RIGHT_DS_PRIMARY_OWNER) Then
			getRights =  "Mailbox Primary Owner"
		End If
End Function