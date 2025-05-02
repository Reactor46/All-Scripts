set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"

report = "<table border=""1"" width=""100%"">" & vbcrlf
report = report & "  <tr>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Group-Name</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Total Number of Members</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Mailboxes</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Contacts-Internal</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Contacts-External</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Groups</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">QueryBased Dl's</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Public Folders</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">InterOrg-Person</font></b></td>" & vbcrlf
report = report & "</tr>" & vbcrlf

GALQueryFilter =  "(&(mail=*)(objectCategory=group))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
Com.ActiveConnection = Conn
Com.Properties("SearchScope") = 2          ' we want to search everything
Com.Properties("Page Size") = 500          ' and we want our records in lots of 500 (must be < query limit)
polQuery = "<LDAP://" & iAdRootDSE.Get("configurationNamingContext") &  ">;(objectCategory=msExchRecipientPolicy);distinguishedName,gatewayProxy;subtree"
Com.CommandText = polQuery 
set sdSMTPDomains = CreateObject("Scripting.Dictionary")
Set adRs = Com.Execute
while not adRs.eof
	for each adr in adRs.fields("gatewayproxy").value
		if instr(lcase(adr),"smtp:") then 
			if sdSMTPDomains.exists(LCase(right(adr,(len(adr)-instr(adr,"@"))))) then
			else
				sdSMTPDomains.Add LCase(right(adr,(len(adr)-instr(adr,"@")))),1
			end if
		end if		
	next
	adRs.movenext
Wend

Com.CommandText = strQuery
Set Rs = Com.Execute

while not rs.eof
	set objgroup = getobject("LDAP://" & replace(rs.fields("distinguishedName"),"/","\/"))	
	numcheck = 0
	usNumberUsers = 0
	cnNumberinContacts = 0
	cnNumberexContacts = 0
	nsNumberGroups = 0
	pfNumberPublicFolders = 0
	ioNumberIorgPersons = 0
	QbNumberQBDLs = 0
	for each member in objgroup.members
		wscript.echo member.class
		If member.mail <> "" then
			Select Case member.Class			
				Case "user" usNumberUsers = usNumberUsers + 1
				Case "contact"  wscript.echo right(member.mail,(len(member.mail)-instr(member.mail,"@")))
								If sdSMTPDomains.exists(LCase(right(member.mail,(len(member.mail)-instr(member.mail,"@"))))) Then
									cnNumberinContacts = cnNumberinContacts + 1
         						Else
									cnNumberexContacts = cnNumberexContacts + 1
								End if		
				Case "group" nsNumberGroups = nsNumberGroups + 1
				Case "publicFolder" pfNumberPublicFolders = pfNumberPublicFolders + 1
				Case "inetOrgPerson" ioNumberIorgPersons = ioNumberIorgPersons + 1
				Case "msExchDynamicDistributionList" QbNumberQBDLs = QbNumberQBDLs + 1
			End Select
			numcheck = numcheck + 1
		End if
	Next
    wscript.echo rs.fields("displayname")
	wscript.echo usNumberUsers
	wscript.echo cnNumberinContacts
	wscript.echo cnNumberexContacts
	wscript.echo nsNumberGroups
	wscript.echo QbNumberQBDLs
	wscript.echo pfNumberPublicFolders
	wscript.echo ioNumberIorgPersons
	report = report & "<tr>" & vbcrlf
	report = report & "<td align=""center"">" & rs.fields("displayname") & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & numcheck & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & usNumberUsers & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & cnNumberinContacts  & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & cnNumberexContacts & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & nsNumberGroups  & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & QbNumberQBDLs  & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & pfNumberPublicFolders  & "&nbsp;</td>" & vbcrlf
	report = report & "<td align=""center"">" & ioNumberIorgPersons  & "&nbsp;</td>" & vbcrlf
	report = report & "</tr>" & vbcrlf
	rs.movenext	
		
wend		

report = report & "</table>" & vbcrlf
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\Groupreport.htm",2,true) 
wfile.write report
wfile.close
set wfile = nothing
set fso = nothing

