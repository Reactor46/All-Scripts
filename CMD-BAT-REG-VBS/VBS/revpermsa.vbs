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
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS UOADDisplayName, " & _
           "  NEW adVarChar(255) AS UOADTrusteeName, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS MRmbox, " & _
           "      NEW adVarChar(255) AS MRTrusteeName, " & _
           "      NEW adVarChar(255) AS MRRights, " & _
           "      NEW adVarChar(255) AS MRAceflags) " & _
           "      RELATE UOADTrusteeName TO MRTrusteeName) AS rsUOMR" 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1

Set Rs = Com.Execute
While Not Rs.EOF
	dn = "LDAP://" & replace(rs.Fields("distinguishedName").Value,"/","\/")
	set objuser = getobject(dn)
	Set oSecurityDescriptor = objuser.Get("ntSecurityDescriptor")
	Set dacl = oSecurityDescriptor.DiscretionaryAcl
	Set ace = CreateObject("AccessControlEntry")
	objParentRS.addnew 
	objParentRS("UOADDisplayName") = rs.fields("displayname")
	objParentRS("UOADTrusteeName") = strdname & "\" & rs.fields("samaccountname")
	objParentRS.update
	Set objChildRS = objParentRS("rsUOMR").Value
	For Each ace In dacl
		   if lcase(ace.ObjectType) = "{ab721a54-1e2f-11d0-9819-00aa0040529b}" and ace.AceType = 5 then
			if ace.Trustee <> "NT AUTHORITY\SELF" and ace.AceFlags <> 6 then
				objChildRS.addnew
				objChildRS("MRmbox") = rs.fields("displayname")
				objChildRS("MRTrusteeName") = ace.Trustee
				objChildRS("MRRights") = "Send As"
				objChildRS("MRAceflags") = ace.AceFlags
				objChildRS.update
			end if
		   end if
		   if lcase(ace.ObjectType) = "{ab721a56-1e2f-11d0-9819-00aa0040529b}" and ace.AceType = 5 then
			if ace.Trustee <> "NT AUTHORITY\SELF" and ace.AceFlags <> 6 then
				objChildRS.addnew
				objChildRS("MRmbox") = rs.fields("displayname")
				objChildRS("MRTrusteeName") = ace.Trustee
				objChildRS("MRRights") = "Recieve As"
				objChildRS("MRAceflags") = ace.AceFlags
				objChildRS.update
			end if
		   end if
	Next
	rs.movenext
Wend
wscript.echo "Number of Mailboxes Checked " & objParentRS.recordcount
Wscript.echo
objParentRS.MoveFirst
Do While Not objParentRS.EOF
	Set objChildRS = objParentRS("rsUOMR").Value
	crec = 0
	if objChildRS.recordcount <> 0 then wscript.echo objParentRS("UOADDisplayName")
	Do While Not objChildRS.EOF 
		wscript.echo "   " & objChildRS.fields("MRmbox")
		wscript.echo "		-" & objChildRS.fields("MRRights")
		objChildRS.movenext
	loop
	objParentRS.MoveNext
loop