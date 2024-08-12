Const RIGHT_DS_DELETE = &H10000
Const RIGHT_DS_READ = &H20000
Const RIGHT_DS_CHANGE = &H40000
Const RIGHT_DS_TAKE_OWNERSHIP = &H80000
Const RIGHT_DS_MAILBOX_OWNER = &H1
Const RIGHT_DS_SEND_AS = &H2
Const RIGHT_DS_PRIMARY_OWNER = &H4

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
Query = "<LDAP://" & strNameingContext & ">;(&(objectCategory=person)(objectClass=user));samaccountname,displayname,distinguishedName;subtree"
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
	objParentRS.addnew 
	objParentRS("UOADDisplayName") = rs.fields("displayname")
	objParentRS("UOADTrusteeName") = strdname & "\" & rs.fields("samaccountname")
	objParentRS.update
	if not objuser.mailnickname = "" then
        	Set oSecurityDescriptor = objuser.Get("msExchMailboxSecurityDescriptor")
		Set dacl = oSecurityDescriptor.DiscretionaryAcl
		Set ace = CreateObject("AccessControlEntry")
		Set objChildRS = objParentRS("rsUOMR").Value
		For Each ace In dacl
			   if ace.AceFlags <> 18 then
				if ace.Trustee <> "NT AUTHORITY\SELF" then
					objChildRS.addnew
					objChildRS("MRmbox") = rs.fields("displayname")
					objChildRS("MRTrusteeName") = ace.Trustee
					objChildRS("MRRights") = ace.AccessMask
					objChildRS("MRAceflags") = ace.AceFlags
					objChildRS.update
				end if
		   	   end if
		Next
	end if
	rs.movenext
Wend
wscript.echo "Number of Mailboxes Checked " & objParentRS.recordcount
Wscript.echo
objParentRS.MoveFirst
Do While Not objParentRS.EOF
	Set objChildRS = objParentRS("rsUOMR").Value
	if objChildRS.recordcount <> 0 then wscript.echo objParentRS("UOADDisplayName")
	Do While Not objChildRS.EOF
		wscript.echo "   " & objChildRS.fields("MRmbox")
		If (objChildRS.fields("MRRights") And RIGHT_DS_SEND_AS) Then
			wscript.echo "		-send mail as"
		End If
		If (objChildRS.fields("MRRights") And RIGHT_DS_CHANGE) Then
			wscript.echo "		-modify user attributes"
		End If
		If (objChildRS.fields("MRRights") And RIGHT_DS_DELETE) Then
			wscript.echo  "		-delete mailbox store"
		End If
		If (objChildRS.fields("MRRights") And RIGHT_DS_READ) Then
			wscript.echo  "		-read permissions"
		End If
		If (objChildRS.fields("MRRights") And RIGHT_DS_TAKE_OWNERSHIP) Then
			wscript.echo  "		-take ownership of this object"
		End If
		If (objChildRS.fields("MRRights") And RIGHT_DS_MAILBOX_OWNER) Then
			wscript.echo "		-is mailbox owner of this object"
		End If
		If (objChildRS.fields("MRRights") And RIGHT_DS_PRIMARY_OWNER) Then
			wscript.echo  "		-is mailbox Primary owner of this object"
		End If
		objChildRS.movenext
	loop
	objParentRS.MoveNext
loop