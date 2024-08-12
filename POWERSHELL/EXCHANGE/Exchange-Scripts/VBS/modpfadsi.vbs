pfnamemail = wscript.arguments(0)
customvalue = wscript.arguments(1)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
pfQuery = "<LDAP://" & strNameingContext & ">;(&(&(&(& (mailnickname=*) (| (objectCategory=publicFolder) )))(objectCategory=publicFolder)(mail=" & pfnamemail & ")));name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = pfQuery
Set Rs = Com.Execute
While Not Rs.EOF
	set objuser = getobject("LDAP://" & rs.fields("distinguishedName"))
	objuser.extensionAttribute1 = customvalue
	objuser.setinfo
	wscript.echo "Modified folder " & objuser.displayname & " Added attribute " & objuser.extensionAttribute1
	rs.movenext
Wend