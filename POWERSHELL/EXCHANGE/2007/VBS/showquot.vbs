set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("DefaultNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
atrQuery = "<LDAP://" & strNameingContext & ">;(&(&(mailnickname=*)(objectClass=User)(mDBUseDefaults=FALSE)));cn,name,samaccountname,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = atrQuery
Set snrs = Com.Execute
wscript.echo "Name,SamAccountName"
wscript.echo 
while not snrs.eof
	wscript.echo snrs.fields("name") & "," & snrs.fields("samaccountname")
	rem ResetQuota(snrs.fields("distinguishedName"))
	snrs.movenext
wend

sub ResetQuota(DN)

userDN = "LDAP://" & DN
set objuser = getobject(userDN)
objuser.mDBUseDefaults = True
objuser.setinfo
Wscript.echo "Modified User" & objuser.displayname


end sub