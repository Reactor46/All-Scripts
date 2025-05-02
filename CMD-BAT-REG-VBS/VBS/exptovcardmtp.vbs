Set objSystemInfo = CreateObject("ADSystemInfo") 
set iper = createobject("CDO.Person")
strdname = objSystemInfo.DomainShortName
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Query = "<LDAP://" & strNameingContext & ">;(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(|(&(objectCategory=person) (objectClass=user)(msExchHomeServerName=*)) )))));samaccountname,displayname,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = Query
Com.Properties("Page Size") = 1000
Set Rs = Com.Execute
While Not Rs.EOF
	iper.datasource.open "LDAP://" & rs.fields("distinguishedName")
	Set strm = iper.GetvCardStream
	strm.savetofile "c:\temp\" & rs.fields("samaccountname") & ".vcf"
	rs.movenext
Wend
Wscript.echo "Done"