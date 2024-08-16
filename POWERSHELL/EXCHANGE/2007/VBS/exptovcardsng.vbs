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
vcardwrite = ""
While Not Rs.EOF
	iper.datasource.open "LDAP://" & rs.fields("distinguishedName")
	Set strm = iper.GetvCardStream
	vcardwrite = vcardwrite & strm.readtext & vbcrlf
	rs.movenext
Wend
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\export.vcf",2,true)
wfile.writeline vcardwrite
wfile.close
set wfile = nothing