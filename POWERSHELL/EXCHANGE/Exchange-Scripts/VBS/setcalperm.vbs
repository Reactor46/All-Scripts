Public Const CdoDefaultFolderCalendar = 0
servername = "servername"
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("C:\" & servername & ".csv",2,true)
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		strperm = getpermissions(servername,rs1.fields("mailnickname"))
		'wscript.echo strperm
		wfile.writeline strperm
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
wfile.close
set fso = nothing
set conn = nothing
set com = nothing
wscript.echo "Done"


function getpermissions(servername,mailboxname)

Set objSession   = CreateObject("MAPI.Session")
objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set ACLObj = CreateObject("MSExchange.aclobject")
set cdocalendar = objSession.GetDefaultFolder(CdoDefaultFolderCalendar)
ACLObj.CDOItem = cdocalendar
Set FolderACEs = ACLObj.ACEs
getpermissions = mailboxname & ": "
For each fldace in FolderACEs
   if fldace.ID = "ID_ACL_DEFAULT"  then
	fldace.rights = 1025
	ACLObj.update
   end if
   getpermissions = getpermissions & fldace.ID & "-" & fldace.rights
Next

End function

