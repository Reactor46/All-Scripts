servername = wscript.arguments(0)
UserToRemove = wscript.arguments(1)

Public Const CdoDefaultFolderCalendar = 0
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
usrQuery = "<LDAP://" & strDefaultNamingContext & ">;(&(&(objectCategory=person)(objectClass=user)(mail=" & UserToRemove & ")));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = usrQuery
Set usRs = Com.Execute
while not usRs.eof
	UserToRemoveLegDN = usRS.fields("legacyExchangeDN")
	usRs.movenext
wend
wscript.echo UserToRemoveLegDN
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mail,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		call dofreebusy(servername,rs1.fields("mailnickname"))
		wscript.echo "Seaching Permissions on: " & rs1.fields("mailnickname") 
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
wscript.echo "Done"


function dofreebusy(servername,mailboxname)

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
   if lcase(GetACLEntryName(fldace.id,objSession)) = lcase(UserToRemoveLegDN) then
	wscript.echo "Found ACL :" & GetACLEntryName(fldace.id,objSession)
   	FolderACEs.delete fldace.id
	ACLObj.Update
	Wscript.echo "ACL Deleted"
   end if
Next
'FreeBusy Update
Set objRoot = objSession.GetFolder("") 
Set objFreeBusyFolder = objRoot.Folders.Item("FreeBusy Data") 
ACLObj.CDOItem = objFreeBusyFolder
Set FolderACEs = ACLObj.ACEs
For each fldace in FolderACEs
   if lcase(GetACLEntryName(fldace.id,objSession)) = lcase(UserToRemoveLegDN) then
	wscript.echo "Found ACL :" & GetACLEntryName(fldace.id,objSession)
   	FolderACEs.delete fldace.id
	ACLObj.Update
	Wscript.echo "ACL Deleted"
   end if
Next
	
End function



Function GetACLEntryName(ACLEntryID,SubSession)

select case ACLEntryID
	case "ID_ACL_DEFAULT"
		GetACLEntryName = "Default"
	case  "ID_ACL_ANONYMOUS"
		GetACLEntryName = "Anonymous"
	case else
		Set tmpEntry = SubSession.GetAddressEntry(ACLEntryID)
		tmpName = tmpEntry.address
		GetACLEntryName = tmpName
end select
Set objSession = nothing
Set CdoFolderRoot = nothing
Set ACLObj = nothing
Set FolderACEs = nothing
Set objSession = nothing

End Function

