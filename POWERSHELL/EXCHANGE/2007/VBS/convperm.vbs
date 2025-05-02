servername = wscript.arguments(0)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("C:\" & wscript.arguments(0) & ".csv",2,true)
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
ACLObj.CDOItem = CdoFolderRoot
Set FolderACEs = ACLObj.ACEs
getpermissions = mailboxname & ",Root,"
For each fldace in FolderACEs
   getpermissions = getpermissions & GetACLEntryName(fldace.ID, objSession) & "-" & DispACERules(fldace) & ", " 
Next
wfile.writeline getpermissions
Set CdoFolders = CdoFolderRoot.Folders
Set CdoFolder = CdoFolders.GetFirst
do while Not (CdoFolder Is Nothing)
	wscript.echo CdoFolder.Name 
	ACLObj.CDOItem = CdoFolder
	Set FolderACEs = ACLObj.ACEs
	getpermissions = mailboxname & "," & CdoFolder.Name  & ","
	For each fldace in FolderACEs
  		getpermissions = getpermissions & GetACLEntryName(fldace.ID, objSession) & "-" & DispACERules(fldace) & ", " 		
	Next
	wfile.writeline getpermissions	
	Set CdoFolder = CdoFolders.GetNext
loop

End function

Function GetACLEntryName(ACLEntryID,SubSession)

select case ACLEntryID
	case "ID_ACL_DEFAULT"
		GetACLEntryName = "Default"
	case  "ID_ACL_ANONYMOUS"
		GetACLEntryName = "Anonymous"
	case else
		Set tmpEntry = SubSession.GetAddressEntry(ACLEntryID)
		tmpName = tmpEntry.Name
		GetACLEntryName = tmpName
end select
Set objSession = nothing
Set CdoFolderRoot = nothing
Set ACLObj = nothing
Set FolderACEs = nothing
Set objSession = nothing

End Function

Function DispACERules(DisptmpACE)

Select Case DisptmpACE.Rights

        Case ROLE_NONE, 0  ' Checking in case the role has not been set on that entry.
                DispACERules = "None"
        Case 1024  ' Check value since ROLE_NONE is incorrect
                DispACERules = "None"
        Case ROLE_AUTHOR
                DispACERules = "Author"
        Case 1051  ' Check value since ROLE_AUTHOR is incorrect
                DispACERules = "Author"
        Case ROLE_CONTRIBUTOR
                DispACERules = "Contributor"
        Case 1026  ' Check value since ROLE_CONTRIBUTOR is incorrect
                DispACERules = "Contributor"
        Case 1147  ' Check value since ROLE_EDITOR is incorrect
                DispACERules = "Editor"
        Case ROLE_NONEDITING_AUTHOR
                DispACERules = "Nonediting Author"
        Case 1043  ' Check value since ROLE_NONEDITING AUTHOR is incorrect
                DispACERules = "Nonediting Author"
        Case 2043  ' Check value since ROLE_OWNER is incorrect
                DispACERules = "Owner"
        Case ROLE_PUBLISH_AUTHOR
                DispACERules = "Publishing Author"
        Case 1179  ' Check value since ROLE_PUBLISHING_AUTHOR is incorrect
                DispACERules = "Publishing Author"
        Case 1275  ' Check value since ROLE_PUBLISH_EDITOR is incorrect
                DispACERules = "Publishing Editor"
        Case ROLE_REVIEWER
                DispACERules = "Reviewer"
        Case 1025  ' Check value since ROLE_REVIEWER is incorrect
                DispACERules = "Reviewer"
        Case Else
                DispACERules = "Custom"
End Select

End Function
