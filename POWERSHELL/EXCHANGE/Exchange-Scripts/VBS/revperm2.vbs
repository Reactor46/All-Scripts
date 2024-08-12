on error resume next
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString	
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("C:\revperms.csv",2,true)
wfile.writeline("Username,Mailbox,FolderName,Permission")
Query = "<LDAP://" & strNameingContext & ">;(&(&(&(& (mailnickname=*)" & _
	"(!msExchHideFromAddressLists=TRUE) (| (&(objectCategory=person)(objectClass=user)" & _
	"(|(homeMDB=*)(msExchHomeServerName=*))) )))))" & _
	";samaccountname,legacyExchangeDN,msExchHomeServerName,displayname,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = Query
Com.Properties("Page Size") = 1000
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS UOADDisplayName, " & _
           "  NEW adVarChar(255) AS UOADTrusteeName, " & _
           "  NEW adVarChar(255) AS UOADLegacyDN, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS MRmbox, " & _
           "      NEW adVarChar(255) AS MRFolder, " & _
           "      NEW adVarChar(255) AS MRTrusteeName, " & _
           "      NEW adVarChar(255) AS MRRights) " & _
           "      RELATE UOADLegacyDN TO MRTrusteeName) AS rsUOMR" 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1

Set Rs = Com.Execute
While Not Rs.EOF
	objParentRS.addnew 
	objParentRS("UOADDisplayName") = rs.fields("displayname")
	objParentRS("UOADTrusteeName") = rs.fields("samaccountname")
	objParentRS("UOADLegacyDN") = rs.fields("legacyExchangeDN")
	objParentRS.update
	Set objChildRS = objParentRS("rsUOMR").Value
	inplinearray = Split(rs.fields("msExchHomeServerName").value, "=", -1, 1)
	CDOAddr = getpermissions(inplinearray(ubound(inplinearray)),rs.fields("samaccountname").value,rs.fields("displayname").value)
	rs.movenext
Wend
wscript.echo "Number of Mailboxes Checked " & objParentRS.recordcount
Wscript.echo

objParentRS.MoveFirst
Do While Not objParentRS.EOF
	Set objChildRS = objParentRS("rsUOMR").Value
	if objChildRS.recordcount <> 0 then wscript.echo objParentRS("UOADDisplayName")
	Do While Not objChildRS.EOF
		wscript.echo "   " & objChildRS.fields("MRmbox") & _
			" "  & objChildRS.fields("MRFolder") & " " &  objChildRS.fields("MRRights")
		wfile.writeline objParentRS("UOADDisplayName") & "," & objChildRS.fields("MRmbox") & _
			","  & objChildRS.fields("MRFolder") & "," &  objChildRS.fields("MRRights")
		objChildRS.movenext
	loop
	objParentRS.MoveNext
loop

function getpermissions(servername,mailboxname,displayname)
wscript.echo "Processing " & mailboxname
Set objSession   = CreateObject("MAPI.Session")
objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
if err.number <> 0 then wscript.echo err.description
err.clear
set objCuser = objSession.CurrentUser
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set ACLObj = CreateObject("MSExchange.aclobject")
ACLObj.CDOItem = CdoFolderRoot
Set FolderACEs = ACLObj.ACEs
For each fldace in FolderACEs
	if cstr(objCuser.address) <> cstr(GetACLEntryName(fldace.ID, objSession)) then
		objChildRS.addnew
		objChildRS("MRmbox") = mailboxname
		objChildRS("MRFolder") = "Root"
		objChildRS("MRTrusteeName") = GetACLEntryName(fldace.ID, objSession)
		objChildRS("MRRights") = DispACERules(fldace)
		objChildRS.update
	end if
Next
Set CdoFolders = CdoFolderRoot.Folders
Set CdoFolder = CdoFolders.GetFirst
do while Not (CdoFolder Is Nothing)
	ACLObj.CDOItem = CdoFolder
	Set FolderACEs = ACLObj.ACEs
	For each fldace in FolderACEs
		if cstr(objCuser.address) <> cstr(GetACLEntryName(fldace.ID, objSession)) then 
			objChildRS.addnew
			objChildRS("MRmbox") = mailboxname
			objChildRS("MRFolder") = CdoFolder.Name
			objChildRS("MRTrusteeName") = GetACLEntryName(fldace.ID, objSession)
			objChildRS("MRRights") = DispACERules(fldace)
			objChildRS.update
		end if
	Next
	Set CdoFolder = CdoFolders.GetNext
loop
if Not objSession Is Nothing Then objSession.Logoff 
set objSession = nothing
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