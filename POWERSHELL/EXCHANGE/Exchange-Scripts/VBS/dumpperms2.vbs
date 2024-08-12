mailboxname = wscript.arguments(0)
servername = wscript.arguments(1)


Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\permissionDump.csv",8,true)

wscript.echo "Processing " & mailboxname
Set objSession = CreateObject("MAPI.Session")
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
	if cstr(objCuser.address) <> cstr(GetACLEntryName(fldace.ID)) then
		if DispACERules(fldace) <> "None" then
			wfile.writeline(objSession.CurrentUser & "," & mailboxname & "," & "Root" & "," & GetACLEntryName(fldace.ID) & "," & DispACERules(fldace))
		end if
	end if
Next
Set CdoFolders = CdoFolderRoot.Folders
Set CdoFolder = CdoFolders.GetFirst
do while Not (CdoFolder Is Nothing)
	ACLObj.CDOItem = CdoFolder
	Set FolderACEs = ACLObj.ACEs
	For each fldace in FolderACEs
		if cstr(objCuser.address) <> cstr(GetACLEntryName(fldace.ID)) then 
			if DispACERules(fldace) <> "None" then
				wfile.writeline(objSession.CurrentUser & "," & mailboxname & "," & CdoFolder.Name & "," & GetACLEntryName(fldace.ID) & "," & DispACERules(fldace))
			end if
		end if
	Next
	Set CdoFolder = CdoFolders.GetNext
loop
if Not objSession Is Nothing Then objSession.Logoff 
set objSession = nothing


Function GetACLEntryName(ACLEntryID)
select case ACLEntryID

	case "ID_ACL_DEFAULT"
		GetACLEntryName = "Default"
	case  "ID_ACL_ANONYMOUS"
		GetACLEntryName = "Anonymous"
	case else
		Set tmpEntry = objSession.GetAddressEntry(ACLEntryID)
		tmpName = tmpEntry.address
		GetACLEntryName = tmpName
end select

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