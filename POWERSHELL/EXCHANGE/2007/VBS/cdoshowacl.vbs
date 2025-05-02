servername = wscript.arguments(0)
mailboxname = wscript.arguments(1)
Set objSession   = CreateObject("MAPI.Session")
objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set ACLObj = CreateObject("MSExchange.aclobject")
ACLObj.CDOItem = CdoFolderRoot
Set FolderACEs = ACLObj.ACEs
For each fldace in FolderACEs
   wscript.echo mailboxname & "," & GetACLEntryName(fldace.ID, objSession) & "," & DispACERules(fldace)
Next


Function GetACLEntryName(ACLEntryID,SubSession)

select case ACLEntryID
	case "ID_ACL_DEFAULT"
		GetACLEntryName = "Default"
	case  "ID_ACL_ANONYMOUS"
		GetACLEntryName = "Anonymous"
	case else
		on error resume next
		Set tmpEntry = SubSession.GetAddressEntry(ACLEntryID)
		tmpName = tmpEntry.Name
		GetACLEntryName = tmpName
end select



End Function

Function DispACERules(DisptmpACE)

Select Case DisptmpACE.Rights

        Case ROLE_NONE, 0  ' Checking in case the role has not been set on that entry.
                DispACERules = "None"
        Case ROLE_AUTHOR
                DispACERules = "Author"
        Case ROLE_CONTRIBUTOR
                DispACERules = "Contributor"
        Case 1147  ' Check value since ROLE_EDITOR is incorrect
                DispACERules = "Editor"
        Case ROLE_NONEDITING_AUTHOR
                DispACERules = "Nonediting Author"
        Case 2043  ' Check value since ROLE_OWNER is incorrect
                DispACERules = "Owner"
        Case ROLE_PUBLISH_AUTHOR
                DispACERules = "Publishing Author"
        Case 1275  ' Check value since ROLE_PUBLISH_EDITOR is incorrect
                DispACERules = "Publishing Editor"
        Case ROLE_REVIEWER
                DispACERules = "Reviewer"
        Case Else
                DispACERules = "Custom"
End Select

End Function
