mailboxname = wscript.arguments(0)
servername = wscript.arguments(1)
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\mbxforwardingRules.csv",8,true)
On Error resume next
Set objSession   = CreateObject("Redemption.RDOSession")
objSession.LogonExchangeMailbox mailboxname, servername
set objCuser = objSession.CurrentUser
unDisplayname = objCuser.Name
Set CdoInfoStore = objSession.Stores.DefaultStore
Set CdoFolderRoot = CdoInfoStore.IPMRootFolder
if err.number <> 0 Then
	wscript.echo "logon Error"
	wscript.echo err.description
	err.clear
else
On Error goto 0
wscript.echo 
wscript.echo "Mailbox Name :" & MailboxName



addcount = 1
frule = false
Set mrMailboxRules = objSession.Stores.DefaultStore.Rules
Wscript.echo "Checking Rules"
fwFirstWrite = 0
bnum = 0
for Each roRule in mrMailboxRules
	agrstr = ""
	acActType = ""
	rname = ""
        set actions = roRule.Actions 
	for i = 1 to actions.count
		acActType = actions(i).ActionType
		if acActType = 6 Or acActType = 8 Or acActType = 7 Then	
		    frule = true
		    If acActType = 8 Then
				rname = "Delegate-Forward-Rule"
				bnum = bnum + 1
		    else
				rname = "Forward-Rule"
		    End if
		    for each aoAdressObject In actions(i).Recipients			
				If agrstr = "" then
					if instr(aoAdressObject.Address,"/ou=") then 
						agrstr = agrstr & aoAdressObject.Name
					else
						agrstr = agrstr & aoAdressObject.Address
					end if
				Else 
					if instr(aoAdressObject.Address,"/ou=") then 
						agrstr = agrstr & ";" & aoAdressObject.Name
					else
						agrstr = agrstr & ";" & aoAdressObject.Address
					end if 
				End if
		   next
		   redim Preserve resarray(1,1,1,1,addcount)
		   resarray(1,0,0,0,addcount) = mailboxname
		   resarray(1,1,0,0,addcount) = rname
	           resarray(1,1,1,0,addcount) = roRule.ConditionsAsSQL	
		   resarray(1,1,1,1,addcount) = agrstr	
	           addcount = addcount + 1		
		end If
	next
	set actions = nothing
Next
if frule = true then
	for r = 1 to ubound(resarray,5)
		wfile.writeline(resarray(1,0,0,0,r) & "," & resarray(1,1,0,0,r) & "," & resarray(1,1,1,0,r)  & "," & resarray(1,1,1,1,r))
	next
end if
objSession.Logoff() 
End if
wfile.close





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

Function DisplayActionType(acActionType)
	Select Case acActionType
		Case 1 DisplayActionType = "Move-Rule"
		Case 2 DisplayActionType = "Assign Cateogry"
		Case 3 DisplayActionType = "Delete-Message"
		Case 4 DisplayActionType = "Delete-Permanently"
		Case 5 DisplayActionType = "Copy-Rule"
		Case 6 DisplayActionType = "Forward-Rule"
		Case 7 DisplayActionType = "Forward-As-Attachment"
		Case 8 DisplayActionType = "Delegate-Forward-Rule"
		Case 9 DisplayActionType = "ServerReply "
		Case 11 DisplayActionType = "Mark-Defer"
		Case 15 DisplayActionType = "Importance"
		Case 16 DisplayActionType = "Sensitivity"
		Case 19 DisplayActionType = "Mark-Read"
		Case 28 DisplayActionType = "Defer"
		Case 1024 DisplayActionType = "Bounce-Message"
		Case 1025 DisplayActionType = "Tag"
		Case Else DisplayActionType = "Unknown"
	End Select
End Function 