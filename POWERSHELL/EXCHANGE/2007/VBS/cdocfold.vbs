Const PR_FOLDER_WEBVIEWINFO = &H36DF0102
propval = "020000000100000001000000000000000000000000000000000000000" _
& "0000000000000000000000034000000680074007400700073003A002F002F006C" _
& "006F00670069006E002E0070006F007300740069006E0069002E0063006F006D000000"

servername = "servername"
mailboxname = "mailbox"

Set objSession   = CreateObject("MAPI.Session")

objSession.Logon "","",false,true,true,true,servername & vbLF & mailboxname
Set objInbox = objSession.Inbox
Set CdoInfoStore = objSession.GetInfoStore
Set CdoFolderRoot = CdoInfoStore.RootFolder
Set CdoFolders = CdoFolderRoot.Folders

bFound = False
Set CdoFolder = CdoFolders.GetFirst
Do While (Not bFound) And Not (CdoFolder Is Nothing)
    If CdoFolder.Name = "Junk E-mail" Then
       bFound = True
    Else
       Set CdoFolder = CdoFolders.GetNext
    End If
Loop
Set ActionFolder = CdoFolder
Set CdoFolder = ActionFolder.Folders.GetFirst
Do While (Not cFound) And Not (CdoFolder Is Nothing)
    If CdoFolder.Name = "postini" Then
       cFound = True
    Else
       Set CdoFolder = ActionFolder.Folders.GetNext
    End If
Loop
if cFound <> True then
	set piniFld = ActionFolder.Folders.add("postini")	
	piniFld.fields.Add PR_FOLDER_WEBVIEWINFO, propval
	piniFld.Update 
	wscript.echo "Folder Created"
else
	wscript.echo "Folder Exists"
end if 

