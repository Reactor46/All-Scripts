servername = "servername"
mailbox = "mailbox"
pubContactsfolderid = "000...etc"
set objSession = CreateObject("MAPI.Session")
strProfile = servername & vbLf & mailbox
objSession.Logon "",,, False,, True, strProfile
Set objcontactfolder = objSession.getdefaultfolder(5)
Set objInfoStore = objSession.GetInfoStore(objSession.Inbox.StoreID)
Set objpubstore = objSession.InfoStores("Public Folders")
set objpubContactsfolder = objSession.getfolder(pubContactsfolderid,objpubstore.id)
for each objcontact in objpubContactsfolder.messages
	set objCopyContact = objcontact.copyto(objcontactfolder.ID,objInfoStore.ID)
	objCopyContact.Unread = false
	objCopyContact.Update
 	Set objCopyContact = Nothing
	wscript.echo objcontact.subject
next
objsession.Logoff




