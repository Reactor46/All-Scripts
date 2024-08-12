mapiserver = "servername"
mapimailbox = "mailbox"
contactEntryID = "000000002F7EDBE600F33A4D9FD823919AF30EFF0700A65D419E23216B4AB84E3BF83385A6E8000000002F0300004ED86BAFBD2F39498EB9D5B4E38DA684000001594A040000"
filename = "c:\temp\ContactPicture.jpg"
set objSession = CreateObject("MAPI.Session")
Const Cdoprop1 = &H7FFF000B
const Cdoprop2 = &H370B0003
const Cdoprop3 = &HE210003
strProfile = mapiserver & vbLf & mapimailbox
objSession.Logon "",,, False,, True, strProfile
Set objInbox = objSession.Inbox
Set objInfoStore = objSession.GetInfoStore(objSession.Inbox.StoreID)
set objmessage = objSession.getmessage(contactEntryID)
set objAttachments = objmessage.Attachments
Set objAttachment = objmessage.Attachments.Add
objAttachment.Position = -1
objAttachment.Type = 1
objAttachment.Source = filename
objAttachment.ReadFromFile filename
For Each objAttachment In objAttachments
if objAttachment.name = "ContactPicture.jpg" then
    objAttachment.fields.add Cdoprop1,"True"
    objAttachment.fields(Cdoprop2).value = -1
    wscript.echo objAttachment.name
end if
objmessage.fields.add "0x8015",11,"true","0420060000000000C000000000000046"
objmessage.update

