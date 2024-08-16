Public Const CdoDefaultFolderContacts = 5
snServername = wscript.arguments(0)
mbMailboxName = wscript.arguments(1)
set csCDOSession = CreateObject("MAPI.Session")
pfProfile = snServername & vbLf & mbMailboxName
csCDOSession.Logon "","",False,True,0,True, pfProfile
set cfContactsFolder = csCDOSession.getdefaultfolder(CdoDefaultFolderContacts)
set cfContactscol = cfContactsFolder.messages
set ofConFilter = cfContactscol.Filter
Set cfContFltFld1 = ofConFilter.Fields.Add("0x8015",vbBoolean,true,"0420060000000000C000000000000046")
For Each ctContact In cfContactscol
	Set collAttachments = ctContact.Attachments 
	For Each atAttachment In collAttachments
		If atAttachment.name = "ContactPicture.jpg" Then
			fname = replace(replace(replace(replace(replace((ctContact.subject & "-" & atAttachment.name),":","-"),"\",""),"/",""),"?",""),chr(34),"")
			fname = replace(replace(replace(replace(replace(replace(fname,"<",""),">",""),chr(11),""),"*",""),"|",""),"(","")
			fname = replace(replace(replace(fname,")",""),chr(12),""),chr(15),"")
			atAttachment.WriteToFile("c:\contactpictures\" & fname)
			wscript.echo "Exported Picture to : " &  fname
		End if		
	next
Next