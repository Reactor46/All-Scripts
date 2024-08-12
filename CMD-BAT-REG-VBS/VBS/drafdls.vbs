Public Const CdoDefaultFolderSentItems = 3
Public Const CdoPR_DRAFTS_FOLDER = &H36D70102
Public const adVarChar = 200
set objSession = CreateObject("MAPI.Session")
strProfile = "servername" & vbLf & "mailbox"
objSession.Logon "",,, False,, True, strProfile
Set objInbox = objSession.Inbox
Set objInfoStore = objSession.GetInfoStore(objSession.Inbox.StoreID)
strDraftsEntryID = objInbox.Fields.Item(CdoPR_DRAFTS_FOLDER)
Set objDraftsFolder = objSession.GetFolder(strDraftsEntryID, Null)
set objSentItemsFolder = objSession.GetDefaultFolder(CdoDefaultFolderSentItems) 
set ColDraftfolders = objDraftsFolder.folders
bFound = False
Set CdoDraftsFolder = ColDraftfolders.GetFirst
Do While (Not bFound) And Not (CdoDraftsFolder Is Nothing)
    If CdoDraftsFolder.Name = "DL Drafts" Then
        bFound = True
   	set colDLDraftmsg = CdoDraftsFolder.messages
    Else
        Set CdoDraftsFolder = ColDraftfolders.GetNext
    End If
Loop
If bFound = False then
      Set CdoDraftsFolder = ColDraftfolders.Add("DL Drafts")
      Wscript.echo "Create DL Drafts Folder"
      set colDLDraftmsg = CdoDraftsFolder.messages
End If
Set objMessages = objSentItemsFolder.Messages 
for each objmessage in objMessages
	Set colRecips = objmessage.Recipients
	if colRecips.Count => 4 then
		wscript.echo objmessage.subject
		wscript.echo colRecips.Count
		strSubject = ""
		ifound = false
		set rs = createobject("ador.recordset")
		rs.fields.append "name", adVarChar, 255
		rs.fields.append "AddressType", adVarChar, 255
		rs.open
		for each objrecip in colRecips
			rs.addnew 
                		rs("name") = objrecip.name
                		rs("AddressType") = objrecip.Type
		        rs.Update
                next
		rs.sort = "AddressType,Name"
		rs.moveFirst
		do until rs.eof
 		    	select case rs.fields("AddressType").Value
				case 1 strSubject = strSubject & " TO:" & replace(rs.fields("name").Value,"'","")
				case 2 strSubject = strSubject & " CC:" & replace(rs.fields("name").Value,"'","")
				case 3 strSubject = strSubject & " BCC:" & replace(rs.fields("name").Value,"'","")
			end select
 		        rs.moveNext
		loop
		set rs = nothing
		for each objDLmessage in colDLDraftmsg
				if objmessage.Recipients.Count = objDLmessage.Recipients.Count then
					if strSubject = objDLmessage.Subject then
						ifound = True
						wscript.echo "Found Skipping"
						exit for
					end if
				end if
		next
		if ifound = false then 
			set objnewdlmsg = colDLDraftmsg.add
			set colnewdlmsgrecp = objnewdlmsg.Recipients
			for each objrecip in colRecips
				Set objOneRecip = colnewdlmsgrecp.Add(objrecip.name,objrecip.Address,objrecip.Type,objrecip.AddressEntry.ID)
				objnewdlmsg.Subject = strSubject
				set objOneRecip = nothing
			next	
			objnewdlmsg.update
			set objnewdlmsg = nothing
		end if
	
	end if
next





