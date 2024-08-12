snServername = wscript.arguments(0)
mbMailboxName = wscript.arguments(1)

Set objSession   = CreateObject("MAPI.Session")
Set catDict = CreateObject("Scripting.Dictionary")
objSession.Logon "","",false,true,true,true,snServername & vbLF & mbMailboxName
set ifInboxFolderCol = objSession.inbox.messages
set attFilter = ifInboxFolderCol.Filter
Set attFilterFiled = attFilter.Fields.Add(&h0E1B000B,true)
attFilter.TimeFirst = DateAdd("m",-1,Now()) 
For Each moMessageobject In ifInboxFolderCol 
	ceCatExists = False
	catDict.RemoveAll
	On Error Resume Next
	ccCurrentCats = moMessageobject.Fields.item("{2903020000000000C000000000000046}Keywords").value
	If Err.number = 0 Then 
		ceCatExists = True
		For Each existingcat In ccCurrentCats 
			catDict.add existingcat,1
		next
	End If
	On Error goto 0 
	oldcatlength = catDict.Count
	Call GetCategories(moMessageobject,catDict)
	If catDict.Count > oldcatlength Then
		wscript.echo moMessageobject.Subject
		ReDim newcats(catDict.Count-1)
		catkeys = catDict.Keys
		For i = 0 to catDict.Count-1
			newcats(i) = catkeys(i)
		Next
		If ceCatExists = True then
			moMessageobject.Fields.item("{2903020000000000C000000000000046}Keywords").value = newcats
		Else
			moMessageobject.Fields.add  "Keywords", vbArray , newcats, "2903020000000000C000000000000046" 
		End If
		moMessageobject.update
	End if
next
sub GetCategories(msgObject,catDict)
For Each attachment In msgObject.Attachments
	On Error Resume Next
	inline = 0
	fnFileName = attachment.fields(&h3704001E)
	Err.clear
	contentid = attachment.fields(&h3712001F)
	If Err.number = 0 Then
		inline = 1
	Else
		inline = 0
	End if
	Err.clear
	attflags = attachment.fields(&h37140003) 
	If  Err.number = 0 Then
		If attflags = 4 Then inline = 1
	End if
	If Len(fnFileName) > 4 And inline = 0  Then
		Select Case Right(LCase(fnFileName),4)
			Case ".doc" If Not catDict.exists("Word Attachment") Then
							catDict.add "Word Attachment",1
						End if
			Case ".ppt"  If Not catDict.exists("PowerPoint Attachment") Then
							catDict.add "PowerPoint Attachment",1
						End if
			Case ".xls"  If Not catDict.exists("Excel Attachment") Then
							catDict.add "Excel Attachment",1
						End if
			Case ".jpg" If Not catDict.exists("Image Attachment") Then
							catDict.add "Image Attachment",1
						End if
			Case ".bmp" If Not catDict.exists("Image Attachment") Then
							catDict.add "Image Attachment",1
						End if
			Case ".mov" If Not catDict.exists("Video Attachment") Then
							catDict.add "Video Attachment",1
						End if
			Case ".mpg" If Not catDict.exists("Video Attachment") Then
							catDict.add "Video Attachment",1
						End if
			Case ".wmv" If Not catDict.exists("Video Attachment") Then
							catDict.add "Video Attachment",1
						End if
			Case ".pdf" If Not catDict.exists("PDF Attachment") Then
							catDict.add "PDF Attachment",1
						End if
			Case ".mp3" If Not catDict.exists("Sound Attachment") Then
							catDict.add "Sound Attachment",1
						End if
			Case ".pps" If Not catDict.exists("PowerPoint Attachment") Then
							catDict.add "PowerPoint Attachment",1
						End if
			Case ".zip" If Not catDict.exists("Zip Attachment") Then
							catDict.add "Zip Attachment",1
						End if
		End select

	End if
	On Error goto 0

next

End sub
