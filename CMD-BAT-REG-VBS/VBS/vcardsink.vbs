<SCRIPT LANGUAGE="VBScript">

public const mapiserver = "servername"
public const mapimailbox = "mailbox"

Sub ExStoreEvents_OnSave(pEventInfo, bstrURLItem, lFlags)
on error resume next

Const EVT_NEW_ITEM = 1 
Const EVT_IS_DELIVERED = 8 

If (lFlags And EVT_IS_DELIVERED) Then  

set objmessage = createobject("CDO.Message")
objmessage.datasource.open bstrURLItem
if objmessage.fields("DAV:contentclass").value = "urn:content-classes:message" then 
	Set objAttachments = objMessage.Attachments
	If objAttachments.Count <> 0 Then
		For Each objAttachment In objAttachments
               		fatt1 = len(objAttachment.filename)
                	fatt2 = fatt1 - 2
	        	attname = UCASE(objAttachment.filename)
	       		if lcase(mid(attname,fatt2,3)) = "vcf" then
				Set Strm = objAttachment.GetDecodedContentStream
	       			call ProcVcard(Strm.readtext,objmessage.fields("DAV:parentname").value)
				delmsg = 1
	    	        end if
   		 Next
	End If
End If
End if
set objmessage = nothing
if delmsg = 1 then
	set rec = createobject("ADODB.Record")
	rec.open bstrURLItem,,3
	rec.deleterecord
	set rec = nothing		
end if

End Sub

sub ProcVcard(vcardstream,pfPublicContactsFolder)
pfPublicContactsFolder = pfPublicContactsFolder & "/"
cphoto = 0
set contobj1 = createobject("CDO.Person")
set stm1 = contobj1.getvcardstream()
stm1.type = 2
stm1.Charset = "x-ansi"
stm1.writetext = vcardstream
stm1.flush
vcararry = split(vcardstream,vbcrlf)
for i = lbound(vcararry) to ubound(vcararry)
	if instr(vcararry(i),"PHOTO;")then
		cphoto = 1
	else
		if cphoto = 1 then
			if instr(vcararry(i),"    ") then
				photovcard = photovcard & vcararry(i) & vbcrlf
			else
				cphoto =2
			end if			
		end if
	end if
next
Randomize   ' Initialize random-number generator.
rndval = Int((20000000000 * Rnd) + 1)
contname = pfPublicContactsFolder & day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".eml"
contobj1.fields("urn:schemas:mailheader:subject").value = contobj1.fileas
contobj1.fields.update
if contobj1.fields("urn:schemas:mailheader:subject").value = "" then 
else
contobj1.datasource.saveto contname 
set contobj1 = nothing
if cphoto = 2 then
	set objmessage = createobject("CDO.Message")
	objmessage.datasource.open contname,,3
	Set objbpart = objmessage.BodyPart.AddBodyPart
	Set Flds = objbpart.Fields
	Flds("urn:schemas:mailheader:content-type") = "image/jpeg"
	Flds("urn:schemas:mailheader:content-disposition") = "attachment;filename=ContactPicture.jpg"
	Flds("urn:schemas:mailheader:content-transfer-encoding") = "base64"
	Flds.Update
	set Stm = createobject("ADODB.Stream")
	Set Stm = objbpart.GetEncodedContentStream
	stm.type = 2
	Stm.writetext photovcard
	Stm.Flush
	Stm.Close
	Set fso = CreateObject("Scripting.FileSystemObject")
	if (fso.FileExists("c:\temp\ContactPicture.jpg")) Then
		fso.deletefile("c:\temp\ContactPicture.jpg")
	End If
	objbpart.savetofile "c:\temp\ContactPicture.jpg"
	objmessage.addattachment "c:\temp\ContactPicture.jpg"
	objmessage.fields("http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/0x00008015") = true
	objmessage.fields("http://schemas.microsoft.com/mapi/proptag/0x0037001E") = objmessage.fields("urn:schemas:contacts:fileas").value
	objmessage.fields.update
	objmessage.datasource.save
	eiEntryID = Octenttohex(objmessage.fields("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102").value)
	set objSession = CreateObject("MAPI.Session")
	Const Cdoprop1 = &H7FFF000B
	const Cdoprop2 = &H370B0003
	const Cdoprop3 = &HE210003
	strProfile = mapiserver & vbLf & mapimailbox
	objSession.Logon "",,, False,, True, strProfile
	Set objInbox = objSession.Inbox
	Set objInfoStore = objSession.GetInfoStore(objSession.Inbox.StoreID)
	Set objpubstore = objSession.InfoStores("Public Folders")
	set objmessage = objSession.getmessage(eiEntryID,objpubstore.ID)
	set objAttachments = objmessage.Attachments
	For Each objAttachment In objAttachments
	objAttachment.fields.add Cdoprop1,"True"
	objAttachment.fields(Cdoprop2).value = -1
Next
objmessage.update

end if
end if

end sub

Function Octenttohex(OctenArry) 
ReDim aOut(UBound(OctenArry)) 
For i = 1 to UBound(OctenArry) + 1 
if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
	aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
else
	aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
end if
Next 
Octenttohex = join(aOUt,"")
End Function 


</SCRIPT>