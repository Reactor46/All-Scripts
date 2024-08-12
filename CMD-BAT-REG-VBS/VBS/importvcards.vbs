public const pfPublicContactsFolder = "file://./backofficestorage/domain.com/public folders/pubcontacts/"
public const vcardfolder = "c:\vcardimport"
public const mapiserver = "servername"
public const mapimailbox = "mailbox"
set fso = createobject("Scripting.FileSystemObject")
set f = fso.getfolder(vcardfolder)
Set fc = f.Files
For Each file1 in fc
call ProcVcard(file1.path)
next
wscript.echo "Done"

sub ProcVcard(fname)
set stm = createobject("ADODB.Stream")
cphoto = 0
stm.open
stm.type = 2
stm.Charset = "x-ansi"
stm.loadfromfile fname
set contobj1 = createobject("CDO.Person")
set stm1 = contobj1.getvcardstream()
stm1.type = 2
stm1.Charset = "x-ansi"
txtvar = stm.readtext
stm1.writetext = txtvar
stm1.flush
vcararry = split(txtvar,vbcrlf)
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
wscript.echo contobj1.fileas
Randomize   ' Initialize random-number generator.
rndval = Int((20000000000 * Rnd) + 1)
contname = pfPublicContactsFolder & day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".eml"
contobj1.fields("urn:schemas:mailheader:subject").value = contobj1.fileas
contobj1.fields.update
if contobj1.fields("urn:schemas:mailheader:subject").value = "" then 
	wscript.echo "Skipping Blank Entry"   
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
