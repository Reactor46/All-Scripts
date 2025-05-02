<SCRIPT LANGUAGE="VBScript">

Sub ExStoreEvents_OnSave(pEventInfo, bstrURLItem, lFlags)

call dbinsert(bstrURLItem)

End Sub

sub dbinsert(murl)
on error resume next
dbDatabaseName = "address@domain.com"
dbAttachmentsName = "Attachments@" & dbDatabaseName
dbDatabaseName = "inbox@" & dbDatabaseName
Set Cnxn1 = CreateObject("ADODB.Connection")
strCnxn1 = "Data Source=server;Initial Catalog=SharedMailbox;User Id=username;Password=pass;"
Cnxn1.Open strCnxn1
set msg = createobject("cdo.message") 
msg.datasource.open murl
eiEntryID = Octenttohex(msg.fields("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102").value)
miMessageID = msg.fields("urn:schemas:mailheader:message-id").value
dhDavhref = msg.fields("DAV:Href").value
stSenttime = msg.fields("urn:schemas:httpmail:datereceived").value
fnFromName =  msg.fields("urn:schemas:httpmail:fromname").value
feFromEmail =  replace(replace(msg.fields("urn:schemas:httpmail:fromemail").value,"<",""),">","")
toToEmail = msg.fields("urn:schemas:mailheader:to").value
sjSubject = msg.Subject
tbTextBody = msg.fields("urn:schemas:httpmail:htmldescription")
haHasAttach = msg.fields("urn:schemas:httpmail:hasattachment").value
line_to_insert = ("'" & eiEntryID & "','" & replace(miMessageID,"'","''") &  "','" & dhDavhref & "','" & stSenttime & "','" & replace(fnFromName,"'","''") & _
"','" &  replace(feFromEmail,"'","''") & "','" & replace(toToEmail,"'","''") & "','"  & replace(sjSubject,"'","''") & _
"','" & left(replace(tbTextBody,"'","''"),255) & "','" &  replace(tbTextBody,"'","''") & "','" & haHasAttach  &  "'")
sqlstate1 = "insert into [" & dbDatabaseName & "] values(" & line_to_insert & ")"
Cnxn1.Execute(sqlstate1)
i = 1
set objattachments = msg.attachments 
for each objattachment in objattachments 
if objAttachment.ContentMediaType = "message/rfc822" then
	set msg1 = createobject("cdo.message") 
	msg1.datasource.OpenObject objattachment, "ibodypart"
	fnFileName = msg1.subject & "(" & i & ")" & ".eml"
	ctContentType = "message/rfc822" 
	ceContentTransferEncoding = "7bit"
	cdContentDisposition = msg1.Fields("urn:schemas:mailheader:content-disposition").value
	set stm = msg1.getstream
	mbMessageBody =  stm.readtext
	line_to_insert = ("'" & eiEntryID & "','" & i  & "','" & replace(fnFileName,"'","''")  & "','" &  replace(ctContentType,"'","''") & _
	"','" &  replace(ceContentTransferEncoding,"'","''") & "','" & replace(cdContentDisposition,"'","''") & "','"  & replace(mbMessageBody,"'","''") & "'")
	sqlstate1 = "insert into [" & dbAttachmentsName & "] values(" & line_to_insert & ")"
	Cnxn1.Execute(sqlstate1)
else 
	fnFileName = objattachment.filename
	ctContentType = objattachment.ContentMediaType 
	ceContentTransferEncoding = objattachment.ContentTransferEncoding
	cdContentDisposition = objattachment.Fields("urn:schemas:mailheader:content-disposition").value
	set stm = objAttachment.getstream
	mbMessageBody =  stm.readtext
	line_to_insert = ("'" & eiEntryID & "','" & i  & "','" & replace(fnFileName,"'","''") & "','" & replace(ctContentType,"'","''") & _
	"','" &  replace(ceContentTransferEncoding,"'","''") & "','" & replace(cdContentDisposition,"'","''") & "','"  & replace(mbMessageBody,"'","''") & "'")
	sqlstate1 = "insert into [" & dbAttachmentsName & "] values(" & line_to_insert & ")"
	Cnxn1.Execute(sqlstate1)	
end if 
i = i + 1
next 
set msg = nothing 

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