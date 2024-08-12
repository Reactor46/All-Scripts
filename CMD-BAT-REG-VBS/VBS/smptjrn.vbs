<SCRIPT LANGUAGE="VBSCRIPT">

Const cdoRunNextSink = 0
Const cdoSkipRemainingSinks = 1

Sub ISMTPOnArrival_OnArrival(ByVal objMessage, EventStatus)
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    set wfile = fso.opentextfile("c:\RECPT2.txt",8,true)
    wfile.writeline("Sink fired")
    wfile.writeline(objMessage.EnvelopeFields("http://schemas.microsoft.com/cdo/smtpenvelope/pickupfilename"))
    wfile.writeline(objMessage.EnvelopeFields("http://schemas.microsoft.com/cdo/smtpenvelope/senderemailaddress"))
    Dim objAttachments	
    Dim objAttachment	
    Dim objFields
    if objMessage.EnvelopeFields("http://schemas.microsoft.com/cdo/smtpenvelope/pickupfilename") = "" then
        	fname = "c:\Program Files\Exchsrvr\Mailroot\vsi 1\PickUp\" & day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".eml"
		objMessage.fields("urn:schemas:mailheader:x-sender:") = "internaluser@yourdomain.com.au"
		objMessage.fields("urn:schemas:mailheader:x-receiver") = "externaluser@yourdomain.com.au"
		objmessage.fields.update
		set stm = objMessage.getstream()
		stm.savetofile fname
		wfile.writeline("no pickup")
		set stm = nothing
		Set objFields = Nothing
		Set objAttachments = Nothing
		Set objAttachment = Nothing
   end if
   wfile.close
   EventStatus = cdoRunNextSink
End Sub

</SCRIPT>

