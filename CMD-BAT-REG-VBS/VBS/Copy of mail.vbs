Set SendMsg = CreateObject("CDO.Message")

With SendMsg
	.Subject = "Job Slot Errors"
	.From = "Scheduler <scheduler@creditone.com>"
	.To = "stephen.bonar@creditone.com"
	.TextBody = "Argent jobs running on Lasargent02 have experienced jobslot errors.  The .js1 files have been recreated.  Please verify jobs are running normally and release dependencies on jobs affected by the errors.  See attached Log."
	.TextBody2 = "TEST2"
rem	.AddAttachment "D:\Argent Jobslot Check\Jobslot.log"
	.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mailgateway.fnbm.corp"
	.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	.Configuration.Fields.Update
	.Send
End With