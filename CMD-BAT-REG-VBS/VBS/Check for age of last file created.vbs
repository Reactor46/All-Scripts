'Script to check when the last file was created, if it is over an hour ago, send an e-mail alert

Set fs = CreateObject("Scripting.FileSystemObject")
Set MainFolder = fs.GetFolder("***INSERT LOCATION TO CHECK HERE***")

'This section finds the most recently used folder in "Unarchived"

For Each fldr In MainFolder.SubFolders
    If fldr.DateLastModified > LastFolderDate Or IsNull(LastFolderDate) Then
        LastFolder = fldr.path
        LastFolderDate = fldr.DateLastModified
    End If
Next

'This section finds the most recently created file in the above folder

Set LastWorkingFolder = fs.GetFolder(LastFolder)

For Each file In LastWorkingFolder.files
    If file.DateCreated > LastFileDate Or IsNull(LastFileDate) Then
        LastFile = file.path
        LastFileDate = file.DateCreated
    End If
Next

'This section find the minute difference between the system time and the file located above

SystemDate = (Now)

FileAge = DateDiff("n", SystemDate, LastFileDate)
FileLimit = -30

If FileAge < FileLimit Then
	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = "***INSERT FROM EMAIL HERE***"
	objEmail.To = "***INSERT TO EMAIL HERE***"
	objEmail.Subject = "***INSERT SUBJECT HERE***" 
	objEmail.Textbody = "***INSERT TEXT HERE***"
	objEmail.Configuration.Fields.Item _
    		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item _
    		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
        	"***INSERT EXCHANGE SERVER HERE***" 
	objEmail.Configuration.Fields.Item _
    		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEmail.Send
End If