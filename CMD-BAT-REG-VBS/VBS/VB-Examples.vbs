6otGXhdDYirbVWj5Gp!BH$



###Move a folder

Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFolder "C:\Scripts" , "M:\HelpDesk\Management"

###Modify Folder Attributes
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFolder = objFSO.GetFolder("C:\FSO") 
 
If objFolder.Attributes = objFolder.Attributes AND 2 Then 
    objFolder.Attributes = objFolder.Attributes XOR 2  
End If 

###Copy Folders and Files with Overwrite
Const OverWriteFiles = TRUE 
 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
objFSO.CopyFolder "C:\Scripts" , "C:\FSO" , OverWriteFiles 

###Create a FolderParentFolder = "C:\"  
 
set objShell = CreateObject("Shell.Application") 
set objFolder = objShell.NameSpace(ParentFolder)  
objFolder.NewFolder "Archive" 


###Rename a Folder
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colFolders = objWMIService.ExecQuery _ 
    ("Select * from Win32_Directory where name = 'c:\\Scripts'") 
 
For Each objFolder in colFolders 
    errResults = objFolder.Rename("C:\Script Repository") 
Next 

###Rename all files in a folder with specific extension
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set FileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\test'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In FileList
    strDate = Left(objFile.CreationDate, 8)
    strNewName = objFile.Drive & objFile.Path & _
       strDate & "." & "jpg"
    strNameCheck = Replace(strNewName, "\", "\\")

    i = 1
    Do While True
        Set colFiles = objWMIService.ExecQuery _
            ("Select * from Cim_Datafile Where Name = '" & strNameCheck & "'")
        If colFiles.Count = 0 Then
            errResult = objFile.Rename(strNewName)
            Exit Do
        Else
            i = i + 1
            strNewName = objFile.Drive & objFile.Path & _
                strDate & "" & i & "." & "jpg"
            strNameCheck = Replace(strNewName, "\", "\\")
        End If
    Loop
Next



###Modify Access Permissions Based on Folder Size
set objFSO = CreateObject("Scripting.FileSystemObject") 
set objFolder = objFSO.GetFolder("D:\ProjectOperations") 
 
Dim objF 
 
 objF = Int(objFolder.size/1048576) 
 Wscript.Echo objF & " MB" 
 
 set oShell = Wscript.CreateObject("Wscript.Shell") 
 sfolder = "I:\Project Operations" 
 if objF > 10 then 
  oShell.Run "%COMSPEC% /c echo y| cacls D:\ProjectOperations /p administrator:f Operations:r" 
 end if  
 if objF < 10 then 
  oShell.Run "%COMSPEC% /c echo y| cacls D:\ProjectOperations /p administrator:f Operations:c" 
  end if 
 
 set oShell = Nothing 
 
 ###Verify the Existence of a Folder
 'DoesTheFolderExist 
'Tells if a folder exists on the hard drive 
'This takes an exact Folder Path without the final slash 
'Usage: cscript DoesTheFolderExist.vbs "F:\WINNT" 
 
set fsobj=WScript.CreateObject("Scripting.FileSystemObject") 
If fsobj.FolderExists(WScript.Arguments(0)) Then 
   WScript.Echo "Folder '" & WScript.Arguments(0) & "' exists."  
End If 
If  Not fsobj.FolderExists(WScript.Arguments(0)) Then 
   WScript.Echo "Folder '" & WScript.Arguments(0) & "' not exists."  
End If 

## Monitor Folder
strComputer =  'TODO: Enter the servers name here
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate, (Security)}!\" & _
strComputer & "rootcimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
("Select * from __instancecreationevent where " _
& "TargetInstance isa 'Win32_NTLogEvent' " _
& "and TargetInstance.EventCode = '560' " ) 'TODO: modify the event code to fire on what ever you require

Do
Set objLatestEvent = colMonitoredEvents.NextEvent
strAlertToSend = objLatestEvent.TargetInstance.User _
& " has accessed a folder on a server"   'TODO: Modify the alert you would like to receive
Wscript.Echo strAlertToSend

Set objEmail = CreateObject("CDO.Message")

objEmail.From = 'TODO: Specify a from address
objEmail.To = 'TODO: Enter a To address
objEmail.Subject = strAlertToSend
objEmail.Textbody = strAlertToSend
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
'TODO: Specify your mail servers name here
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
Loop

##Monitor Folder for files larger than XXX
'FolderMonitor.vbs - 1.0a, 10/01/10
'
'Rob Dunn 
'
'I put this together a few years ago, and updated it today for a user 
' on the Spiceworks forums.
'
'This script will check a folder (share) and email you if any files in the
' folder are larger than the byte size designated in the watch.ini file.
' If so, it will attach any files to the email if blnAttachFiles = true and 
' send it to the recipient listed in the 'SendTo' INI value.
'
' Reads INI file for settings - here's the format (name it 'watch.ini' in the same folder):
'
'By the way - FileSizeThreshold number is in byte format

'[Main]
'WatchServers="server"
'WatchDirectory="\sharename"
'FileSizeThreshold="200"
'DontSendFilesLargerThan="400"
'
'[Email]
'SMTPServer="x.x.x.x"
'SendFrom="folderwatcher_noreply@yourdomain.com"
'SendTo="you@yourdomain.com"                                    
'Subject="Folder Contents"
'CC=""

Dim blnAttachFiles, iDontSendFilesLargerThan, sMessage

Const ForReading = 1
Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing", _
      cdoSendUsingPort = 2, _
      cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"

'//  Create the CDO connections.
Dim iMsg, iConf, Flds, blnSendMail, debug, strDate

debug = 0

blnAttachFiles = true
blnSendMail = false           
            
strWatchDir = GetINIString("Main", "WatchDirectory", "", ".\watch.ini")
strWatchServers = GetINIString("Main", "WatchServers", "", ".\watch.ini")
strCompare = GetINIString("Main", "Comparison", "", ".\watch.ini")
strFileSize = Int(GetINIString("Main", "FileSizeThreshold", "", ".\watch.ini"))
iDontSendFilesLargerThan = Int(GetINIString("Main", "DontSendFilesLargerThan", "", ".\watch.ini"))
strSMTPServer = GetINIString("Email", "SMTPServer", "", ".\watch.ini")
strSendTo = GetINIString("Email", "SendTo", "", ".\watch.ini")
strSendFrom = GetINIString("Email", "SendFrom", "", ".\watch.ini")
strSubject = GetINIString("Email", "Subject", "", ".\watch.ini")


'If debug = 0 Then wscript.echo strWatchServers & vbcrlf & strWatchDir & vbcrlf & strFileSize _
' & vbcrlf & strSMTPServer & vbcrlf & strSendTo & vbcrlf _
' & strSendFrom & vbcrlf & strSubject & vbcrlf & strCompare


strWatchServers = split(strWatchServers,",", -1,1)
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

'//  SMTP server configuration.
With Flds
	.Item(cdoSendUsingMethod) = cdoSendUsingPort
    
	'//  Set the SMTP server address here.
    	.Item(cdoSMTPServer) = strSMTPServer
    	.Update
End With

For i=0 to UBound(strWatchServers)
  Set iMsg = CreateObject("CDO.Message")

  If debug = 1 Then wscript.echo strWatchServers(i)
  
  strWatchDir_a = "\\" & strWatchServers(i) & strWatchDir 
  

  
  '//  Set the message properties.
  With iMsg
    Set .Configuration = iConf
        .To       = strSendTo
        .From     = strSendFrom
        .Subject  = "Contents of " & strWatchDir_a
        .TextBody = strSubject & " - " & strWatchDir_a & " - " & Now
  End With
  If debug = 1 Then wscript.echo strWatchDir_a
  Call ShowFolderList(strWatchDir_a)

Next
                  
Function AddAttachment(strFile)
	'//  An attachment can be included.
	iMsg.AddAttachment strFile
End Function

Function SendMail
	'//  Send the message.
  iMsg.TextBody = iMsg.TextBody & vbcrlf & vbcrlf & sMessage
	iMsg.Send ' send the message.
 	Set imsg = nothing
End Function

Function ShowFolderList(folderspec)
   Dim fso, f, f1, fc, s
    
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)

   Set fc = f.Files
   For Each f1 in fc
	   if debug = 1 then wscript.echo f1.name & vbtab & f1.size
	   If f1.size > strFileSize Then
     	  blnSendMail = true
        s = s & f1.name & vbtab & f1.size
        s = s & vbcrlf
        'fso.CopyFile f1.path, folderspec & "\" & strDate & "\" & f1.name
        If blnAttachFiles = true and f1.size < iDontSendFilesLargerThan then  
          Call AddAttachment(f1.path)
        End If

        sMessage = sMessage & f1.name & vbtab & f1.size & " bytes" & vbnewline   
      End If
      
      if debug = 1 then wscript.echo sMessage
   Next
   

   ShowFolderList = s
   If debug = 1 Then wscript.echo s

   If CStr(err.number) <> 0 Then 
        Call ErrorHandle(folderspec)
   ElseIf blnSendMail = True Then
        Call SendMail
   End If
 
   blnSendMail = false
   
End Function

Sub ErrorHandle(folderspec)
   	WshShell.Popup "Error: Path on " & folderspec & " does not exist or no files " _ 
   	 & "were found.", 7, "No files found in " & folderspec, 64
End Sub

Sub WriteINIString(Section, KeyName, Value, FileName)
  Dim INIContents, PosSection, PosEndSection
  
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    Dim OldsContents, NewsContents, Line
    Dim sKeyName, Found
    OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
    OldsContents = split(OldsContents, vbCrLf)

    'Temp variable To find a Key
    sKeyName = LCase(KeyName & "=")

    'Enumerate section lines
    For Each Line In OldsContents
      If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
        Line = KeyName & "=" & Value
        Found = True
      End If
      NewsContents = NewsContents & Line & vbCrLf
    Next

    If isempty(Found) Then
      'key Not found - add it at the end of section
      NewsContents = NewsContents & KeyName & "=" & Value
    Else
      'remove last vbCrLf - the vbCrLf is at PosEndSection
      NewsContents = Left(NewsContents, Len(NewsContents) - 2)
    End If

    'Combine pre-section, new section And post-section data.
    INIContents = Left(INIContents, PosSection-1) & _
      NewsContents & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    'Section Not found. Add section data at the end of file contents.
    If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then 
      INIContents = INIContents & vbCrLf 
    End If
    INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
      KeyName & "=" & Value
  end if'if PosSection>0 Then
  WriteFile FileName, INIContents
End Sub


Function GetINIString(Section, KeyName, Default, FileName)
  Dim INIContents, PosSection, PosEndSection, sContents, Value, Found
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)
  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  
  If PosSection > 0 Then
    'Section exists. Find end of section
       
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1
    
    'Separate section contents
    sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

    If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
      Found = True
      'Separate value of a key.
      Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
    End If
  End If
  If isempty(Found) Then Value = Default
  
  GetINIString = replace(Value,Chr(34),"")
End Function

'Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function

'File functions
Function GetFile(ByVal FileName)
  Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
  'Go To windows folder If full path Not specified.
  If InStr(FileName, ":\") = 0 And Left (FileName,2) <> "\\" And Left (FileName,2) <> ".\" Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  'On Error Resume Next
End Function

'Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function

'File functions
Function GetFile(ByVal FileName)

  Set FS = CreateObject("Scripting.FileSystemObject")
  
  'Go To windows folder If full path Not specified.
  If InStr(FileName, ":\") = 0 And Left(FileName,2)<> "\\" And Left(FileName,2) <> ".\" Then 
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  On Error Resume Next

  GetFile = FS.OpenTextFile(FileName).ReadAll 
  'wscript.echo getfile
End Function

Sub WriteINIStringVirtual(Section, KeyName, Value, FileName)
  WriteINIString Section, KeyName, Value, _
    Server.MapPath(FileName)
End Sub

Function GetINIStringVirtual(Section, KeyName, Default, FileName)
  GetINIStringVirtual = GetINIString(Section, KeyName, Default, _
    Server.MapPath(FileName))
End Function 

Function MakeSureDirectoryTreeExists(dirName)
Dim aFolders, newFolder
	If debug = 1 Then wscript.echo "Makesuredirectorytreexists " & dirname
	dim delim
	' Creates the FSO object.
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Checks the folder's existence.
	If Not fso.FolderExists(dirName) Then

		' Splits the various components of the folder name.
		If instr(dirname,"\\") then
		    delim = "-_-_-_-"
			dirname = replace(dirname,"\\",delim)
			'wscript.echo dirname
		End if
		aFolders = split(dirName, "\")
		if instr(dirname,delim) Then
			dirname = replace(aFolders(0),delim,"\\")
			'wscript.echo "aFolders = " & dirname
		End if
		' Obtains the drive's root folder.
		
		newFolder = fso.BuildPath(dirname, "\")
	
		' Scans each component in the array, and create the appropriate folder.
		For i=1 to UBound(aFolders)
			newFolder = fso.BuildPath(newFolder, aFolders(i))

			If Not fso.FolderExists(newFolder) Then

				fso.CreateFolder newFolder
			

			End If
		Next
	End If
End Function


##Zip something

newzip "C:\test.zip"
Createzip "c:\test.zip", "c:\directorytozip\"
msgbox "Zip file created"

Sub NewZip(pathToZipFile)
 
   'WScript.Echo "Newing up a zip file (" & pathToZipFile & ") "
 
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   Dim file
   Set file = fso.CreateTextFile(pathToZipFile)
 
   file.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
 
   file.Close
   Set fso = Nothing
   Set file = Nothing
 
   WScript.Sleep 500
 
End Sub
 
Sub CreateZip(pathToZipFile, dirToZip)
 
   'WScript.Echo "Creating zip  (" & pathToZipFile & ") from (" & dirToZip & ")"
 
   Dim fso
   Set fso= Wscript.CreateObject("Scripting.FileSystemObject")
 
   pathToZipFile = fso.GetAbsolutePathName(pathToZipFile)
   dirToZip = fso.GetAbsolutePathName(dirToZip)
 
   If fso.FileExists(pathToZipFile) Then
       'WScript.Echo "That zip file already exists - deleting it."
       fso.DeleteFile pathToZipFile
   End If
 
   If Not fso.FolderExists(dirToZip) Then
       'WScript.Echo "The directory to zip does not exist."
       Exit Sub
   End If
 
   NewZip pathToZipFile
 
   dim sa
   set sa = CreateObject("Shell.Application")
 
   Dim zip
   Set zip = sa.NameSpace(pathToZipFile)
 
   'WScript.Echo "opening dir  (" & dirToZip & ")"
 
   Dim d
   Set d = sa.NameSpace(dirToZip)
 
   zip.CopyHere d.items, 4
 
   Do Until d.Items.Count <= zip.Items.Count
       Wscript.Sleep(200)
   Loop
 
End Sub

###Systematically remove old files

Dim fso, startFolder, OlderThanDate
 
Set fso = CreateObject("Scripting.FileSystemObject")
startFolder = "C:\New folder" ' folder to start deleting (subfolders will also be cleaned)
OlderThanDate = DateAdd("d", -30, Date)  ' 30 days (adjust as necessary)
 
DeleteOldFiles startFolder, OlderThanDate
 
Function DeleteOldFiles(folderName, BeforeDate)
   Dim folder, file, fileCollection, folderCollection, subFolder
 
   Set folder = fso.GetFolder(folderName)
   Set fileCollection = folder.Files
   For Each file In fileCollection
      If file.DateLastModified < BeforeDate Then
         fso.DeleteFile(file.Path)
      End If
   Next
 
    Set folderCollection = folder.SubFolders
    For Each subFolder In folderCollection
       DeleteOldFiles subFolder.Path, BeforeDate
    Next
End Function

### Files per folder -- HTML Output

Set objDlg = WScript.CreateObject("Shell.Application")
Set objFile = objDlg.BrowseForFolder(&H0,"Select the Directory to Review",&H0001 + &H0010)

objStartFolder = objfile.ParentFolder.ParseName(objfile.Title).Path 
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files
Set objHTML = objfso.CreateTextFile("c:\File_Output.html",True)

With objHTML
	.Write("<html>"+vbcrlf)
	.Write("<table border = " & """1""" & ">"+vbcrlf)
	.Write("<tr>"+vbcrlf)
	.Write("<td>Folder</td>"+vbcrlf)
	.Write("<td>File Name</td>"+vbcrlf)
	.write("</tr>"+vbcrlf)
	For Each Subfolder In OBJFolder.SubFolders
		For Each objFile in colFiles
			.Write("<tr>"+vbcrlf)
			.Write("<td>"+objfolder.Name+"</td>"+vbCrLf)
			.Write("<td>"+objfile.Name+"</td>"+vbcrlf)
			.Write("</tr>"+vbcrlf)
		Next
		Set objFolder = objFSO.GetFolder(Subfolder)
	Set colFiles = objFolder.Files
	Next	
	.Write("</table>"+vbcrlf)
	.Write("</html>"+vbcrlf)
End With

MsgBox "Process Complete"

Set SubFolder = Nothing
Set objHTML = Nothing
Set objFSO = Nothing
Set objFolder = Nothing 
Set colFiles = Nothing
Set objStartFolder = Nothing 