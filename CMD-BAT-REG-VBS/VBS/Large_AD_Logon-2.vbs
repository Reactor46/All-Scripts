'************************************************************************
'Login Script
'Created 5-19-10
'Logon.vbs
'Version 2.7
'Jeff Berndsen
'************************************************************************

Option Explicit
On Error Resume Next

'***************************************
'Declare Variables
'***************************************

'Main logon script variables
Dim WSHShell
Dim WSHNetwork
Dim strDomain
Dim strWinDir
Dim strUser
Dim strComputer
Dim UserObj
Dim objDrives
Dim i
Dim GroupObj
Dim strPrinter1
Dim strPrinter2
Dim strPrinter3
Dim objFSO
Const DESKTOP = &H10& 'Used for creation of shortcut on desktop
Dim objAppShell
Dim objFolder
Dim strFilePath
Dim strLogonServer

'Logging specific variables
Dim strLogFileName
Dim strLogFolderName
Dim strLogPath
Dim objLogFolderName
Dim objLogFileName
Dim objLogTextFile
Dim strMyDate
Const ForAppending = 8

'Variables for drive mapping
Dim strDriveU
Dim strDriveUPath
Dim strDriveZ
Dim strDriveZPath
Dim strDriveG
Dim strDriveGPath
Dim strDriveX
Dim strDriveXPath
Dim strDriveV
Dim strDriveVPath
Dim strDriveH
Dim strDriveHPath

'Variables to set drives as mapped
'Set them to false right away
Dim strDriveUMapped
Dim strDriveZMapped
Dim strDriveGMapped
Dim strDriveXMapped
Dim strDriveVMapped
Dim strDriveHMapped

strDriveUMapped = False
strDriveZMapped = False
strDriveGMapped = False
strDriveXMapped = False
strDriveVMapped = False
strDriveHMapped = False

Set WSHShell = CreateObject("WScript.Shell")
Set WSHNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objAppShell = CreateObject("Shell.Application")
Set objFolder = objAppShell.Namespace(DESKTOP)

'***************************************
'Find the logon server name
'***************************************
strLogonServer = WSHShell.ExpandEnvironmentStrings("%LOGONSERVER%")

'***************************************
'Get the user name
'***************************************
strUser = WSHNetwork.UserName

'***************************************
'Get the computer name
'***************************************
strComputer = WSHNetwork.ComputerName

'***************************************
'Pull the date into a variable and
'replace the / date separator with _
'***************************************
strMyDate = Date
strMyDate = Replace(strMyDate,"/","_")

'***************************************
'Set variables for drive letters and
'paths
'***************************************
strDriveU = "U:"
strDriveUPath = "\\server\share" & strUser
strDriveZ = "Z:"
strDriveZPath = "\\server\share"
strDriveG = "G:"
strDriveGPath = "\\server\share"
strDriveX = "X:"
strDriveXPath = "\\server\share"
strDriveV = "V:"
strDriveVPath = "\\server\share"
strDriveH = "H:"
strDriveHPath = "\\server\share"

'***************************************
'Set printer variables to printers
'***************************************
strPrinter1 = "\\server\printer"
strPrinter2 = "\\server\printer"
strPrinter3 = "\\server\printer"

'*************************************
'Folder path for the log file
'*************************************
strLogFolderName = "\\server\share\" & strUser

'*************************************
'File Name of the log
'*************************************
strLogFileName = "\" & strMyDate & ".log"

'*************************************
'Full path for the log file
'*************************************
strLogPath = strLogFolderName & strLogFileName

'*************************************
'Check that the strLogFolderName
'folder exists 
'*************************************
If objFSO.FolderExists(strLogFolderName) Then
Set objLogFolderName = objFSO.GetFolder(strLogFolderName)
	Else
		'**********************************
		'Create the folder
		'**********************************
		Set objLogFolderName = objFSO.CreateFolder(strLogFolderName)
		'WScript.Echo "Just created " & strLogFolderName
End If

'*************************************
'Check that the strLogFileName file
'exists
'*************************************
If objFSO.FileExists(strLogFolderName & strLogFileName) Then
Set objLogFolderName = objFSO.GetFolder(strLogFolderName)
	Else
		'**********************************
		'Create the file
		'**********************************
		Set objLogFileName = objFSO.CreateTextFile(strLogFolderName & strLogFileName, True)
		objLogFileName.Close 'Log file has to be closed before it can be appended to
		'WScript.Echo "Just created " & strLogFileName & " in the directory " & strLogFolderName
End If

'*************************************
'Open the text file for appending
'*************************************
Set objLogTextFile = objFSO.OpenTextFile(strLogPath, ForAppending, True)

'***************************************
'Get the user's domain name
'***************************************
strDomain = WSHNetwork.UserDomain

'***************************************
'Start logging
'***************************************
objLogTextFile.WriteLine("*************************************************************")
objLogTextFile.WriteLine("[" & Date & "]" & " " & "[" & Time & "]")
objLogTextFile.WriteLine("<<<Logon Script Started>>>")
objLogTextFile.WriteLine("*************************************************************" & _
	 vbCrLf)
objLogTextFile.WriteLine("Authenticating Server: " & strLogonServer)
objLogTextFile.WriteLine("Computer: " & strComputer)
objLogTextFile.WriteLine("Username: " & strDomain & "\" & strUser)


'***************************************
'Get the path of the Windows
'Directory
'***************************************
strWinDir = WSHShell.ExpandEnvironmentStrings("%WinDir%")
objLogTextFile.WriteLine("Windows directory located at: " & strWinDir)

'***************************************
'If shortcut icon is on the DESKTOP
'this part of the script deletes it
'Also logs what happened
'***************************************
strFilePath = objFolder.Self.Path & "\"
If objFSO.FileExists(strFilePath & "shortcut.lnk") Then
	objFSO.DeleteFile(strFilePath & "shortcut.lnk")
	If Err.Number <> 0 Then
		objLogTextFile.WriteLine("---------------------------------------------------------")
		objLogTextFile.WriteLine("*** Error while attempting to delete the shortcut ***")
	Else
		objLogTextFile.WriteLine("---------------------------------------------------------")
		objLogTextFile.WriteLine("*** shortcut.lnk was sucessfully removed from the desktop ***")
	End If
End If

'***************************************
'Bind to the user object to check for
'group memberships later
'***************************************
Set UserObj = GetObject("WinNT://" & strDomain & "/" & strUser)

'***************************************
'Enumerate mapped network drives and
'remove currently mapped drives and log
'the results
'***************************************
Set objDrives = WSHNetwork.EnumNetworkDrives

objLogTextFile.WriteLine("---------------------------------------------------------")
objLogTextFile.WriteLine("Network drives removed: ")

For i = 0 To objDrives.Count -1 Step 2
If Not objDrives.Item(i) = "" Then
	WSHNetwork.RemoveNetworkDrive objDrives.Item(i), True, True
	objLogTextFile.WriteLine(objDrives.Item(i))
End If
Next
objLogTextFile.WriteLine("---------------------------------------------------------")

'***************************************
'Give the script some time
'***************************************
WScript.Sleep 300

'***************************************
'Map the drives and printers that are 
'needed by all users (I.E. HOME DRIVE)
'and log rhe results
'***************************************

If strDriveUMapped = False Then
	WSHNetwork.MapNetworkDrive strDriveU, strDriveUPath, True

	If Err.Number <> 0 Then
		LogDriveMappingErrors Err.Number, Err.Description, strDriveU, strDriveUPath
	Else
		objLogTextFile.WriteLine("Drive U: was sucessfully mapped to " & strDriveUPath)
		strDriveUMapped = True
	End If
Else
	objLogTextFile.WriteLine("**** Drive U: is already mapped to " & strDriveUPath)
End If
Err.Clear

WSHNetwork.AddWindowsPrinterConnection strPrinter1

If Err.Number <> 0 Then
	LogPrinterMappingErrors Err.Number, Err.Description, strPrinter1
Else
	objLogTextFile.WriteLine("***** Printer " & strPrinter1 & " sucessfully added")
End If
Err.Clear

'***************************************
'Set Printer as default
'***************************************
WSHNetwork.SetDefaultPrinter strPrinter1

'***************************************
'Check for group memberships and map
'appropriate drives and printers
'WILL ONLY CHECK GLOBAL DOMAIN GROUPS
'Also logs the user's assigned groups
'***************************************
objLogTextFile.WriteLine("---------------------------------------------------------")
objLogTextFile.WriteLine("Active directory groups: " & vbCrLf)
For Each GroupObj In UserObj.Groups

'***************************************
'Comment this if you do not want to 
'list all groups the user belongs to
'***************************************
objLogTextFile.WriteLine(GroupObj.Name)
Next

'***************************************
'Force uppercase comparison of the 
'group names (All groups have to be
'named in CAPITAL LETTERS)
'Maps drives and printers based On
'Active Directory Groups and log
'the results
'***************************************
For Each GroupObj In UserObj.Groups

Select Case UCase(GroupObj.Name)
	Case "DOMAIN ADMINS"
		objLogTextFile.WriteLine("---------------------------------------------------------")
		objLogTextFile.WriteLine("**** Member of the [Domain Admins] group, started mapping drives... ****")
		
		If strDriveZMapped = False Then
			WSHNetwork.MapNetworkDrive strDriveZ, strDriveZPath, True
			If Err.Number <> 0 Then
				LogDriveMappingErrors Err.Number, Err.Description, strDriveZ, strDriveZPath
			Else
			objLogTextFile.WriteLine("**** Drive Z: was sucessfully mapped to " & strDriveZPath)
			strDriveZMapped = True
			End If
		Else
			objLogTextFile.WriteLine("**** Drive Z: is already mapped")
		End If
		Err.Clear
		
		WSHNetwork.AddWindowsPrinterConnection strPrinter1
		
		If Err.Number <> 0 Then
			LogPrinterMappingErrors Err.Number, Err.Description, strPrinter1
		Else
			objLogTextFile.WriteLine("***** Printer " & strPrinter1 & " sucessfully added")
		End If
		Err.Clear
		
	Case "GROUP2"
		objLogTextFile.WriteLine("---------------------------------------------------------")
		objLogTextFile.WriteLine("**** Member of the [GROUP2] group, started mapping drives... ****")
		
		If strDriveGMapped = False Then
			WSHNetwork.MapNetworkDrive strDriveG, strDriveGPath, True
			If Err.Number <> 0 Then
				LogDriveMappingErrors Err.Number, Err.Description, strDriveG, strDriveGPath
			Else
				objLogTextFile.WriteLine("**** Drive G: was sucessfully mapped to " & strDriveGPath)
				strDriveGMapped = True
			End If
		Else
			objLogTextFile.WriteLine("**** Drive G: is already mapped")
		End If
		Err.Clear
			
		If strDriveXMapped = False Then
			WSHNetwork.MapNetworkDrive strDriveX, strDriveXPath, True
			If Err.Number <> 0 Then
				LogDriveMappingErrors Err.Number, Err.Description, strDriveX, strDriveXPath
			Else
				objLogTextFile.WriteLine("**** Drive X: was sucessfully mapped to " & strDriveXPath)
				strDriveXMapped = True
			End If
		Else
			objLogTextFile.WriteLine("**** Drive X: is already mapped")
		End If
		Err.Clear
		
		If strDriveVMapped = False Then
			WSHNetwork.MapNetworkDrive strDriveV, strDriveVPath, True
			If Err.Number <> 0 Then
				LogDriveMappingErrors Err.Number, Err.Description, strDriveV, strDriveVPath
			Else
				objLogTextFile.WriteLine("**** Drive V: was sucessfully mapped to " & strDriveVPath)
				strDriveVMapped = True
			End If
		Else
			objLogTextFile.WriteLine("**** Drive V: is already mapped")
		End If
		Err.Clear
		
		WSHNetwork.AddWindowsPrinterCOnnection strPrinter1
		
		If Err.Number <> 0 Then
			LogPrinterMappingErrors Err.Number, Err.Description, strPrinter1
		Else
			objLogTextFile.WriteLine("***** Printer " & strPrinter1 & " sucessfully added")
		End If
		Err.Clear
		
		WSHNetwork.AddWindowsPrinterConnection strPrinter2
		If Err.Number <> 0 Then
			LogPrinterMappingErrors Err.Number, Err.Description, strPrinter2
		Else
			objLogTextFile.WriteLine("***** Printer " & strPrinter2 & " sucessfully added")
		End If
		Err.Clear
		
		WSHNetwork.SetDefaultPrinter strPrinter1
		If Err.Number <> 0 Then
			objLogTextFile.WriteLine("***** Error setting default printer to " & strPrinter1)
		Else
			objLogTextFile.WriteLine("***** Printer " & strPrinter1 & " sucessfully set as default")
		End If
		
	'Users must be added to this group to have the icon added to their desktop
	Case "ICON_GROUP"
		objLogTextFile.WriteLine("---------------------------------------------------------")
		objLogTextFile.WriteLine("**** Member of the [ICON_GROUP] group, adding shortcut to desktop... ****")
		
		strFilePath = objFolder.Self.Path & "\"
		'Shortcut must be added to the domain controllers in the below path for this to work
		objFSO.CopyFile strLogonServer & "\NETLOGON\shortcuts\shortcut.lnk", strFilePath, True
		If Err.Number <> 0 Then
			objLogTextFile.WriteLine("*** Error copying shortcut.lnk from the netlogon share ***")
		Else
			objLogTextFile.WriteLine("*** Sucessfully copied icon to the desktop ***")
		End If
	
	Case Else
		'Nothing to Do
End Select
Next
objLogTextFile.WriteLine(VbCrLf & "---------------------------------------------------------" & _
	vbCrLf)

'***************************************
'Map drives and printer connections
'based on user name (All user names
'have to be named in CAPITAL LETTERS)
'***************************************
Select Case UCase(strUser)
	Case "USERNAME"
		objLogTextFile.WriteLine("Started mapping specific drives for user: " & _
			strDomain & "\" & strUser)
		
		If strDriveYMapped = False Then
			WSHNetwork.MapNetworkDrive strDriveY, strDriveYPath, True
			If Err.Number <> 0 Then
				LogDriveMappingErrors Err.Number, Err.Description, strDriveY, strDriveYPath
			Else
				objLogTextFile.WriteLine("**** Drive Y: was sucessfully mapped to " & strDriveYPath)
				strDriveYMapped = True
			End If
		Else
			objLogTextFile.WriteLine("**** Drive Y: is already mapped")
		End If
		Err.Clear

		WSHNetwork.AddWindowsPrinterConnection strPrinter3
		
		If Err.Number <> 0 Then
			LogPrinterMappingErrors Err.Number, Err.Description, strPrinter3
		Else
			objLogTextFile.WriteLine("***** Printer " & strPrinter3 & " sucessfully added")
		End If
		Err.Clear
		
	Case Else
		'Nothing to Do
End Select

'***************************************
'Create desktop shortcut for all users
'***************************************
'Once again, the files have to be added to the domain controllers for this to work.
strFilePath = objFolder.Self.Path & "\"
objFSO.CopyFile strLogonServer & "\NETLOGON\shortcuts\link.url", strFilePath, True
objFSO.CopyFile strLogonServer & "\NETLOGON\shortcuts\shortcut.lnk", strFilePath, True

'***************************************
'Log the date and time the logon
'script finished
'***************************************
objLogTextFile.WriteLine(VbCrLf & "*************************************************************")
objLogTextFile.WriteLine("[" & Date & "]" & " " & "[" & Time & "]")
objLogTextFile.WriteLine("<<<Logon Script Finished>>>")
objLogTextFile.WriteLine("*************************************************************")

Sub LogDriveMappingErrors(errNum, errDesc, driveLetter, drivePath)
	objLogTextFile.WriteLine(VbCrLf & "---------------------------------------------------------")
	objLogTextFile.WriteLine("***** Error: " & errNum & VbCrLf & "Description: " & errDesc & " *****")
	objLogTextFile.WriteLine("This error occurred while mapping the " & driveLetter & " drive to " & drivePath)
	objLogTextFile.WriteLine(VbCrLf & "---------------------------------------------------------")
End Sub

Sub LogPrinterMappingErrors(errNum, errDesc, PrinterPath)
	objLogTextFile.WriteLine(VbCrLf & "---------------------------------------------------------")
	objLogTextFile.WriteLine("***** Error: " & errNum & VbCrLf & "Description: " & errDesc & " *****")
	objLogTextFile.WriteLine("This error occurred while mapping " & PrinterName & " to " & PrinterPath)
	objLogTextFile.WriteLine(VbCrLf & "---------------------------------------------------------")
End Sub



Option Explicit
Dim WshShell
Dim objNetwork, strUNCPrinter

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegWrite "HKCU\Printers\Settings\Wizard\Set As Default", 0, "REG_DWORD"

strUNCPrinter = "\\PRINTSERVER1\Ricoh Laser"
Set objNetwork = CreateObject("WScript.Network") 
objNetwork.AddWindowsPrinterConnection strUNCPrinter

' Here is where we set the default printer to strUNCPrinter
objNetwork.SetDefaultPrinter strUNCPrinter
WScript.Echo "Check the Printers folder for : " & strUNCPrinter

WScript.Quit

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPorts =  objWMIService.ExecQuery _
    ("Select * from Win32_TCPIPPrinterPort Where Name = 'IP_10.0.0.11'")

For Each objPort in colInstalledPorts 
    objPort.Delete_
Next

Set objNewPort = objWMIService.Get _
    ("Win32_TCPIPPrinterPort").SpawnInstance_

objNewPort.Name = "IP_10.0.0.22"
objNewPort.Protocol = 1
objNewPort.HostAddress = "10.0.0.22"
objNewPort.PortNumber = "9999"
objNewPort.SNMPEnabled = False
objNewPort.Put_

