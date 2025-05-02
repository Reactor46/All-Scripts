''Logon Script by group
'===============================================================================================
'<CUSTOMER> Network Login Script
'===============================================================================================

'=================================================
' Set Environment Variables
'=================================================
Set WSHNetwork = WScript.CreateObject("WScript.Network")
Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")

On Error Resume Next

Domain = WSHNetwork.UserDomain
UserName = ""

While UserName = ""
	UserName = WSHNetwork.UserName
	MyGroups = GetGroups(Domain, UserName)
Wend


'=================================================
'DISCONNECTS ALL DEFINED MAPPED DRIVES
'NOTE:  DO NOT INCLUDE T: DRIVE!!!!!!!!
'TEMPORARY CODE TO CLEAN UP ALL USERS DRIVE MAPPINGS
'=================================================
If INGROUP ("<GROUPNAME>") Then  'CAN USE DOMAIN USERS IF NECESSARY
	WSHNetwork.RemoveNetworkDrive "F:", True, True
	WSHNetwork.RemoveNetworkDrive "G:", True, True
	'WSHNetwork.RemoveNetworkDrive "H:", True, True  'HOMEDRIVE IS REMARKED OUT SO AS NOT TO UNMAP IT!!!
	WSHNetwork.RemoveNetworkDrive "I:", True, True
	WSHNetwork.RemoveNetworkDrive "J:", True, True
	WSHNetwork.RemoveNetworkDrive "K:", True, True
	WSHNetwork.RemoveNetworkDrive "L:", True, True
	WSHNetwork.RemoveNetworkDrive "M:", True, True
	WSHNetwork.RemoveNetworkDrive "N:", True, True
	WSHNetwork.RemoveNetworkDrive "O:", True, True
	WSHNetwork.RemoveNetworkDrive "P:", True, True
	WSHNetwork.RemoveNetworkDrive "Q:", True, True
	WSHNetwork.RemoveNetworkDrive "R:", True, True
	WSHNetwork.RemoveNetworkDrive "S:", True, True
	WSHNetwork.RemoveNetworkDrive "T:", True, True
	WSHNetwork.RemoveNetworkDrive "U:", True, True
	WSHNetwork.RemoveNetworkDrive "V:", True, True
	WSHNetwork.RemoveNetworkDrive "W:", True, True
	WSHNetwork.RemoveNetworkDrive "X:", True, True
	WSHNetwork.RemoveNetworkDrive "Y:", True, True
	WSHNetwork.RemoveNetworkDrive "Z:", True, True
End If

'=================================================
'GIVES PC TIME TO DISCONNECT MAPPED DRIVES
'=================================================
WScript.Sleep 300


'=================================================
'Map Drives by Group
'=================================================
'USAGE:	MapDrive "X:", "\\SERVER\SHARE", "Drive Name"
'NOTE: <HOMEDRIVE:> IS NOT TO BE MAPPED AS IT IS THE HOME DRIVE!!!

objShell.NameSpace("<HOMEDRIVE:>").Self.Name = UCase(UserName) & " Home Drive"

'=================
'<GROUPNAME1> Drives:
'=================
If INGROUP ("<GROUPNAME1>") Then
	MapDrive "<DRIVE:>", "<SHARE>", "<DRIVE NAME>"
	MapDrive "<DRIVE:>", "<SHARE>", "<DRIVE NAME>"
End If

'=================
'<GROUPNAME2> Drives:
'=================
If INGROUP ("<GROUPNAME2>") Then
	MapDrive "<DRIVE:>", "<SHARE>", "<DRIVE NAME>"
	MapDrive "<DRIVE:>", "<SHARE>", "<DRIVE NAME>"
End If


'=================================================
'Map Drives by User
'=================================================
'USAGE:	MapDrive "X:", "\\SERVER\SHARE", "Drive Name"
'NOTE: <HOMEDRIVE:> IS NOT TO BE MAPPED AS IT IS THE HOME DRIVE!!!

'=================
'<USERNAME's> Drives:
'=================
If UCase(UserName) = "<USERNAME>" Then
	MapDrive "<DRIVE:>", "<SHARE>", "<DRIVE NAME>"
	MapDrive "<DRIVE:>", "<SHARE>", "<DRIVE NAME>"
End If


'=================================================
'Map Printers by Group
'=================================================
'USAGE:	MapPrinter "\\SERVER\PRINTER SHARE", "True"
'USAGE:	MapPrinter "\\SERVER\PRINTER SHARE", "False"
'NOTE: Drivers must be installed on pc/server running this login script or user must have permissions to install driver!

'=================
'<GROUPNAME1> Printers
'=================
If INGROUP ("<GROUPNAME1>") Then
	MapPrinter "<PRINTERSHARE>", "<DEFAULT TRUE/FALSE>"
	MapPrinter "<PRINTERSHARE>", "<DEFAULT TRUE/FALSE>"
End If

'=================
'<GROUPNAME2> Printers
'=================
If INGROUP ("<GROUPNAME2>") Then
	MapPrinter "<PRINTERSHARE>", "<DEFAULT TRUE/FALSE>"
	MapPrinter "<PRINTERSHARE>", "<DEFAULT TRUE/FALSE>"
End If


'=================================================
'Map Printers by User
'=================================================
'USAGE:	MapPrinter "\\SERVER\PRINTER SHARE", "True"
'USAGE:	MapPrinter "\\SERVER\PRINTER SHARE", "False"
'NOTE: Drivers must be installed on pc/server running this login script or user must have permissions to install driver!

'=================
'<USERNAME> Printers
'=================
If UCase(UserName) = "<USERNAME>" Then
	MapPrinter "<PRINTERSHARE>", "<DEFAULT TRUE/FALSE>"
	MapPrinter "<PRINTERSHARE>", "<DEFAULT TRUE/FALSE>"
End If


'=================================================
'Run THINREG Based upon Group Membership
'=================================================
'NOTE: This requires THINREG.EXE to be in the NETLOGON SHARE!
If INGROUP ("<GROUPNAME1>") Then
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE1.EXE>"" /Q")
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE2.EXE>"" /Q")
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE3.EXE>"" /Q")
End If

If INGROUP ("<GROUPNAME2>") Then
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE1.EXE>"" /Q")
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE2.EXE>"" /Q")
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE3.EXE>"" /Q")
End If


'=================================================
'Run THINREG Based upon User
'=================================================
'NOTE: This requires THINREG.EXE to be in the NETLOGON SHARE!
If UCase(UserName) = "<USERNAME>" Then
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE1.EXE>"" /Q")
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE2.EXE>"" /Q")
	WSHShell.Exec("""%LOGONSERVER%\NETLOGON\THINREG.EXE"" ""<UNC or Drive\path>\<EXECUTABLE3.EXE>"" /Q")
End If


'=================================================
'Exit Script
'=================================================
WScript.Quit

'===============================================================================================
'Subfunctions and Routines
'===============================================================================================

'=================================================
'Function: GetGroups
'=================================================
Function GetGroups(Domain, UserName)
	Set objUser = GetObject("WinNT://" & Domain & "/" & UserName)
	GetGroups=""
	For Each objGroup In objUser.Groups
		GetGroups=GetGroups & "[" & UCase(objGroup.Name) & "]"
	Next
End Function

'=================================================
'Function: InGroup
'=================================================
Function InGroup(strGroup)
	InGroup=False
	If InStr(MyGroups,"[" & UCase(strGroup) & "]") Then
		InGroup=True
	End If
End Function

'=================================================
' MapDrives Subroutine
'=================================================
Sub MapDrive(sDrive,sShare,sName)
	On Error Resume Next
	WSHNetwork.RemoveNetworkDrive sDrive, 1, 1
	WScript.Sleep 300
	Err.Clear
	WSHNetwork.MapNetworkDrive sDrive, sShare, 0
	objShell.NameSpace(sDrive).Self.Name = sName
End Sub

'=================================================
' MapPrinters Subroutine
'=================================================
Sub MapPrinter(sPrinterPath,sPrinterDefault)
	On Error Resume Next
	WSHNetwork.AddWindowsPrinterConnection sPrinterPath
	WScript.Sleep 300
	If sPrinterDefault = "1" Or sPrinterDefault = UCase("TRUE") Then
		WSHNetwork.SetDefaultPrinter sPrinterPath
	End If
End Sub

''' End Logon script by group