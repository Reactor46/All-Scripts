' ===========================================================================================
' Check If SQL Server is Installed on the Computer
' Check If 32-Bit OR 64-Bit
' Check If OS Is Windows 2000 OR Windows 2003 OR Windows 2008 OR Windows 2008 R2
' ===========================================================================================

Option Explicit

Dim ObjWMIService, ObjFSO, ObjColServices, WshShell, ColFiles
Dim StrComputer, StrFileName, StrPath, ChkSQLInstall, Counter
Dim OurOSBit, StrName, StrVersion, ArrNames, ObjFile
Dim StrOSName, StrSQLName

StrComputer = "."
ChkSQLInstall = False:	FindIfSQLPresent
CheckOSVersion
If ChkSQLInstall = True Then
	GetOSBit:	CheckSQLVersion	
End If
If ChkSQLInstall = True And OurOSBit = "X86" Then
	' -- On 32-Bit SQL possibilities are: SQL 2000, SQL 2005 and SQL 2008
	' -- On 32-Bit OS possibilities are: Win2000, Win2003, Win2003 R2, Win2008
	If StrComp(StrSQLName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2000", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 32-Bit Windows Server 2000"
	End If
	If StrComp(StrOSName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 32-Bit Windows Server 2003"
	End If
	If StrComp(StrOSName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 32-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 32-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2000", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 32-Bit Windows Server 2000"
	End If	
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 32-Bit Windows Server 2003"
	End If	
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 32-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 32-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2000", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 32-Bit Windows Server 2000"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 32-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 32-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 32-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 32-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 32-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 32-Bit Windows Server 2008"
	End If
End If

If ChkSQLInstall = True And OurOSBit = "X64" Then
	' -- On 64-Bit SQL possibilities are: SQL 2008, SQL 2008 R2 and SQL 2012
	' -- On 64-Bit OS possibilities are: Win2003, Win2003 R2, Win2008, Win2008 R2
	If StrComp(StrSQLName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 64-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 64-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 64-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2000", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2000 Installed on 64-Bit Windows Server 2008 R2"
	End If
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 64-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 64-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 64-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2005", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2005 Installed on 64-Bit Windows Server 2008 R2"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 64-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 64-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 64-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2008", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 Installed on 64-Bit Windows Server 2008 R2"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 64-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 64-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 64-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2008 R2", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2008 R2 Installed on 64-Bit Windows Server 2008 R2"
	End If
	If StrComp(StrSQLName, "SQL 2012", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2012 Installed on 64-Bit Windows Server 2003"
	End If
	If StrComp(StrSQLName, "SQL 2012", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2003 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2012 Installed on 64-Bit Windows Server 2003 R2"
	End If
	If StrComp(StrSQLName, "SQL 2012", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2012 Installed on 64-Bit Windows Server 2008"
	End If
	If StrComp(StrSQLName, "SQL 2012", vbTextCompare) = 0 AND StrComp(StrOSName, "Windows 2008 R2", vbTextCompare) = 0 Then
		WScript.Echo "SQL Server 2012 Installed on 64-Bit Windows Server 2008 R2"
	End If
End If

Private Sub FindIfSQLPresent
	Dim ObjService
	Set ObjWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & StrComputer & "\root\cimv2")
	Set ObjColServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name = 'MSSQLServer'")
	If ObjColServices.Count > 0 Then
		For Each ObjService in ObjColServices
			WScript.Echo "Microsoft SQL Server is present and " & ObjService.State & "."
			ChkSQLInstall = True
		Next
	Else
		WScript.Echo "Microsoft SQL Server is not installed on this Machine."
		ChkSQLInstall = False
	End If
	Set ObjColServices = Nothing:	Set ObjWMIService = Nothing
End Sub

Private Sub GetOSBit
	Dim ObjCompWin32
	Set ObjWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate,authenticationLevel=Pkt}!\\" & StrComputer & "\root\cimv2")
	Set ObjColServices = ObjWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each ObjCompWin32 in ObjColServices
		If ObjCompWin32.SystemType = "X86-based PC" Then
			OurOSBit = "X86"
		Else
			OurOSBit = "X64"
		End If
	Next
	Set ObjColServices = Nothing:	Set ObjWMIService = Nothing
End Sub

Private Sub CheckSQLVersion
	Dim ObjDrive
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
	Set WshShell = CreateObject("WScript.Shell")
		StrPath = Trim(WshShell.ExpandEnvironmentStrings("%temp%\CheckDirList.txt"))
	Set WshShell = Nothing
		WScript.Echo "Checking Microsoft SQL Server Version. Please wait ... "
		For Each ObjDrive In ObjFSO.Drives
			If ObjDrive.DriveType = 2 Then 
				DoSearch (ObjDrive.DriveLetter)
			End If
		Next:	Set ObjDrive = Nothing
		Set ObjFile = ObjFSO.OpenTextFile(StrPath, 1)
		ArrNames = Split(ObjFile.ReadAll, VbCrLf)
		ObjFile.Close:	Set ObjFile = Nothing
		For Each StrName In ArrNames
			If InStr(1, StrName, "sqlservr.exe", 1) > 0 Then
				If Right(Trim(StrName),12)="sqlservr.exe" And InStr(Trim(StrName),"Program Files") > 0 Then
					StrFileName = Trim(StrName):	StrVersion = ObjFSO.GetFileVersion(Trim(StrName))
					WScript.Echo "File Version -- " & StrVersion
					If Trim(StrVersion) >= "2000" And Trim(StrVersion) < "2001" Then 
						WScript.Echo "SQL 2000 Installed":	StrSQLName = "SQL 2000"
					End If
					If Trim(StrVersion) >= "2005" And Trim(StrVersion) < "2006" Then 
						WScript.Echo "SQL 2005 Installed":	StrSQLName = "SQL 2005"
					End If
					If Trim(StrVersion) >= "2007" And Trim(StrVersion) < "2008" Then 
						WScript.Echo "SQL 2008 Installed":	StrSQLName = "SQL 2008"
					End If
					If Trim(StrVersion) >= "2009" And Trim(StrVersion) < "2010" Then 
						WScript.Echo "SQL 2008 R2 Installed":	StrSQLName = "SQL 2008 R2"
					End If
					If Trim(StrVersion) >= "2010" And Trim(StrVersion) < "2012" Then 
						WScript.Echo "SQL 2012 Installed":	StrSQLName = "SQL 2012"
					End If
				End If	
			End If
		Next		
		If ObjFSO.FileExists(StrPath) Then 
			ObjFSO.DeleteFile(StrPath)
		End If
	Set ObjFSO = Nothing
End Sub

Private Sub DoSearch(StrDrive) 
	Set WshShell = CreateObject("WScript.Shell")
		WshShell.Run "cmd /c dir /s /b " & StrDrive & ":\" & StrName & " >>" & StrPath, 0, True
	Set WshShell = Nothing
End Sub

Private Sub CheckOSVersion
	Dim ObjOS
	On Error Resume Next
	' ---  Connect to WMI and obtain instances of Win32_OperatingSystem
	For Each ObjOS in GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem")
		StrOSName = Trim(ObjOS.Caption)
	Next
	If InStr(StrOSName, "2000") > 0 Then
		StrOSName = "Windows 2000"
	End If
	If InStr(StrOSName, "2003") > 0 Then
		If InStr(StrOSName, "R2") > 0 Then
			StrOSName = "Windows 2003 R2"
		Else
			StrOSName = "Windows 2003"
		End If
	End If
	If InStr(StrOSName, "2008") > 0 Then
		If Instr(StrOSName, "R2") > 0 Then
			StrOSName = "Windows 2008 R2"
		Else
			StrOSName = "Windows 2008"
		End If
	End If
	If InStr(StrOSName, "7") > 0 Then
		StrOSName = "Windows 7"
	End If
	If InStr(StrOSName, "XP") > 0 Then
		StrOSName = "Windows XP"
	End If
End Sub