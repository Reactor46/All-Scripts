' ===========================================================
' VB Script: Obtain the Uninstall String and MSI GUID
' ===========================================================

Option Explicit

Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1

Dim ObjWshShell, ObjWMI, WinDir, RegExp, CurPos
Dim ObjColWin32, ObjCompWin32, ObjFSO
Dim StrSoftwareName, StrApplicationType, StrDisplayName
Dim StrSystemType, StrProgramFiles, StrSoftwareRegistryKey, StrValue
Dim ObjReadFSO, ReadThisFile, OneLine, StrReadPath

Set ObjWshShell = WScript.CreateObject("WScript.Shell")
Set ObjWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
WinDir = ObjWshShell.ExpandEnvironmentStrings("%windir%")
Set RegExp = New RegExp
RegExp.IgnoreCase = True

Set ObjColWin32 = ObjWMI.ExecQuery ("Select * from Win32_ComputerSystem")
For Each ObjCompWin32 in ObjColWin32
	If ObjCompWin32.SystemType = "X86-based PC" Then
		StrSystemType = "X86"
		WScript.Echo "This is a 32 Bit System"
	Else
		StrSystemType = "X64"
		WScript.Echo "This is a 64 Bit System"
	End If
Next

' --- Now Delete the CMD File, if it exists
DoThisDeleteJob

Set ObjReadFSO = CreateObject("Scripting.FileSystemObject")
If ObjReadFSO.FileExists(ObjReadFSO.GetFile(WScript.ScriptFullName).ParentFolder & "\UnInstallList.txt") = True Then
	StrReadPath = Trim(ObjReadFSO.GetFile(WScript.ScriptFullName).ParentFolder)
	WScript.Echo "<<**>> Checking Software Uninstallation Status. Please wait ..." & VbCrLf
	Set ObjReadFSO = Nothing
Else
	StrReadPath = vbNullString
	WScript.Echo VbCrLf & "Cannot continue." & VbCrLf & "The required File -- UnInstallList.txt -- is missing."
	Set ObjReadFSO = Nothing:	WScript.Quit
End If

CurPos = 0
Select Case StrSystemType
	Case "X86"
		Set ObjReadFSO = CreateObject("Scripting.FileSystemObject")
		StrReadPath = Trim(ObjReadFSO.GetFile(WScript.ScriptFullName).ParentFolder)
		Set ReadThisFile = ObjReadFSO.OpenTextFile(StrReadPath & "\UnInstallList.txt", ForReading)
		Do Until ReadThisFile.AtEndOfStream
			OneLine = ReadThisFile.ReadLine
			CurPos = CurPos + 1:	WScript.Echo " ===> Checking Status: " & CurPos
			GetTheAppItem "32-bit"
			NowDoUnInstall OneLine				
		Loop			
		ReadThisFile.Close:	Set ReadThisFile = Nothing:	Set ObjReadFSO = Nothing		
	Case "X64"
		Set ObjReadFSO = CreateObject("Scripting.FileSystemObject")
		StrReadPath = Trim(ObjReadFSO.GetFile(WScript.ScriptFullName).ParentFolder)
		Set ReadThisFile = ObjReadFSO.OpenTextFile(StrReadPath & "\UnInstallList.txt", ForReading)
		Do Until ReadThisFile.AtEndOfStream
			OneLine = ReadThisFile.ReadLine
			CurPos = CurPos + 1:	WScript.Echo " ===> Checking Status: " & CurPos
			GetTheAppItem "64-bit"
			NowDoUnInstall OneLine
			GetTheAppItem "32-bit"
			NowDoUnInstall OneLine				
		Loop			
		ReadThisFile.Close:	Set ReadThisFile = Nothing:	Set ObjReadFSO = Nothing		
End Select

WScript.Echo vbNullString
WScript.Echo "<<**>> Software Uninstallation Check Completed."

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
If ObjFSO.FileExists(StrReadPath & "\SoftwareList.cmd") = True Then
	WScript.Echo "<<**>> Please check and execute the Batch File - SoftwareList.cmd." & VbCrlf & "<<**>> It has been created."
End If
Set ObjFSO = Nothing:	WScript.Quit

Private Sub GetTheAppItem (StrAppType)
	' -- Set the path to the correct "Program Files" folder and "SOFTWARE" registry key based on the application and OS type
	StrApplicationType = StrAppType
	If StrSystemType = "X64" Then
		If StrApplicationType = "64-bit" Then
			StrSoftwareRegistryKey = "SOFTWARE"
			If UCase(WScript.FullName) = UCase(WinDir & "\SysWOW64\WScript.exe") Then
				' -- 64-bit Application on a 64-bit OS from a 32-bit WScript.exe Process
				StrProgramFiles = ObjWshShell.ExpandEnvironmentStrings("%ProgramW6432%")
			Else
				' -- 64-bit Application on a 64-bit OS from a 64-bit WScript.exe Process
				StrProgramFiles = ObjWshShell.ExpandEnvironmentStrings("%ProgramFiles%")
			End If
		Else
			StrSoftwareRegistryKey = "SOFTWARE\Wow6432Node"
			' -- 32-bit Application on a 64-bit OS (the WScript.exe process type does not matter)
			StrProgramFiles = ObjWshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
		End If
	Else
		StrSoftwareRegistryKey = "SOFTWARE"
		' -- 32-bit Application on a 32-bit OS from a 32-bit WScript.exe Process
		StrProgramFiles = ObjWshShell.ExpandEnvironmentStrings("%ProgramFiles%")
	End If
End Sub

Private Sub NowDoUnInstall (StrValueCheck)
	Dim ObjReg, StrKeyPath, Subkey, ArrSubKeys, StrCheckKey
	StrKeyPath = StrSoftwareRegistryKey & "\Microsoft\Windows\CurrentVersion\Uninstall"
	' --- WScript.Echo "Software GUID Search In Location: " & StrKeyPath
	Set ObjReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	ObjReg.EnumKey HKEY_LOCAL_MACHINE, StrKeyPath, ArrSubKeys
	For Each Subkey In ArrSubKeys
		StrValue = ""
		StrCheckKey = StrKeyPath & "\" & Subkey
		ObjReg.GetStringValue HKEY_LOCAL_MACHINE, StrCheckKey, "DisplayName", StrValue
		If Not IsNull(StrValue) Then
			RegExp.Pattern = StrValueCheck
			If RegExp.Test (StrValue) = TRUE Then
				StrDisplayName = StrValue
				' -- WScript.Echo "Display Name --- " & StrValue
				' -- Attempts to Obtain the UninstallString for the Matching String.
				ObjReg.GetStringValue HKEY_LOCAL_MACHINE, StrCheckKey, "UninstallString", StrValue
				WriteToUninstallFile
			End If
		End If
	Next
	Set ObjReg = Nothing
End Sub

Private Sub WriteToUninstallFile
	Dim ThisObjFSO, NewFileToCreate
	Set ThisObjFSO = CreateObject("Scripting.FileSystemObject")
	If ThisObjFSO.FileExists(StrReadPath & "\SoftwareList.cmd") = False Then
		Set NewFileToCreate = ThisObjFSO.OpenTextFile(StrReadPath & "\SoftwareList.cmd", 8, True, 0)
		NewFileToCreate.WriteLine "@ECHO OFF"
		NewFileToCreate.WriteLine vbNullString	
	Else
		Set NewFileToCreate = ThisObjFSO.OpenTextFile(StrReadPath & "\SoftwareList.cmd", 8, True, 0)
	End If
	NewFileToCreate.WriteLine "REM --- Uninstall " & StrDisplayName 
	NewFileToCreate.WriteLine "start /wait psexec cmd /c " & StrValue
	NewFileToCreate.WriteLine vbNullString:	NewFileToCreate.Close
	Set NewFileToCreate = Nothing:	Set ThisObjFSO = Nothing
End Sub

Private Sub DoThisDeleteJob
	Dim NewObjFSO, StrFilePath
	Set NewObjFSO = CreateObject("Scripting.FileSystemObject")
	If NewObjFSO.FileExists(NewObjFSO.GetFile(WScript.ScriptFullName).ParentFolder & "\SoftwareList.cmd") = True Then
		StrFilePath = Trim(NewObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
		NewObjFSO.DeleteFile StrFilePath & "\SoftwareList.cmd", True
	End If
	Set NewObjFSO = Nothing
End Sub