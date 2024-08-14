'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------


' UAC prompt for elevation

If WScript.Arguments.Count = 0 Then
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else

Message  = "Yes      :    Set always showing all Users at Sign-in"_
 & vbNewLine & "No       :    Set always showing last user at Sign-in"_
 & vbNewLine & "Cancel:   Cancel operation."
Choice = MsgBox(Message,3,"System Message")

If Choice = VbYes Then 
	'All users
	'take ownership of the registry key
	CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
	PsPath = CurrentDirectory & "TakeOwnerShip.ps1"
	set FSO = CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(PsPath) Then 
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "PowerShell -nologo " & PsPath , 0, True
		' Create a temp file with the script that regini.exe will use
		strFileName = FSO.GetTempName
		set oFile = FSO.CreateTextFile(strFileName)
		oFile.WriteLine "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\UserSwitch [2 8 19]"
		oFile.WriteLine "Enabled = REG_DWORD 1"
		oFile.WriteLine "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
		oFile.WriteLine "AutoAdminLogon = REG_SZ 0"
		oFile.Close
		
	'	Change registry permissions with regini.exe
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "regini " & strFileName, 0, true
		
	'	 Delete temp file
		
		FSO.DeleteFile strFileName
		
		WScript.Echo "Windows 8 now set to show 'all users' at sign in"
	Else 
		MsgBox "Do not find TakeOwnerShip.ps1, failed to execute script."
	End If 
ElseIf Choice = vbNo Then 
	'Last logon user

	' Create a temp file with the script that regini.exe will use
	CurrentDirectory = left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(len(WScript.ScriptName)))
	PsPath = CurrentDirectory & "TakeOwnerShip.ps1"
	Set FSO = CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(PsPath) Then 
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "PowerShell -nologo " & PsPath , 0, True
	
		strFileName = FSO.GetTempName
		set oFile = FSO.CreateTextFile(strFileName)
		oFile.WriteLine "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\UserSwitch [1 8 17]"
		oFile.WriteLine "Enabled = REG_DWORD 0"
		oFile.WriteLine "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
		oFile.WriteLine "AutoAdminLogon = REG_SZ 0"
		oFile.Close
		
		
		' Change registry permissions with regini.exe
		
		R=0
		Do until R=2
		R=R+1
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "regini " & strFileName, 0, true
		Loop
		
		' Delete temp file
		
		FSO.DeleteFile strFileName
		
		WScript.Echo "Windows 8 now set to show only 'last user that signed out' at sign in"
	Else 
		MsgBox "Do not find TakeOwnerShip.ps1, failed to execute script."
	End If 
End If 



End If