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

'Run script with adminsitrator
If WScript.Arguments.Count = 0 Then
	Dim objShell
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else
	'Create wscript.shell object 
	Set objshell = CreateObject("wscript.shell")
	'get user password 
	WScript.StdOut.Write "Please enter current user password:"
	Password = WScript.StdIn.ReadLine()
	'Check if password is null
	If password <> "" Then 
		Set result = objshell.Exec("cmd /c  whoami")
		Do While Not  result.StdOut.AtEndOfStream 
			username =  result.StdOut.ReadLine()
		Loop 
		
		'Change the registry key value
		DefaultDomain = left( username, InStr(username,"\")-1)
		UserProfile =  objshell.ExpandEnvironmentStrings("%UserProfile%")
		DefaultUser =  Right(UserProfile, InStrRev(UserProfile,"\")-2)
		path = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
		objshell.RegWrite path & "\AutoAdminLogon", 1 , "REG_SZ"
		objshell.RegWrite path & "\DefaultDomainName", DefaultDomain, "REG_SZ"
		objshell.RegWrite path & "\DefaultuserName", defaultuser , "REG_SZ"
		objshell.RegWrite path & "\DefaultPassword", Password , "REG_SZ"
		WScript.StdOut.Write "Operation done.Press enter to exit."
	Else 
		WScript.StdOut.Write "Operation cancelled by user.Press enter to exit."
		
	End If 
	
	WScript.StdIn.ReadLine()
End If 