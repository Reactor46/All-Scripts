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
Option Explicit 
On Error Resume Next 
If WScript.Arguments.Count = 0 Then
	Dim objShell
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "open", 1
Else
	Dim WshShell, USERPROFILE, receiver, ShortName 
'	Create wscript.shell object 
	Set WshShell = CreateObject("WScript.Shell")
'	get the user profile path
	USERPROFILE =  WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
	WScript.StdOut.Write("Please input receivers' address(usng semicolon between the receivers):")
	receiver = WScript.StdIn.ReadLine()
	WScript.StdOut.Write("Give a name to the shortcut(such as ToTom&john):")
	ShortName = WScript.StdIn.ReadLine()
	
	Dim FSO, ShortPath, Shortcut
'	create scripting.filesystemobject object 
	If  Len(ShortName) = 0  Then 
		ShortPath = USERPROFILE & "\Desktop\MailTo" & receiver & ".lnk"
	Else 
		ShortPath = USERPROFILE & "\Desktop\" & ShortName & ".lnk"	
	End If 
	Set Shortcut = WshShell.CreateShortcut(ShortPath)
	'	set the target path
	Shortcut.TargetPath = "MailTo:" & receiver 
	'	set the arguments
	Shortcut.Arguments = Filepath
	'	set the iconlocation
	Shortcut.IconLocation = "%Systemdrive%\Program Files\Windows Mail\wabmig.exe"
	Shortcut.Description = "To " & receiver
	Shortcut.Save()
	'	err checking
	Dim FSObject 
	Set FSObject = CreateObject("scripting.filesystemobject")
	
	If FSObject.FileExists(ShortPath) Then 
		WScript.StdOut.WriteLine "Create shortcut successfully."
		
	Else 
		WScript.StdOut.WriteLine "Failed to create shortcut"
	End If
WScript.Sleep(3000)

End If 