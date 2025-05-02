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
	Dim objSh
	Set objSh = CreateObject("Shell.Application")
	objSh.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "runas", 1
Else

'Declare two variables
Dim objshell, Path
'Create wscript.shell object 
Set objshell = CreateObject("Wscript.shell")
'Run history registry path
Path = objshell.ExpandEnvironmentStrings("%HOMEDRIVE%") & "\Windows.old"

Set Fsobject = CreateObject("Scripting.filesystemobject")
If Fsobject.FolderExists(Path) Then 
	 Regpath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations\StateFlags1221"
	 objshell.RegWrite Regpath, 2, "REG_DWORD"
	 objshell.Run "cleanmgr /SAGERUN:1221", 1, True 
	 WScript.Echo "Delete Windows.old folder successfully."
Else 
	
	WScript.Echo "Not find Windows.old folder"
End If 

End If 