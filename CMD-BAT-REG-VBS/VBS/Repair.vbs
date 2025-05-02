
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


Sub main()
	Dim objshell, AppPath,AppData,Path,FSO,Destination
	'Create wscript.shell object 
	Set objshell = CreateObject("wscript.shell")
	'Create scripting.filesystemobject 
	Set FSO = CreateObject("scripting.FileSystemObject")
	'Get AppData foler path 
	AppData = objshell.ExpandEnvironmentStrings("%APPDATA%")
	'Get shortcut path 
	Path = AppData + "\Microsoft\Windows\SendTo\Desktop (create shortcut).DeskLink"
	'Get shortcut parent folder path 
	Destination = AppData & "\Microsoft\Windows\SendTo\"
	'Check if the file exists
	If FSO.FileExists(path) Then
		WScript.Echo "Desktop (create shortcut).DeskLink file exsits.Your problem may not be caused by file missing."
	Else 
		Dim DefaultPath
		DefaultPath = "C:\Users\Default\AppData\Roaming\Microsoft\Windows\SendTo\Desktop (create shortcut).DeskLink"
		If FSO.FileExists(DefaultPath) Then 
			'Copy file 
			FSO.CopyFile DefaultPath,Destination 
		Else 
			'Create file
			FSO.CreateTextFile(Path)
		End If 
		If Err.Number <> 0 Then 
			WScript.Echo Err.Description
		Else
			WScript.Echo "Missing file has been fixed.Please check your context menu."
		End If 
	End If  
End Sub 

Call main 