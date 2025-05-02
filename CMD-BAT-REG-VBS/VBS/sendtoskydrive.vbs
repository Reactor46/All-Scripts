'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 
Option Explicit 
Dim objshell, Target, UserFile, Appdata, FSO, intAnswer, ShortCut, path 
'Create wscript.shell object 
Set objshell = CreateObject("wscript.shell")
'Create FileSystemObject
Set FSO = CreateObject("Scripting.FileSystemObject")
'Get user profile path
UserFile = objshell.ExpandEnvironmentStrings("%UserProfile%")
'SkyDrive path
Target =  UserFile & "\SkyDrive"
If FSO.FolderExists(Target) Then 
	Appdata = objshell.ExpandEnvironmentStrings("%APPDATA%")
	'ShortCut path
	path = Appdata & "\Microsoft\Windows\SendTo\SkyDrive.lnk"
	'Check if shortcut exists
	If FSO.FileExists(path) Then 
		intAnswer = MsgBox ("'Send to SkyDrive' already exists, do you want remove it?",vbYesNo, "Delete File")
		If intAnswer = vbYes Then 
			FSO.DeleteFile path
			WScript.Echo "Remove 'Send to SkyDrive' successfully."
		End If 
	Else 
	'Create shortcut
		Set ShortCut = objshell.CreateShortcut(path)
		ShortCut.Targetpath = Target
		ShortCut.save
		WScript.Echo "Create 'Send to SkyDrive' successfully."
	End If 
Else 
	WScript.Echo "There is no SkyDrive or the SkyDrive is not in default path."
End If 