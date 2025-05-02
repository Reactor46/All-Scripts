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
Dim objShell,WshShell
Dim WindowsAppsLnkPath,WindowsAppsShortcut
Dim envUSER,Flag

Set objShell = CreateObject("Shell.Application")  
Set WshShell = CreateObject("WScript.Shell")

envUSER = WshShell.expandEnvironmentStrings("%username%")

'Create a shortcut link
WindowsAppsLnkPath = "C:\Users\" & envUSER & "\Desktop\WindowsApps.lnk"
Set WindowsAppsShortcut = WshShell.CreateShortcut(WindowsAppsLnkPath)
WindowsAppsShortcut.TargetPath = "C:\Windows\explorer.exe"
WindowsAppsShortcut.Arguments = "shell:::{4234d49b-0245-4df3-b780-3893943456e1}"
'Change the default icon of shortcut
WindowsAppsShortcut.IconLocation = "C:\Windows\System32\twinui.dll, 91"
WindowsAppsShortcut.Save


'Pin shortcut to taskbar
Dim objFolder,objFolderItem,colVerbs,objVerb

Set objFolder = objShell.Namespace("C:\Users\" & envUSER & "\Desktop\")
Set objFolderItem = objFolder.ParseName("WindowsApps.lnk") 
Set colVerbs = objFolderItem.Verbs 

'Verify the file can be pinned to taskbar
Flag=0

For Each objVerb in colVerbs	
	If Replace(objVerb.name, "&", "") = "Pin to Taskbar" Then 
		objVerb.DoIt
		Flag = 1
	End If
Next

If Flag = 1 Then 
	msgbox "Create Windows apps shortcut file successfully."
Else 
	msgbox "Failed to create Windows apps shortcut file."
End If 