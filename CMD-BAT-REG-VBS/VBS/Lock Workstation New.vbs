Option Explicit 

Dim Text, Title, Lnk_Title, Shell
Dim WshShell    ' Object variable
Dim Shortcut, fso, Windows, AllUsersDesktop, Link

Set fso = CreateObject("Scripting.FileSystemObject")
Set Windows = fso.GetSpecialFolder(0)

Set WshShell = WScript.CreateObject("WScript.Shell")
Set Shell = CreateObject("WScript.Shell")
AllUsersDesktop = Shell.SpecialFolders("AllUsersDesktop")


	Set Link = WshShell.CreateShortcut(AllUsersDesktop & "\Lock Workstation.lnk")
        Link.Description = "Locks the computer"
	Link.TargetPath = "%windir%\system32\rundll32.exe" 
        Link.Arguments = "user32.dll,LockWorkStation"
	Link.IconLocation = "%SystemRoot%\system32\SHELL32.dll,47"
	Link.Save

'End
