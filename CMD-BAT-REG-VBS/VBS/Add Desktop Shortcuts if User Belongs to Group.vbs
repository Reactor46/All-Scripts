' Script installs desktop icons based on group membership
' Bill Reid, 2006
' -----------------------------------------------------------------------------------

Option Explicit
Dim objNetwork, objUser, CurrentUser, strDesktop
Dim strGroup, WshShell, oShellLink
'
' ================================================
'  Configure
'
Const userGroup1 = "cn=AD Group 1 users"
' ================================================

' Create objects and extract strGroup values
Set objNetwork = CreateObject("WScript.Network")
Set objUser = CreateObject("ADSystemInfo")
Set CurrentUser = GetObject("LDAP://" & objUser.UserName)
strGroup = LCase(Join(CurrentUser.MemberOf))

set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

' Assign shortcuts

' == userGroup1 ====================================
If InStr(strGroup, lcase(userGroup1)) Then

	' Example of an Internet Shortcut
	set oShellLink = WshShell.CreateShortcut(strDesktop & "\ShortcutName.lnk")
	oShellLink.TargetPath = "http://url.com"
	oShellLink.IconLocation = "C:\pathtoicon\icon.ico" 
	oShellLink.Description = "ShortcutName"
	oShellLink.Save
	
End If
' ================================================

WScript.Quit