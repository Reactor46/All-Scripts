'--------------------------------------------------------------------------------- 
'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 
Option Explicit

Dim objShell,objWSHShell
Set objShell = CreateObject("Shell.Application")
Set objWSHShell = CreateObject("Wscript.Shell")

Dim SystemRoot,ProgramData,objDesktop
SystemRoot = objWSHShell.expandEnvironmentStrings("%SystemRoot%")
ProgramData = objWSHShell.expandEnvironmentStrings("%ProgramData%")
Set objDesktop = objShell.Namespace(ProgramData & "\Microsoft\Windows\Start Menu\Programs\")

' #############################
' Create Windows SwitchUserTileSwitchUser tile
' #############################
Dim SwitchUserShortcutPath
Dim objSwitchUserShortcut
'create a new shortcut of SwitchUser
SwitchUserShortcutPath = ProgramData & "\Microsoft\Windows\Start Menu\Programs\SwitchUser.lnk"
Set objSwitchUserShortcut = objWSHShell.CreateShortcut(SwitchUserShortcutPath)
objSwitchUserShortcut.TargetPath = SystemRoot & "\System32\tsdiscon.exe"
'change the default icon of SwitchUser shortcut
objSwitchUserShortcut.IconLocation = SystemRoot & "\System32\SHELL32.dll,264"
objSwitchUserShortcut.Save

Dim objSwitchUserLnk,SwitchUserVerbs

'pin application to windows Start menu
Set objSwitchUserLnk = objDesktop.ParseName("SwitchUser.lnk")
Set SwitchUserVerbs = objSwitchUserLnk.Verbs

'Verify the file can be pinned to taskbar
Dim userFlag,SwitchUserVerb
userFlag=0

For Each SwitchUserVerb in SwitchUserVerbs
    If Replace (SwitchUserVerb.Name,"&","") = "Pin to Start" Then
        SwitchUserVerb.DoIt
        userFlag = 1
	End If
Next
    
If userFlag = 1 Then
	WScript.Echo "Create Windows SwitchUser Tile successfully." 
Else
	WScript.Echo "Failed to create Windows SwitchUser Tile."
End If