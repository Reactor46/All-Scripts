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

' #############################################
' Create Windows Hibernation tile
' #############################################
Dim HibernationShortcutPath
Dim objHibernationShortcut
'create a new shortcut of Hibernation
HibernationShortcutPath = ProgramData & "\Microsoft\Windows\Start Menu\Programs\Hibernation.lnk"
Set objHibernationShortcut = objWSHShell.CreateShortcut(HibernationShortcutPath)
objHibernationShortcut.TargetPath = SystemRoot & "\System32\rundll32.exe"
objHibernationShortcut.Arguments = "powrprof.dll,SetSuspendState"
'change the default icon of Hibernation shortcut
objHibernationShortcut.IconLocation = SystemRoot & "\System32\SHELL32.dll,297"
objHibernationShortcut.Save

Dim objHibernationLnk,HibernationVerbs

'pin application to windows Start menu
Set objHibernationLnk = objDesktop.ParseName("Hibernation.lnk")
Set HibernationVerbs = objHibernationLnk.Verbs

'Verify the file can be pinned to taskbar
Dim hFlag,HibernationVerb
hFlag=0

For Each HibernationVerb in HibernationVerbs
    If Replace (HibernationVerb.Name,"&","") = "Pin to Start" Then
        HibernationVerb.DoIt
        hFlag = 1
	End If
Next
    
If hFlag = 1 Then
	WScript.Echo "Create Windows Hibernation tile successfully." 
Else
	WScript.Echo "Failed to create Windows Hibernation tile."
End If