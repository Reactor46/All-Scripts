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
' Create Windows shutdown tile
' #############################
Dim ShutdownShortcutPath
Dim objShutdownShortcut
'create a new shortcut of shutdown
ShutdownShortcutPath = ProgramData & "\Microsoft\Windows\Start Menu\Programs\Shutdown.lnk"
Set objShutdownShortcut = objWSHShell.CreateShortcut(ShutdownShortcutPath)
objShutdownShortcut.TargetPath = SystemRoot & "\System32\shutdown.exe"
objShutdownShortcut.Arguments = "-s -t 0"
'change the default icon of shutdown shortcut
objShutdownShortcut.IconLocation = SystemRoot & "\System32\SHELL32.dll,27"
objShutdownShortcut.Save

Dim objShutdownLnk,ShutdownVerbs

'pin application to windows Start menu
Set objShutdownLnk = objDesktop.ParseName("Shutdown.lnk")
Set ShutdownVerbs = objShutdownLnk.Verbs

'Verify the file can be pinned to taskbar
Dim shtdFlag,ShutdownVerb
shtdFlag=0

For Each ShutdownVerb in ShutdownVerbs
    If Replace (ShutdownVerb.Name,"&","") = "Pin to Start" Then
        ShutdownVerb.DoIt
        shtdFlag = 1
	End If
Next
    
If shtdFlag = 1 Then
	WScript.Echo "Create Windows shutdown tile successfully."
Else
	WScript.Echo "Failed to create Windows shutdown tile."
End If

' #############################
' Create Windows restart tile
' #############################
Dim RestartShortcutPath
Dim objRestartShortcut
'create a new shortcut of restart
RestartShortcutPath = ProgramData & "\Microsoft\Windows\Start Menu\Programs\Restart.lnk"
Set objRestartShortcut = objWSHShell.CreateShortcut(RestartShortcutPath)
objRestartShortcut.TargetPath = SystemRoot & "\System32\shutdown.exe"
objRestartShortcut.Arguments = "-r -t 0"
'change the default icon of restart shortcut
objRestartShortcut.IconLocation = SystemRoot & "\System32\SHELL32.dll,238"
objRestartShortcut.Save

Dim objRestartLnk,RestartVerbs

'pin application to windows Start menu
Set objRestartLnk = objDesktop.ParseName("Restart.lnk")
Set RestartVerbs = objRestartLnk.Verbs

'Verify the file can be pinned to taskbar
Dim rstFlag,RestartVerb
rstFlag=0

For Each RestartVerb in RestartVerbs
    If Replace (RestartVerb.Name,"&","") = "Pin to Start" Then
        RestartVerb.DoIt
        rstFlag = 1
	End If
Next
    
If rstFlag = 1 Then
	WScript.Echo "Create Windows restart tile successfully."
Else
	WScript.Echo "Failed to create Windows restart tile."
End If


' #############################
' Create Windows logoff tile
' #############################
Dim LogoffShortcutPath
Dim objLogoffShortcut
'create a new shortcut of logoff
LogoffShortcutPath = ProgramData & "\Microsoft\Windows\Start Menu\Programs\Logoff.lnk"
Set objLogoffShortcut = objWSHShell.CreateShortcut(LogoffShortcutPath)
objLogoffShortcut.TargetPath = SystemRoot & "\System32\shutdown.exe"
objLogoffShortcut.Arguments = "-s -t 0"
'change the default icon of logoff shortcut
objLogoffShortcut.IconLocation = SystemRoot & "\System32\SHELL32.dll,44"
objLogoffShortcut.Save

Dim objLogoffLnk,LogoffVerbs

'pin application to windows Start menu
Set objLogoffLnk = objDesktop.ParseName("Logoff.lnk")
Set LogoffVerbs = objLogoffLnk.Verbs

'Verify the file can be pinned to taskbar
Dim lgFlag,LogoffVerb
lgFlag=0

For Each LogoffVerb in LogoffVerbs
    If Replace (LogoffVerb.Name,"&","") = "Pin to Start" Then
        LogoffVerb.DoIt
        lgFlag = 1
	End If
Next
    
If lgFlag = 1 Then
	WScript.Echo "Create Windows logoff tile successfully."
Else
	WScript.Echo "Failed to create Windows logoff tile."
End If