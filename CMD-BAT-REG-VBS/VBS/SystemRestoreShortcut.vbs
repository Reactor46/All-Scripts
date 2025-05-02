
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
Dim WshShell, USERPROFILE, Filepath 
'Create wscript.shell object 
Set WshShell = CreateObject("WScript.Shell")
'get the user profile path
USERPROFILE =  WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
'set the vbscript file path
Filepath = USERPROFILE & "\SystemRestore.vbs"
Dim FSO,VBSFile, ShortPath, Shortcut
'create scripting.filesystemobject object 
Set FSO = CreateObject("Scripting.filesystemobject")
'create the vbscript file 
Set VBSFile = FSO.CreateTextFile(Filepath)
'write some content into vbscript
VBSFile.WriteLine "rp = ""Restore point created in "" & Time & Date "
VBSFile.WriteLine "GetObject(""winmgmts:\\.\root\default:Systemrestore"").CreateRestorePoint rp, 0, 100"
VBSFile.WriteLine "Msgbox ""Create restore point successfully."""
'close the file 
VBSFile.Close
'create the shortcut on desktop
ShortPath = USERPROFILE & "\Desktop\Create System Restore Point.lnk"
Set Shortcut = WshShell.CreateShortcut(ShortPath)
'set the target path
Shortcut.TargetPath = "C:\windows\System32\wscript.exe" 
'set the arguments
Shortcut.Arguments = Filepath
'set the iconlocation
Shortcut.IconLocation = "%SystemRoot%\system32\rstrui.exe"
Shortcut.Description ="System Restore Point"
Shortcut.Save()
'err checking
If Err.Number = 0 Then 
	MsgBox "Create 'Create System Restore Point' shortcut successfully."
Else 
	MsgBox "Failed to create 'Create System Restore Point' shortcut"
End If 