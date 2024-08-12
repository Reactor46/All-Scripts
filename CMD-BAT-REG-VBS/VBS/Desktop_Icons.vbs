''Add Icons to the Desktop
option explicit
on error resume next
dim shell, desktopPath, link, sys32Path
Set shell = WScript.CreateObject("WScript.shell")
desktoplocation = shell.SpecialFolders("Desktop")
sys32Path = "%SystemRoot%\system32"

set link = shell.CreateShortcut(desktopPath & "\CAMS Local Logon.lnk")
link.Description = "insert name of icon here"
link.TargetPath = "path to program or file"
link.WindowStyle = 1
link.WorkingDirectory = desktoplocation
link.Save

set shell = nothing

'msgBox "Your desktop icon has been created." & vbCrLf & "Please check your Windows Desktop.", vbOKOnly-vbInformation, "Icon Added"

''End Icons on Desktop ----------------------------------------------------