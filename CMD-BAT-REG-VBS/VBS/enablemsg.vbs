' enablemsg.vbs
' script enable the use of the MSG utility from other computers
' Version 1.0
' By Pedro Lima (pedrofln.blogspot.com)
' Information Technology, MBA
' MCT, MCSE, MCSA, MCP+I, Network+ Certified Professional
' ------------------------------------------------------------

Option Explicit
Dim objShell
Dim strKey, strValue, strType

' Configure MSG to receive messages from other computers
strKey = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\AllowRemoteRPC"
strValue = "1"
strType = "REG_DWORD"
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite strKey, strValue, strType
