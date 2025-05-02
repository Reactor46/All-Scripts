'Script to write Logoff Data Username, Computername to eventlog.
Dim objShell, WshNetwork, PCName, UserName, strMessage
Dim strQuery, objWMIService, colItems, strIP
' Constants for type of event log entry
const EVENTLOG_AUDIT_SUCCESS = 8

set objShell = CreateObject("WScript.Shell")
Set WshNetwork = WScript.CreateObject("WScript.Network")

PCName = WshNetwork.ComputerName 
UserName = WshNetwork.UserName

strMessage = "Logoff Event Data" & VBCrLf & "PC Name: " & PCName & VBCrLf & "Username: " & UserName
objShell.LogEvent EVENTLOG_AUDIT_SUCCESS, strMessage
WScript.Quit

'End Script