'Script to write Logon Data Username, Computername, and IP address to eventlog.

Dim objShell, WshNetwork, PCName, UserName, strMessage
Dim strQuery, objWMIService, colItems, strIP
' Constants for type of event log entry
const EVENTLOG_INFORMATION = 4

set objShell = CreateObject("WScript.Shell")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objWMIService = GetObject( "winmgmts:" & "{impersonationLevel=impersonate}!\\" & ".\root\CIMV2")
PCName = WshNetwork.ComputerName 
UserName = WshNetwork.UserName

strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"
Set colItems = objWMIService.ExecQuery( strQuery, "WQL", 48 )

For Each objItem In colItems
    If IsArray( objItem.IPAddress ) Then
        If UBound( objItem.IPAddress ) = 0 Then
            strIP = objItem.IPAddress(0)
        Else
            strIP = Join( objItem.IPAddress, ", " )
        End If
    End If
Next

strMessage = "Logon Event Data" & VBCrLf & "PC Name: " & PCName & VBCrLf & "Username: " & UserName & VBCrLf & "IP Addresses: " & strIP
objShell.LogEvent EVENTLOG_INFORMATION, strMessage
WScript.Quit

' End Logon Script to add Login Info to Event Log