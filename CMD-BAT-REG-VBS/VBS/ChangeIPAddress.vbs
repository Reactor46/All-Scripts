' ChangeIPAddress.vbs 172.16.11.54

Option Explicit
Dim objWMIService
Dim objNetAdapter
Dim strComputer     
Dim strAddress     
Dim arrIPAddress
Dim arrSubnetMask
Dim colNetAdapters
Dim errEnableStatic

If WScript.Arguments.Count = 0 Then
     Wscript.Echo "Usage: ChangeIPAddress.vbs new_IP_address"
     WScript.Quit
End If

strComputer = "."
strAddress = Wscript.Arguments.Item(0)
arrIPAddress = Array(strAddress)
arrSubnetMask = Array("255.255.255.0")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each objNetAdapter in colNetAdapters
     errEnableStatic = objNetAdapter.EnableStatic(arrIPAddress, arrSubnetMask)
Next
