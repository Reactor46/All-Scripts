Set wshShell = WScript.CreateObject( "WScript.Shell" )

strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

strComputer = "."
strIP ="."
strGateway="."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
 
For Each objAdapter in colAdapters
      strIP = objAdapter.IPAddress(0)
      strGateway = objAdapter.DefaultIPGateway(0)
Next

WScript.Echo "Computer Name: " & strComputerName & vbCr & "IP Address: " & strIP & vbCr & "Gateway: " & strGateway
