Set objFirewall = CreateObject("HNetCfg.FwMgr") 
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile 
 
Set colPorts = objPolicy.GloballyOpenPorts 
 
For Each objPort in colPorts 
    Wscript.Echo "Port name: " & objPort.Name 
    Wscript.Echo "Port number: " & objPort.Port 
    Wscript.Echo "Port IP version: " & objPort.IPVersion 
    Wscript.Echo "Port protocol: " & objPort.Protocol 
    Wscript.Echo "Port scope: " & objPort.Scope 
    Wscript.Echo "Port remote addresses: " & objPort.RemoteAddresses 
    Wscript.Echo "Port enabled: " & objPort.Enabled 
    Wscript.Echo "Port built-in: " & objPort.Builtin 
Next 