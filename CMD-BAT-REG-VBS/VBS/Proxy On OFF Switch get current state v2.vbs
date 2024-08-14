
Set WshShell = WScript.CreateObject("WScript.Shell")

DimSetting = WshShell.RegRead ("HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")

If DimSetting = 1 Then 
'Wscript.Echo "Proxy is Currently Enabled"

returnvalue=MsgBox ("Do you want to DISABLE the Proxy?", 36, "Proxy state is Currently ENABLED")


if returnvalue = 6 then 
' BEGIN CALLOUT A
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\" & _
	"CurrentVersion\Internet Settings\ProxyEnable", _
	0, "REG_DWORD"
' END CALLOUT A

wShshell.Run "taskkill /im iexplore.exe",0,True
WScript.Sleep 500 'give kill time to take effect

'Inform User that the task is done.
Mybox = MsgBox("Proxy is now disabled" ,vbOkOnly,"Proxy is Disabled")
're-launch internet explorer
wShshell.Run "iexplore.exe"
End If
WScript.Quit


Else 
'Wscript.Echo "Proxy is Currently Disabled" 
End If


'On Error Resume Next

returnvalue=MsgBox ("Do you want to ENABLE the Proxy?", 36, "Proxy state is Currently DISABLED")

If returnvalue = "" Then
  WScript.Quit
End If

if returnvalue = 6 then 
' BEGIN CALLOUT A
'Replace <Domain Proxy Server> in line 50  with IP or your proxy server name without the <> eg. ContosoProxy:80
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\" & _
	"CurrentVersion\Internet Settings\ProxyEnable", _
	1, "REG_DWORD"
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\" & _
	"CurrentVersion\Internet Settings\ProxyServer","<Domain Proxy Server>:80", "REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Windows\" & _
	"CurrentVersion\Internet Settings\ProxyOverride","<local>", "REG_SZ"
' END CALLOUT A


wShshell.Run "taskkill /im iexplore.exe",0,True

WScript.Sleep 500 'give kill time to take effect

'Inform User that the task is done.
Mybox = MsgBox("Proxy is enabled and set to enter your companies proxy server",vbOkOnly,"Proxy Enabled")
're-launch internet explorer
wShshell.Run "iexplore.exe"
End If

WScript.Quit





