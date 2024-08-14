If WScript.Arguments.Named.Exists("elevated") = False Then
  CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
  WScript.Quit
End If

Set objArgs = WScript.Arguments 
Set WshShell = Wscript.CreateObject("Wscript.Shell") 
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

'Message Box with just prompt message
  'Message Box with title, yes no and cancel Butttons 
  MsgBox("In order to reactivate MS Licensing you must login to Terminal server. By clicking OK the fix will be applied and remote desktop will be launched, please log-in and log-off using your credentials. Be noted that the registry will be exported for backup purposes to your desktop 'mslicense_backup.reg'.")


Dim wsh
set wsh = createobject("WScript.Shell")
wsh.run ("regedit /e C:\users\" & strUserName & "\Desktop\mslicense_backup.reg ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSLicensing""")
WScript.Sleep 2000
wsh.run("REG.EXE DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSLicensing" & " /f")
set wsh = nothing 

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "C:\Windows\system32\mstsc.exe", "/v:YOURSERVER-FQDN", "", "runas", 1