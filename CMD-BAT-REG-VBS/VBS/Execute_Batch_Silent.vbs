Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.run "%comspec% /c C:\test.bat",0
Set WshShell = Nothing