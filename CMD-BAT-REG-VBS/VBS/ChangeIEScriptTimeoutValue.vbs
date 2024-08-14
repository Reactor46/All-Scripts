Option Explicit

Dim Registry 
Dim sKeyPath
Dim sValueName
Dim dwValue
Dim sComputer

CONST HKEY_CURRENT_USER = &H80000001
 
sKeyPath = "Software\Microsoft\Internet Explorer\Styles"
sValueName = "MaxScriptStatements"
dwValue = -1
sComputer = "."
 
On Error Resume Next
Set Registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
				sComputer & "\root\default:StdRegProv")
 
Registry.CreateKey HKEY_CURRENT_USER, sKeyPath
Registry.SetDWORDValue HKEY_CURRENT_USER, sKeyPath, sValueName, CLng(dwValue)