' This script will open the Outlook Secure Temp folder location for Outlook 2003, 2007, 2010 and now includes Office 2013

Dim WshShell, strValue, openlocation

Set WshShell = WScript.CreateObject("WScript.Shell")
const HKEY_CURRENT_USER = &H80000001
on error resume next
strComputer = "."

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Office\11.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

If IsNull(dwValue) Then

Else
strKeyPath = "Software\Microsoft\Office\11.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
openlocation = "EXPLORER.exe /e," & strValue
wShshell.Run openlocation
End If

strKeyPath = "Software\Microsoft\Office\12.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

If IsNull(dwValue) Then

Else
strKeyPath = "Software\Microsoft\Office\12.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
openlocation = "EXPLORER.exe /e," & strValue
wShshell.Run openlocation
End If

strKeyPath = "Software\Microsoft\Office\14.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

If IsNull(dwValue) Then

Else
strKeyPath = "Software\Microsoft\Office\14.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
openlocation = "EXPLORER.exe /e," & strValue
wShshell.Run openlocation
End If

strKeyPath = "Software\Microsoft\Office\15.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

If IsNull(dwValue) Then

Else
strKeyPath = "Software\Microsoft\Office\15.0\Outlook\Security\"
strValueName = "OutlookSecureTempFolder"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
openlocation = "EXPLORER.exe /e," & strValue
wShshell.Run openlocation
End If

WScript.Quit