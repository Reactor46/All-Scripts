On Error Resume Next
const KEY_QUERY_VALUE = &H0001
const KEY_SET_VALUE = &H0002
const KEY_CREATE_SUB_KEY = &H0004
const DELETE = &H00010000
const HKEY_CURRENT_USER = &H80000001
const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set objADSystemInfo = CreateObject("ADSystemInfo")
Set objUser = GetObject("LDAP://" & objADSystemInfo.UserName)
varProxy = objUser.GetEx("proxyAddresses")
sEmail = "NeEmail@usonv.com"
For i = 0 To UBound(varProxy)
	If left(varProxy(i),5) = "SMTP:" Then
		sEmail = Mid(varProxy(i),6)
	End if
Next

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\GFI FAX & VOICE\Intranet"
oReg.CreateKey HKEY_CURRENT_USER,strKeyPath 
strValueName = "smtp"
strValue = "usonvsvrfax01.uson.local"
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
strValueName = "sender"
strValue = sEmail
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue