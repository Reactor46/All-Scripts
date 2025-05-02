const KEY_QUERY_VALUE = &H0001
const KEY_SET_VALUE = &H0002
const KEY_CREATE_SUB_KEY = &H0004
const DELETE = &H00010000
const HKEY_CURRENT_USER = &H80000001
const HKEY_LOCAL_MACHINE = &H80000002
Set objShell = CreateObject("Wscript.Shell")
sUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
strComputer = "."
Set StdOut = WScript.StdOut

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")

strKeyPath = "Software\Shoreline Teleworks\ShoreWare Client"
oReg.CreateKey HKEY_CURRENT_USER,strKeyPath 

strKeyPath = "Software\Shoreline Teleworks\ShoreWare Client"
strValueName = "UserName"
strValue = sUserName
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue

strValueName = "Server"
strValue = "MSOVOIP01"
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue

strValueName = "Password"
strValue = "a26f5087aeb1e80f8916e8644b4e4250"
oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
 
