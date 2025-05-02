''' Set Home Page in Browser
const HKEY_CURRENT_USER = &H80000001

strComputer ="."

set objReg = Getobject("winmgmts:\\" & strComputer & "\root\default:stdRegProv")

Set objWshShell = CreateObject("WScript.Shell") 

strHomePage = "address" 

objWshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\Start Page", strHomePage 

Value = "Secondary Start Pages"

value1 = "Address1"
value2 = "Address2"
arrValue = Array(value1, value2)

objReg.SetMultiStringValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main",value, arrValue

Set objWshShell = Nothing

''' End Browser Home Page Script