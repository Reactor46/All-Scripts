'--------------------------------------------------------------------------------- 
' The sample scripts are not supported under any Microsoft standard support 
' program or service. The sample scripts are provided AS IS without warranty  
' of any kind. Microsoft further disclaims all implied warranties including,  
' without limitation, any implied warranties of merchantability or of fitness for 
' a particular purpose. The entire risk arising out of the use or performance of  
' the sample scripts and documentation remains with you. In no event shall 
' Microsoft, its authors, or anyone else involved in the creation, production, or 
' delivery of the scripts be liable for any damages whatsoever (including, 
' without limitation, damages for loss of business profits, business interruption, 
' loss of business information, or other pecuniary loss) arising out of the use 
' of or inability to use the sample scripts or documentation, even if Microsoft 
' has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 

Option Explicit 
On Error Resume Next 
Dim objshell, value, keypath, Choice
'Create wscript.shell object 
Set objshell = CreateObject("wscript.shell")
'The registry key path 
keypath = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Isolation"
'Read the value
value = objshell.RegRead(keypath)
'If error, the mode is setted by default.
If Err.Number = 0 Then 
	If value = "PMIL" Then 
		'enable
		Choice = MsgBox("Enhanced Protected Mode is disabled, do you want to enable it?(Need to restart IE to take effect)", 4 , "Comfirming")
		If Choice =  vbYes Then 
			objshell.RegWrite keypath, "PMEM"
		End If 
	End If 
	
	If value = "PMEM" Then 
		'Dsiable
		Choice = MsgBox("Enhanced Protected Mode is enabled, do you want to disable it?(Need to restart IE to take effect)", 4 , "Comfirming")
		If Choice =  vbYes Then 
			objshell.RegWrite keypath, "PMIL"
		End If 
	End If 
Else
	'Disable the Enhanced Protected Mode
	Const HKEY_CURRENT_USER = &H80000001
	Dim strComputer, strKeyPath, strValueName, strvalue, objRegistry
	strComputer = "."
	'Get wmi obejct 
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "Software\Microsoft\Internet Explorer\Main\"
	strValueName = "Isolation"
	strValue = "PMIL"
	'Get user choice 
	Choice = MsgBox("Enhanced Protected Mode is setted by defualt, do you want to disable it?(Need to restart IE to take effect)", 4 , "Comfirming")
	If Choice =  vbYes Then 
		'Set regstry value
		objRegistry.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strValue
	End If 

End If 