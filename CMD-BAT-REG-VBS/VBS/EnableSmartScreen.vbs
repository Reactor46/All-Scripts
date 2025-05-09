'#--------------------------------------------------------------------------------- 
'#The sample scripts are not supported under any Microsoft standard support 
'#program or service. The sample scripts are provided AS IS without warranty  
'#of any kind. Microsoft further disclaims all implied warranties including,  
'#without limitation, any implied warranties of merchantability or of fitness for 
'#a particular purpose. The entire risk arising out of the use or performance of  
'#the sample scripts and documentation remains with you. In no event shall 
'#Microsoft, its authors, or anyone else involved in the creation, production, or 
'#delivery of the scripts be liable for any damages whatsoever (including, 
'#without limitation, damages for loss of business profits, business interruption, 
'#loss of business information, or other pecuniary loss) arising out of the use 
'#of or inability to use the sample scripts or documentation, even if Microsoft 
'#has been advised of the possibility of such damages 
'#--------------------------------------------------------------------------------- 

Const HKEY_LOCAL_MACHINE = &H80000002
Dim StrComputer,strKeyPath,strValueName
Dim objRegistry

strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\System\"
strValueName = "EnableSmartScreen"
objRegistry.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue


Dim RegKeyPath
Set objShell = CreateObject("Wscript.Shell")
regKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\System\EnableSmartScreen"

'determine if a registry key exists 
If IsNull(dwValue) Then
	'if the registry key does not exist, create a new registry key
	objShell.RegWrite regKeyPath,2,"REG_DWORD"
	WScript.Echo "Turn on SmartScreen successfully."
Else
	If dwValue = 2 Then
		WScript.Echo "You have already turn on SmartScreen successfully."
	Else
		objShell.RegWrite regKeyPath,2,"REG_DWORD"
		WScript.Echo "Turn on SmartScreen successfully."
	End If
End If


