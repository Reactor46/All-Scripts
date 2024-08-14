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
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------


Dim count 
count = WScript.Arguments.Count 
Const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
'Get Wmi object. 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
'Registry key path 
strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\Personalization"

Select Case count 
	Case 0 
		'create a registry key
		objRegistry.CreateKey HKEY_LOCAL_MACHINE, strKeyPath
		objRegistry.SetDWORDValue  HKEY_LOCAL_MACHINE, strKeyPath, "NoLockScreen" , 1 
		WScript.Echo "Disable lock screen successfully."
	Case 1
		value = WScript.Arguments.Item(0)
		'Delete the key.
		If  UCase(value) = "ENABLE" Then 
			objRegistry.DeleteKey HKEY_LOCAL_MACHINE, strKeyPath
			WScript.Echo "Eanble lock screen successfully."
		Else
			WScript.Echo "Invalid value, try again."
		End If 
	Case Else 
		WScript.Echo "Invalid value, try again."
End Select 

