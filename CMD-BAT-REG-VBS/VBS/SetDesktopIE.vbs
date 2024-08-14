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

Option Explicit
On Error Resume Next
Dim objShell,regKeyPath
Dim AssociationActivationMode
Set objShell = CreateObject("Wscript.Shell")

regKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\AssociationActivationMode"
AssociationActivationMode = objShell.RegRead(regKeyPath)

If Err.Number = 0 Then
	'Setting the registry key
	If AssociationActivationMode = 2 Then
		WScript.Echo "You have already been set the desktop Internet Explorer mode."
	Else
		objShell.RegWrite regKeyPath,2,"REG_DWORD"
		WScript.Echo "Set the desktop Internet Explorer mode default successfully."
	End If
Else
	objShell.RegWrite regKeyPath,2,"REG_DWORD"
	WScript.Echo "Set the desktop Internet Explorer mode default successfully."
End If


