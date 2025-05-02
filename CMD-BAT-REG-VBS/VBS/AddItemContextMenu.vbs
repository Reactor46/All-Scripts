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

On Error Resume Next
Dim strKeypath1,strKeyPath2
'Define a key registry path
strKeyPath1 = "Software\Classes\*\shellex\ContextMenuHandlers\PintoStartScreen"
addItemContextMenu(strKeyPath1)

strKeyPath2 = "Software\Classes\AllFileSystemObjects\shellex\ContextMenuHandlers\PintoStartScreen"
addItemContextMenu(strKeyPath2)			 
WScript.Echo "Add the 'Pin to Start' to context menu successfully."

Sub addItemContextMenu(strKeyPath)
Dim strComputer,objRegistry,strValueName,dwValue
	Const HKEY_CURRENT_USER = &H80000001
	strComputer = "."
	
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	objRegistry.CreateKey HKEY_CURRENT_USER, strKeyPath
	
	strValueName = ""
	dwValue = "{470C0EBD-5D73-4d58-9CED-E91E22E23282}"
	
	objRegistry.SetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
End Sub
