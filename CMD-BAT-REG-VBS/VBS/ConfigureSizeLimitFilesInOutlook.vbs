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

'Define a key registry path
Dim regKeyPath,value,filesizeValue
value = Inputbox("Please enter the file size that you want to set,the unit is in 'MB'")
filesizeValue = value *1024*1024

Dim objRegistry,strValueName,dwValue,strComputer
If filesizeValue > 2147483648 Then
	WScript.Echo "The size you entered should be smaller than the maximum size of the file 2048(MB)."
Else
	Const HKEY_CURRENT_USER = &H80000001
	strComputer = "."
	
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	strKeyPath = "Software\Microsoft\Office\15.0\Outlook\PST"
	objRegistry.CreateKey HKEY_CURRENT_USER, strKeyPath
	
	strValueName = "MaxFileSize"
	dwValue = filesizeValue
	
	objRegistry.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue
	WScript.Echo "Set the value of properties of MaxFileSize successfully."
	
	'Call function
	Choice
End If

'Prompt message
Sub Choice
	Dim result

	result = MsgBox ("It will take effect after log off, do you want to log off right now?", vbYesNo, "Log off computer")
	
	Select Case result
	Case vbYes
		objShell.Run("logoff")
	Case vbNo
		Wscript.Quit
	End Select
End Sub	
