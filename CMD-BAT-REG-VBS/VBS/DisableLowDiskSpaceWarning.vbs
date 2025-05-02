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

'Define a key registry path
Dim regKeyPath,objShell,NoLowDiscSpaceChecksValue

Set objShell = CreateObject("Wscript.Shell")

regKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLowDiscSpaceChecks"
NoLowDiscSpaceChecksValue = objShell.RegRead(regKeyPath)

If Err.Number = 0 Then
	'Setting the registry key
	If NoLowDiscSpaceChecksValue = 1 Then
		WScript.Echo "You have already removed low disk space warning."
	Else
		objShell.RegWrite regKeyPath,1,"REG_DWORD"
		WScript.Echo "Setting to remove the low disk space warning successfully."
		'Call function
		Choice
	End If
Else
	objShell.RegWrite regKeyPath,1,"REG_DWORD"
	WScript.Echo "Setting to remove the low disk space warning successfully."
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
