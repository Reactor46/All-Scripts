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


Sub main()
	'Main function 
	Content = "Select a action among follows:"_ 
			  & vbNewline & "SwitchUser"_
			  & vbNewline & "LogOff"_
			  & vbNewline & "Lock"_
			  & vbNewline & "Restart"_
			  & vbNewline & "Sleep"_
			  & vbNewline & "Hinermate"_
			  & vbNewline & "ShutDown"
	value = InputBox(Content,"Power Button Action")
	If IsEmpty(value) Then 
		WScript.Quit
	Else 
		Select Case UCase(value) 
		 	Case "SWITCHUSER"
		 		ChangeReg 256
			Case "LOGOFF"
				ChangeReg 1
			Case "LOCK" 
				ChangeReg 512
			Case "RESTART" 
				ChangeReg 4
			Case "SLEEP" 
				ChangeReg 16
			Case "HIBERNATE" 
				ChangeReg 64
			Case "SHUTDOWN" 
				ChangeReg 2
			Case Else 
				WScript.Echo "Invalid input value,please try again!"
				WScript.Quit
		End Select
		
		Call ChoicePromt
	End If 

End Sub 

Function ChangeReg(value)
	'Change the registry value
	On Error Resume Next 
	Set objshell = CreateObject("wscript.shell")
	'The registry value path
	RegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Start_PowerButtonAction"
	ReturnValue = objshell.RegRead(RegPath)
	If Err.Number = 0 Then 
		objshell.RegWrite Regpath, Value,"REG_DWORD"
	Else 
		objshell.RegWrite Regpath, Value,"REG_DWORD"
	End If 	
	Set objshell = Nothing 
End Function 

Sub ChoicePromt
	'reboot the computer
	Set objshell = CreateObject("wscript.shell")
	resultReboot = MsgBox("You may need to reboot of windows for the change to take effect. Do you want to reboot the computer right now?",vbYesNo+vbQuestion,"Reboot Computer")
	If resultReboot = 6 Then
		objShell.Exec("shutdown -r -t 0")
	End If
	Set objshell = Nothing 
End Sub

Call main 