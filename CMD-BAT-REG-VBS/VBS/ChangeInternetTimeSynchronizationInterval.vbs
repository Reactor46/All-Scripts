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

Dim regKeyPath
Dim optValue
Dim value
Dim dwValue
optValue = Inputbox("ID  TimeTypes" & vbCrLf &_                                                                            
                    "--  -----------" & vbCrLf &_                                                                          
                    "1   (Minutes)" & vbCrLf &_                                                                                  
                    "2   (Hours)" & vbCrLf &_                                                                                  
                    "3   (Days)" & vbCrLf & vbCrLf &_                                                                                  
"Which type of time do you want to set for Internet Time Synchronization? Input the corresponding ID number:")

If IsEmpty(optValue) = True Then
	WScript.Quit
Else
	If optValue = 1 Or optValue = 2 Or optValue = 3 Then
		value = InputBox("Please input the time you want to set for Internet Time Synchronization:")
		
		Select Case optValue
		Case 1
			dwValue = value *60
			SetTime
		Case 2
			dwValue = value *3600
			SetTime
		Case 3
			dwValue = value *86400
			SetTime
		End Select
	Else
		WScript.Echo "Input incorrect, please follow the prompt and enter the number."
	End If
End If
	
Sub SetTime
	Dim objRegistry
	Dim strKeyPath
	Dim strValueName
	Dim strComputer
	
	Const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."
	
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	
	'Define a key registry path
	strKeyPath = "SYSTEM\CurrentControlSet\Services\w32Time\TimeProviders\NtpClient"
	objRegistry.CreateKey HKEY_LOCAL_MACHINE, strKeyPath
	strValueName = "SpecialPollInterval"
	
	objRegistry.SetDWORDValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, dwValue
	WScript.Echo "Set the Internet Time Synchronization successfully."
End Sub
