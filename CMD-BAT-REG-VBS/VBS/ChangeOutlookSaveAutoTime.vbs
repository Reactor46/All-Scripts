'--------------------------------------------------------------------------------- 
'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 
Option Explicit

Dim objOutlook
Dim OutlookVersion
Dim regKey
Dim dwValue
Dim strVersion

dwValue = InputBox("Please input the value that used to set AutoSaveTime interval in minutes.")

Set objOutlook = CreateObject("Outlook.Application")
OutlookVersion = objOutlook.Version
strVersion = Mid(OutlookVersion,1,4)

Select Case strVersion
	Case "12.0"	
			regKey = "12.0"
		SetRegValue
	Case "14.0"	
		regKey = "14.0"
		SetRegValue
	Case "15.0"	
		regKey = "15.0"
		SetRegValue
	Case Else 	WScript.Echo "The script can not find proper office version."
End Select

Sub SetRegValue
	Dim objRegistry
	Dim strKeyPath
	Dim strComputer
	Dim strValueName
	Const HKEY_CURRENT_USER = &H80000001
	
	strComputer = "."
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	
	'Define a key registry path
	strKeyPath = "Software\Microsoft\Office\" & regKey & "\Common\MailSettings"
	strValueName = "AutosaveTime"
	
	objRegistry.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
	WScript.Echo "Successfully set the AutoSaveTime registry value."
End Sub