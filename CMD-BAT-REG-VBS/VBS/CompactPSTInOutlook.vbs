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
Dim regVersionKey
Dim dwValue
Dim strVersion

Set objOutlook = CreateObject("Outlook.Application")
OutlookVersion = objOutlook.Version
strVersion = Mid(OutlookVersion,1,4)

Select Case strVersion
	Case "12.0"	
		regVersionKey = "12.0"
		SetRegValue
	Case "14.0"	
		regVersionKey = "14.0"
		SetRegValue
	Case "15.0"	
		regVersionKey = "15.0"
		SetRegValue
	Case Else WScript.Echo "The script can not find proper office version."
End Select

Set objOutlook = Nothing

Sub SetRegValue
	Dim objRegistry
	Dim strKeyPath
	Dim strComputer
	Dim strValueName
	Const HKEY_CURRENT_USER = &H80000001
	
	strComputer = "."
	Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
	' sets the value of registry
	dwValue = 1
	
	'Define a key registry path
	strKeyPath = "Software\Microsoft\Office\" & regVersionKey & "\Outlook"
	strValueName = "PSTNullFreeOnClose"
	
	objRegistry.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
	WScript.Echo "Successfully enabled Outlook to compact the .PST file every time."
End Sub