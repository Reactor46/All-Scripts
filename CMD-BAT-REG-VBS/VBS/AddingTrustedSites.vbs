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

Dim str
Dim Id
Dim strPrefix
Dim strDomains,strDomain
str = Inputbox("Please enter the one or more domains,such as ""contoso1.com"",""contoso2.com""")
strDomains = Split(str,",")

Id = InputBox("Which prefix do you want to add? Please enter the corresponding number" & vbCrLf &_
					"1: The domain will be use the http:// prefix." & vbCrLf &_
					"2: The domain will be use the https:// prefix.")
Select Case Id
	Case "1" strPrefix = "http"
		AddTrustedSitesReg strDomains,strPrefix
	Case "2" strPrefix = "https"
		AddTrustedSitesReg strDomains,strPrefix
	Case Else MsgBox "You input the wrong option, please enter again."
End Select


Function AddTrustedSitesReg(strDomains,strPrefix)
	Dim strComputer
	Dim dwValue
	Dim strKeyPath
	Dim objReg
	Const HKEY_CURRENT_USER = &H80000001
	
	strComputer = "."
	dwValue = 2
	strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\"
	Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate}\\" & strComputer &"\root\default:StdRegProv")
	
	Dim strPath
	For Each strDomain in strDomains
		strPath = strKeyPath & "\" & strDomain
		objReg.CreateKey HKEY_CURRENT_USER,strPath
		objReg.SetDWORDValue HKEY_CURRENT_USER,strPath,strPrefix,dwValue
		MsgBox "Successfully added " & strDomain & " to trusted sites in Internet Explorer." 
	Next
End Function