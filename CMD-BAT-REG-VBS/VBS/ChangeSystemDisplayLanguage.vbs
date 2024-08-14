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

Dim LanguageTag,Flag
Dim objShell,objExecObject
Dim getInstalledLgp,strText,returnValue
Dim strKeyword
Dim InstalledLngTag,strTextLength,strKeywordLength,strLngTagLength

LanguageTag = Inputbox("Please enter an installed language tag that you want to change,such as ""en-US""")

If IsEmpty(LanguageTag) Then
	WScript.Quit
Else
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("Dism /online /Get-Intl")
	
	Flag = 0
	strKeyword = "Installed language(s):"

	Do While Not objExecObject.StdOut.AtEndOfStream
	    strText = objExecObject.StdOut.ReadLine()
	    returnValue = InStr(strText,strKeyword)
	    If returnValue > 0 Then
	    	'Get installed language tag
	    	strTextLength = Len(strText)
	    	strKeywordLength = Len(strKeyword)
	    	strLngTagLength = strTextLength - strKeywordLength 
	    	InstalledLngTag = LCase(Mid(strText,strKeywordLength+2,strLngTagLength))
	    	
	    	'Check if installed language exists
	    	If InstalledLngTag = LCase(LanguageTag) Then
	    		Flag = 1
	    	End If
	    End If
	Loop
	
	If Flag = 1 Then
		ChangeDisplayLanguage(LanguageTag)
	Else
		WScript.Echo "Please make sure you input the correct language tag or this language is installed."
	End If
End If


Function ChangeDisplayLanguage(LanguageTag)
	Dim strComputer
	Dim strValue
	Dim strKeyPath
	Dim objReg,str
	Const HKEY_CURRENT_USER = &H80000001
	
	strComputer = "."
	str = "PreferredUILanguages"
	strValue = LanguageTag
	strKeyPath = "Control Panel\Desktop"
	Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate}\\" & strComputer &"\root\default:StdRegProv")
	
	Dim strPath
	objReg.SetStringValue HKEY_CURRENT_USER,strKeyPath,str,strValue
	MsgBox "Successfully change the system display language." 
	'Prompt message

	Dim result,objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	result = MsgBox ("It will take effect after log off, do you want to log off right now?", vbYesNo, "Log off computer")
	
	Select Case result
	Case vbYes
		objShell.Run("logoff")
	Case vbNo
		Wscript.Quit
	End Select
End Function

