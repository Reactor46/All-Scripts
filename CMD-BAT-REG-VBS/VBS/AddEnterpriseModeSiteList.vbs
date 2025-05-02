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

Dim strSiteList
strSiteList = Inputbox("Spcifies the xml file that contains about the websites you want to enabled on Enterprise Mode.")

If strSiteList="" Then
	WScript.Echo "You must specify a location of xml file."
Else
	AddEnterpriseModeSiteList(strSiteList)
End If

'The function which can be used to add site list to Enterprise Mode.
Function AddEnterpriseModeSiteList(strSiteList)
	Dim strComputer
	Dim strName
	Dim strPath
	Dim strKeyPath
	Dim StringValue
	Dim objReg
	Const HKEY_CURRENT_USER = &H80000001
	
	strComputer = "."

	Set objReg = GetObject("winmgmts:{impersonationlevel=impersonate}\\" & strComputer &"\root\default:StdRegProv")
	

	strValue = strSiteList
	strName = "SiteList"
	strKeyPath = "Software\Policies\Microsoft"
 	strSubPath = "Internet Explorer\Main\EnterpriseMode"
	strPath = strKeyPath & "\" & strsubPath
	objReg.CreateKey HKEY_CURRENT_USER,strPath
	objReg.SetStringValue HKEY_CURRENT_USER,strPath,strName,strValue
	MsgBox "Successfully added the site list to Internet Explorer Enterprise Mode." 
End Function