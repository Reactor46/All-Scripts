'By K8l0
strPathTemplates = "\\contoso.com\SYSVOL\CONTOSO.COM\scripts\Templates_Office"

Const OverwriteExisting = TRUE 
Const HKEY_CURRENT_USER = &H80000001

Set oShell = CreateObject("Wscript.Shell")
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strUserProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
strExcelStart = oShell.ExpandEnvironmentStrings("%APPDATA%")


'EUA-EN
If objFSO.FolderExists(strUserProfile & "\AppData\Roaming\Microsoft\Templates") Then strPathOffice = strUserProfile & "\AppData\Roaming\Microsoft\Templates"
If objFSO.FolderExists(strExcelStart & "\Microsoft\Excel\XLSTART") Then strExcelStart = strExcelStart & "\Microsoft\Excel\XLSTART"
'BR-PT
If objFSO.FolderExists(strUserProfile & "\AppData\Roaming\Microsoft\Modelos") Then strPathOffice = strUserProfile & "\AppData\Roaming\Microsoft\Modelos"
If objFSO.FolderExists(strExcelStart & "\Microsoft\Excel\XLINICIO") Then strExcelStart = strExcelStart & "\Microsoft\Excel\XLINICIO"

If objFSO.FolderExists(strPathOffice & "\OFFICE-CT") Then
	objFSO.CopyFile strPathTemplates & "\*.*" , strPathOffice & "\OFFICE-CT", OverwriteExisting
	objFSO.CopyFile strPathOffice & "\OFFICE-CT\*.xl*" , strExcelStart, OverwriteExisting
	strKeyPath12 = "Software\Microsoft\Office\12.0\Common\General"
	strKeyPath14 = "Software\Microsoft\Office\14.0\Common\General"
	strValueName = "UserTemplates" 
	strValue = strPathOffice & "\OFFICE-CT"
	oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath12,strValueName,strValue
	oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath14,strValueName,strValue
Else
	objFSO.CreateFolder(strPathOffice & "\OFFICE-CT")
	objFSO.CopyFile strPathTemplates & "\*.*" , strPathOffice & "\OFFICE-CT", OverwriteExisting
	objFSO.CopyFile strPathOffice & "\OFFICE-CT\*.xl*" , strExcelStart, OverwriteExisting
	strKeyPath12 = "Software\Microsoft\Office\12.0\Common\General"
	strKeyPath14 = "Software\Microsoft\Office\14.0\Common\General"
	strValueName = "UserTemplates" 
	strValue = strPathOffice & "\OFFICE-CT"
	oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath12,strValueName,strValue
	oReg.SetStringValue HKEY_CURRENT_USER,strKeyPath14,strValueName,strValue
End If



