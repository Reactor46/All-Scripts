' ----------------------------------------------------------------------------------------------------------------'
' Outlook - Configure to use new CAS Array for Exchange 2010 Migration                                            ' 
' ================================================================================================================'
' Script Name ......: Outlook2007_CASArray_Config.vbs                                             '
' Version History ..: 1.0 19-March-2014 - Initial Release, Author - Jamie Powell                                  '
' Description.......: This script has been written to assist with the transition to the new Exchange CAS Array as '
'                     part of the Exchange 2010 Migration. The following functions are performed in this script:  '
'                     1. Find the Default Outlook Profile name on the running computer via the Registry;          '
'                     2. Remove 4 Registry String Values that contain the old CAS Server name of "SVWWMX001"      '
'                        (removing these 4 strings forces the Outlook client to AutoDiscover the new CAS Array);  '
'                     3. Reboot Workstation
' ----------------------------------------------------------------------------------------------------------------'
'*****************************************************************************************************************'
'-----------------------------Set Constants and DIMs--------------------------------------------------------------'
Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

'-----------------------------Find Default Outlook Profile Name---------------------------------------------------
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
strValueName = "DefaultProfile"
oReg.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
strRegPath = "HKCU" & "\" & strKeyPath & "\" & strValue

'-----------------------------Remove Reg String Values to force AutoDiscover--------------------------------------
Set objRegistry=GetObject("winmgmts:\\" & _ 
    strComputer & "\root\default:StdRegProv")
	
strKeyPath1 = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\" & strValue & "\" & "13dbb0c8aa05101a9bb000aa002fc45a"

strValueName1 = "001e6602"
strValueName2 = "001e6603"
strValueName3 = "001e6608"
strValueName4 = "001e6612"

objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath1, strValueName1
objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath1, strValueName2
objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath1, strValueName3
objRegistry.DeleteValue HKEY_CURRENT_USER, strKeyPath1, strValueName4
'-----------------------------Reboot Workstation------------------------------------------------------------------
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Shutdown)}!\\" & _
        strComputer & "\root\cimv2")
		
Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
	
intAnswer = _
    Msgbox("Can I reboot your PC to complete the Outlook Change?", _
        vbYesNo, "Reboot PC")

If intAnswer = vbYes Then
    Msgbox "Rebooting Now."
	For Each objOperatingSystem in colOperatingSystems
		objOperatingSystem.Reboot()
	Next
Else
    Msgbox "Please reboot at your earliest convenience."
End If
'-----------------------------Complete the script and tidy up-----------------------------------------------------
Set WSShell = Nothing
Set objShell = Nothing
WScript.Quit
