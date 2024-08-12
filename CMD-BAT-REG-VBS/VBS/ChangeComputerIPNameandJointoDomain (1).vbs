'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' Date : 01/05/06
'=*=*=*=*=*=*=*=*=*=*=*=

'======================================================================================================================
'This script Changes the IP address and computer name of the Current computer (Never tested on remote computers)
'After doing so it joins the computer to the domain in a selected OU
'All the script is loged
'The Script uses the NewSID program of SysInternals and the NetDom program found in the Resource Kit of Windows
'Joining a computer to the domain works with no problem on XP Machines, on 2000 Machines you need to Join them with the NetDom command
'You can still use all the Vars you define here
'======================================================================================================================

strComputer = "."
Set WshShell = Wscript.CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
WinDir = WshShell.ExpandEnvironmentStrings("%WINDIR%")
strLogDir = WinDir & "\system32\dllCache\"
UserProfile = WshShell.ExpandEnvironmentStrings("%userprofile%")


'=*=*=*=*=*=
'=Functions=
'=*=*=*=*=*=

Function ShutDown(strComputer)
'======================================================================================================================
'This Function first closes all Office Program if open and then shuts down the computer
'It receives strComputer as the Computer Name and shuts it down
'======================================================================================================================
Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Process")
'======================================================================================================================
' First Try to Restart the Computer with a 30 Seconds timeout
'======================================================================================================================
E = objWMIService.Create("cmd /c shutdown -r -t 30 ", null, null, intProcessID) 

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
'======================================================================================================================
'Close all Microsoft Office Programs
'======================================================================================================================
'Outlook
'======================================================================================================================
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'OUTLOOK.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
'======================================================================================================================
'Excel
'======================================================================================================================
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Excel.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
'======================================================================================================================
'Word
'======================================================================================================================
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'WinWORD.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
'======================================================================================================================
'Power Point
'======================================================================================================================
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'PowerPnt.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
'======================================================================================================================
'Wait for 30 seconds
'======================================================================================================================
wscript.sleep 30000
'======================================================================================================================
'Restart the computer using WMI
'======================================================================================================================
Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" & strComputer & "\root\cimv2")
	Set OperatingSystems = WMIService.ExecQuery("Select * From Win32_OperatingSystem")
	For Each OperatingSystem in OperatingSystems
		OperatingSystem.Reboot()
	Next
if E = 0 then
  ShutDown  = "OK"
end if

End Function

'=*=*=*=
'=Subs =
'=*=*=*=

Sub ChangeSID
'======================================================================================================================
'This sub uses the SysInternals newSID program to generate a new SID to the computer
'NewSID is Avialable Here: \\live.sysinternals.com\Tools\newsid.exe
'This script asumes that the program is in a mapped drive (Z)
'You can use another script that maps it with Subst to a local folder..
'If your using SysPrep you dont need this Sub at all
'======================================================================================================================

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Process")

Error = objWMIService.Create("cmd /c z:\newsid.exe /a" , null, null, intProcessID)
If Error = 0 Then
    objFile.WriteLine "SID Changer was started with a process ID of " _
         & intProcessID & "."
    Wscript.echo "SID Changer is Running." & vbclrf & " This may take a few minutes, Please wait"
Else
    objFile.WriteLine "SID Changer could not be started due to error " & _
        Error & "."
End If

End Sub

Sub ChangeIP(Seg,IP)
'======================================================================================================================
'This Sub cahnges the IP of the computer entered in the strComputer variable
'======================================================================================================================
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
strIPAddress = Array("10.5." & Seg & "." & IP)
strSubnetMask = Array("255.255.255.0")
strGateway = Array("10.5." & Seg & ".254")
strGatewayMetric = Array(1)
 
For Each objNetAdapter in colNetAdapters
    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
    errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
    If errEnable = 0 Then
        ObjFile.WriteLine "The IP address has been changed."
    Else
        ObjFile.WriteLine "The IP address could not be changed."
    End If
Next

End Sub

Sub ChangeCompName(Seg,IP)
'======================================================================================================================
'This Sub Changes the Computer name according to a Constant Value and the IP address enterd
'It Changes the ComputerName key in the Registry and some other keys to do so
'======================================================================================================================
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName"
strValueName = "ComputerName"
strValue = "CompName" & seg & IP

oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue


strKeyPath = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
strValueName = "Hostname"
strValue = "CompName" & seg & IP

oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

strKeyPath = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
strValueName = "NV Hostname"
strValue = "CompName" & seg & IP

oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue


objFile.WriteLine "Computer Name: " & strValue
'======================================================================================================================
'Write the New computer name to the Log
'======================================================================================================================
End Sub

Sub Join2Domain()
'======================================================================================================================
'This Sub Joins the computer to the domain we enter here
'You can also designate a spesific OU for the Computer Account in the strOU Const
'This command works only on XP if you want to use it on Win2K you need to use the NetDom command that exist in windows Res Kit
'you can use the same Constants but a different syntax is required
'======================================================================================================================
Const JOIN_DOMAIN             = 1
Const ACCT_CREATE             = 2
Const ACCT_DELETE             = 4
Const WIN9X_UPGRADE           = 16
Const DOMAIN_JOIN_IF_JOINED   = 32
Const JOIN_UNSECURE           = 64
Const MACHINE_PASSWORD_PASSED = 128
Const DEFERRED_SPN_SET        = 256
Const INSTALL_INVOCATION      = 262144
Const strOU		      = "ou=OUComputers, ou=test,  ou=OUProjects, dc=Domain, dc=Com"
'======================================================================================================================
'Enter a valid User Name and password to join the computer to the Domain
'The Password is saved ClearText
'======================================================================================================================
strDomain   = "Domain"
strPassword = "ClearTextPassword"
strUser     = "MyUserName"
 
'======================================================================================================================
'This is a Check Message
'======================================================================================================================
wscript.echo "Trying to join to domain " & strDomain

' Check Computer OS
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "ProductName"

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
If InStr(strValue,"XP") Then
	' On WinXP Computers
	Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & _
	    strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & _
	        strComputer & "'")

	ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, _
                                                strPassword, _
                                                strDomain & "\" & strUser, _
                                                strOU, _
                                                JOIN_DOMAIN + ACCT_CREATE)
Else
	' On Win 2000 Computers
	ReturnValue = WshShell.run "NetDom add " & strComputer & " /Domain:" & strDomain _
		& " /userD:" & strDomain & "\" & strUser & " /PasswordD: " _
		& strPassword & " /ou:" & strOU
End If
'======================================================================================================================
'Some Error Handling
'Any Error is written to the Log File
'======================================================================================================================
If ReturnValue = 5 then
	objFile.WriteLine "Access was Denied for adding the Computer to the Domain"
End If
If ReturnValue = 2224 then
	objFile.WriteLine "Account exists in the Domain"
End If
if ReturnValue > 2691 then
	objFile.WriteLine "The Computer was sucssefully added to the domain"
Else
	objFile.WriteLine "The Computer was NOT added to the domain"
End if

End Sub

'=*=*=*=
'=Code =
'=*=*=*=
'======================================================================================================================
'Here Begins our script and Calles the Subs and Functions
'First it checks if the Log file Exists, if it does it begins the process of joining it to the Domain and ending the Log file
'Else it creates the file and asks for IP address - the IP address is used to cahnge the IP of the Computer and its Name
'The Script will keep asking for the IP in the correct Order until it have it. If you want to stop it just End the Wscript process
'======================================================================================================================
If objFSO.FileExists(strLogDir & "Domain Log.txt") Then
	Set objFile = objFso.OpenTextFile(strLogDir & "Domain Log.txt",ForWriting)
	Join2Domain
	objFile.Writeline "Log Ended :" & NOW
	objFile.close
Else
	fname=strLogDir & "Domain Log.txt"
	Set objFile = objFso.CreateTextFile (fname, True)

	objFile.WriteLine "Log Started : " & Now
	objFile.WriteLine
	IP = inputBox("Enter Computer IP","Enter Computer IP Between X.1 - 253")
	Do Until InStr(IP,".") 
	'======================================================================================================================
	'Keep Looping until the enterd text will have a " . " in it
	'======================================================================================================================
		IP = inputBox("Enter Computer IP","Enter Computer IP Between X.1 - 253")
	Loop
	arrIP = Split(IP,".")
	'======================================================================================================================
	'Divide the IP entered to Segment and IP address
	'======================================================================================================================
	Seg = arrIP(0)
	IP = arrIP(1)
	'======================================================================================================================
	'Change the IP and the Computer Name accourding to it
	'======================================================================================================================
	ChangeIP Seg,IP
	ChangeCompName Seg,IP
	ChangeSID
	'======================================================================================================================
	'Here we write in the Log file the Computer Booted and Close the File before the computer realy boots to k eep the file from Corrupting
	'======================================================================================================================
	objFile.Writeline "Computer Booted :" & NOW
	objFile.close

End If
'======================================================================================================================
'Restart the Computer !
'======================================================================================================================
ShutDown strComputer