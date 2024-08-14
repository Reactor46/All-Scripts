'script will use wmi to do the following on a local computer
'Created by Atul Mishra
' Supported OS - Win XP, Win7, Win 2003 server, Win 2008 server
'1. Stop the Automatic Updates Service
'2. Delete the %WINDIR%\softwaredistribution folder
'3. Delete the HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate keys and subkeys.
'4. Start the Automatic updates Service
'5. Send a wuauclt.exe /resetauthorization /detectnow command to sysem
'6. Initiate software update scan cycle actions
dim objFSO, objShell, objTempFile, objTS
dim sCommand, sReadLine
dim oCPAppletMgr 'Control Applet manager object.
dim oClientAction 'Individual client action.
dim oClientActions 'A collection of client actions.
'dim bReturn
set objShell = WScript.CreateObject("Wscript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject")
strComputer = "."
'----------------Stop Automatic Updates Service--------------------
'
'
'On Error Resume Next
' NB strService is case sensitive.
strService = " 'wuauserv' "
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery _
("Select * from Win32_Service Where Name ="_
& strService & " ")
For Each objService in colListOfServices
objService.StopService()
Next 
WScript.Echo "Service has been stopped" 
'
'
'
'----------------- Delete Folder and Reg Keys --------------------------------
strExe = "cmd.exe /C rmdir %WINDIR%\SoftwareDistribution /S /Q && cmd.exe /C REG DELETE HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate /f"
' Connect to WMI
'
set objWMIService = getobject("winmgmts://"_
& strComputer & "/root/cimv2") 
' Obtain the Win32_Process class of object.
Set objProcess = objWMIService.Get("Win32_Process")
Set objProgram = objProcess.Methods_( _
"Create").InParameters.SpawnInstance_
objProgram.CommandLine = strExe 
'Execute the program now at the command line.
Set strShell = objWMIService.ExecMethod( _
"Win32_Process", "Create", objProgram)
WScript.Echo "The software Distribution Folder and WindowsUpdate regsitry keys have been deleted." 
'------------------ Start Automatic Update Service -------------
'On Error Resume Next
' NB strService is case sensitive.
strService = " 'wuauserv' "
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery _
("Select * from Win32_Service Where Name ="_
& strService & " ")
For Each objService in colListOfServices
objService.StartService()
Next 
WScript.Echo "Service has been Started on " & strcomputer 

'-------------- Force Checking to WSUS server by issueing a wuauclt.exe /resetauthorization /detectnow ------------

strExe = "cmd.exe /C wuauclt.exe /resetauthorization /detectnow"
' Connect to WMI
'
set objWMIService = getobject("winmgmts://"_
& strComputer & "/root/cimv2") 
' Obtain the Win32_Process class of object.
Set objProcess = objWMIService.Get("Win32_Process")
Set objProgram = objProcess.Methods_( _
"Create").InParameters.SpawnInstance_
objProgram.CommandLine = strExe 
'Execute the program now at the command line.
Set strShell = objWMIService.ExecMethod( _
"Win32_Process", "Create", objProgram)
WScript.Echo "Force checkin has been sent. Process Complete." 
'Initiate software update scan cycle actions
'Get the Control Panel manager object.
set  oCPAppletMgr=CreateObject("CPApplet.CPAppletMgr")
if err.number <> 0 then
    Wscript.echo "Couldn't create control panel application manager" 
    WScript.Quit
end if
'Get a collection of actions.
set oClientActions=oCPAppletMgr.GetClientActions
if err.number<>0 then
    wscript.echo "Couldn't get the client actions"
    set oCPAppletMgr=nothing
    WScript.Quit
end if
'Display each client action name and perform it.
For Each oClientAction In oClientActions
    if oClientAction.Name = "Software Updates Assignments Evaluation Cycle" then
        wscript.echo "Performing action " + oClientAction.Name 
        oClientAction.PerformAction
    end if
next
set oClientActions=nothing
set oCPAppletMgr=nothing