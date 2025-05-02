'********************************************************************
' Language VBS 
' App/Proccess/Service.... Starter
' This script can start application's on machines from other domain's
' so proccesses and utilites whit cmd line, like *bat code
' By Roma Krilov@26/10/2011
' Version 1.0.0
' *******************************************************************

Option Explicit

Dim strShell, objProgram, strComputer, strExe,objSWbemLocator,objWMIService, objProcess, strInput

' App Name InputBox here
strExe = InputBox("Enter AppName or Utillite Name\Command Line, you wish to start","Remote App Starter")
If strExe = "" Then
	Wscript.Quit
End If

' Program name InputBox here
Do
 strComputer = InputBox("Enter Server Ip Address or NBTName or FQDN","Remote App Starter")
	If StrComputer <> "" Then
	strInput = True 
	End If
Loop until strInput = True

' Connect to WMI
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
' Conect to Server\Computer in other domain, MS_409 is system local by default(US-English).
Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     "UserNameHere", _
     "PasswordHere", _
     "MS_409", _
     "ntlmdomain:MyDomain.com")
	
' Obtain the Win32_Process class of object.
Set objProcess = objWMIService.Get("Win32_Process")
Set objProgram = objProcess.Methods_( _
"Create").InParameters.SpawnInstance_
objProgram.CommandLine = strExe 

'Execute the program now at the command line.
Set strShell = objWMIService.ExecMethod( _
"Win32_Process", "Create", objProgram) 

WScript.Sleep 20000

WScript.echo "Executed Successfully: " & strExe & " on " & strComputer

Set strShell = Nothing : Set objProgram = Nothing : Set strComputer = Nothing : Set strExe = Nothing
Set objSWbemLocator = Nothing : Set objWMIService = Nothing : Set objProcess = Nothing : Set strInput = Nothing
