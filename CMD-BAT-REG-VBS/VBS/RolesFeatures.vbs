'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
'   Author         : Syed Abdul Khader
'   Last Modified  : 14 April 2014
'   Description    : This script will generate the enabled Roles and Features in the Windows 2008 Server
'   Prerequisite   : ServerList.txt needs to be populated with targeted Windows 2008 server host name
'                    ServerList.txt files should be in location "C:\temp"		
'   Version        : 1.0
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
Option Explicit
Const ForAppending = 8
Const ForReading = 1

Dim objFSO, objtextfile, objoutputfile, srvlist, strcomputer, srvip, srvhostname, objWMIService, colRoleFeatures, objRoleFeatures, colrole, ExcelSheet
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile ("C:\temp\ServerList.txt", ForReading)
Set objOutputFile = objFSO.OpenTextFile("C:\temp\ServerRoles.csv", ForAppending, True)
Set SrvList = objFSO.OpenTextFile("ServerList.txt", ForReading)

objOutputFile.WriteLine("Report generated as of " & Now)
objOutputFile.writeLine("Server Name, Roles & Features")
wscript.echo "Please wait ..."

Do Until SrvList.AtEndOfStream
            strComputer = lcase(srvlist.readline)
            if checkserverResponse(strComputer) then
                        srvIP = localusers(strComputer)
            else
                        objoutputfile.write strComputer & " - Server is unreachable."
			objOutputFile.writeline " "
            end if
Loop

objTextFile.Close
objOutputFile.Close
Wscript.echo "Script completed" 
Set ExcelSheet = CreateObject("Excel.Application")
ExcelSheet.Application.Visible = True
ExcelSheet.Workbooks.Open ("C:\temp\ServerRoles.csv")

Function LocalUsers(strcomputer)
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colRoleFeatures = objWMIService.ExecQuery ("Select * from Win32_ServerFeature")
	colrole=isnull(colRoleFeatures)
	If colrole=false then
		For Each objRoleFeatures in colRoleFeatures
			objOutputFile.write strcomputer & "," & objrolefeatures.name
			objOutputFile.writeline " "
		Next
	Else
		objOutputFile.writeline " "
	End If
End Function

Function checkServerResponse(serverName)
	Dim Strtarget, objshell, objexec, strpingresults
	strTarget = serverName
	Set objShell = CreateObject("WScript.Shell")
	Set objExec = objShell.Exec("ping -n 1 -w 1000 " & strTarget)
	strPingResults = LCase(objExec.StdOut.ReadAll)
	If InStr(strPingResults, "reply from") Then
		checkServerResponse = true
	Else
		checkServerResponse = false
	End If
End Function