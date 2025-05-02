'Option Explicit

Dim objFSO, strTextFile, strData
CONST ForReading = 1
Dim strComputer, strOS, profile, list, OS
Dim colFolders, output, response, strAnswer
dim objShell, objPing, strPingOut

'Set objSh = CreateObject("Shell.Application")
'Set FSO =CreateObject("Scripting.FileSystemObject")
'strPath = FSO.GetParentFolderName(Wscript.ScriptFullName)
'objSh.ShellExecute "wscript.exe", strPath & "\LastLoggedOn.vbs", "", "runas", 1

start

function start
	strAnswer = InputBox("Please enter computer Name or (*) to use 'list.txt':", "Logged on user", "ComputerName")
	If strAnswer <> "*" Then
		strComputer = strAnswer
		'getOS(strAnswer)
		pingTest(strAnswer)
	Else
		'name of the text file
		strTextFile = "list.txt"
		readComputers
	End If
end function
	
function pingTest(strComputer)
    set objShell = CreateObject("Wscript.Shell")    
	set objPing = objShell.Exec("ping -n 1 " & strComputer)
	strPingOut = objPing.StdOut.ReadAll
	'msgbox strPingOut
    if instr(LCase(strPingOut), "host unreachable") or instr(LCase(strPingOut), "Request Timed Out")  then
		msgbox(strComputer & " --> Offline")
    else
        getOS(strComputer)
    end if
end function

function readComputers
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strData = objFSO.OpenTextFile(strTextFile,ForReading).ReadAll
	list = Split(strData,vbCrLf)
	Set objShell = CreateObject("WScript.Shell")
		
	For Each strComputer in list		
		pingTest(strComputer)
	Next	
end function

function getOS(strComputer)
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * From Win32_OperatingSystem")
	
	For Each objItem in colItems		
		If Instr(objItem.Caption, "Vista") Then
			profile = "\\" & strComputer & "\c$\users\"
			OS = "Vista"
		Else 
			profile = "\\" & strComputer & "\c$\documents and settings\"
			OS = "XP"
		End If
		getModDate(profile)
	Next	
end function

function getModDate(path)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(path)
	Set colSubfolders = objFolder.Subfolders

	For Each folder in colSubfolders
		tempModDate = folder.DateLastModified
		If tempModDate > modDate Then
			modDate = tempModDate
			name = folder.name
			'msgBox(strComputer & "(" & OS & ") - Last Logged on user: " & name)
		End If		
	Next	
	'wscript.echo(strComputer & "(" & OS & ") - Last Logged on user: " & name)
	'wscript.echo ""
	msgBox(strComputer & "(" & OS & ") - Last Logged on user: " & name)	
end function

intAnswer = msgBox("Would you like to run the script again?", vbYesNo, "Re-Run")
If intAnswer = vbYes Then
	start
Else
	'msgBox("Script Complete")	
	WScript.Quit
End if

