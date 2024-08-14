'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: disableLGPO.vbs
'
' AUTHOR: Darren Mar-Elia , SDM Software, Inc.
' DATE  : 2/20/2007
'
' COMMENT: This script disables the local GPO in pre-Vista systems
'
'==========================================================================
On Error Resume Next
Const ForReading = 1
Const ForWriting = 2
Set WshShell = WScript.CreateObject("WScript.Shell")
' get the current system folder
sysDir = WshShell.ExpandEnvironmentStrings("%WinDir%")
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Err.Number <>0 Then
	WScript.Echo "Unable to create FileSystemObject"
	Err.Clear
	WScript.Quit(1)
End If
Set objFile = objFSO.OpenTextFile(sysDir+"\system32\grouppolicy\gpt.ini", ForReading)
If Err.Number <>0 Then
	WScript.Echo "Unable to open gpt.ini for reading"
	Err.Clear
	WScript.Quit(1)
End If
counter = 0
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    If InStr(strLine,"Options=") > 0 Then
        strLine = "Options=3"
        counter=counter+1
    End If
    strContents = strContents & strLine & VbCrLf
Loop
'if we didn't find options= in the file then append it
If counter =0 Then
	strContents=strContents+"Options=3"+VbCrLf
End If

objFile.Close

Set objFile = objFSO.OpenTextFile(sysDir+"\system32\grouppolicy\gpt.ini", ForWriting)
If Err.Number <>0 Then
	WScript.Echo "Unable to open gpt.ini for writing"
	Err.Clear
	WScript.Quit(1)
End If
objFile.Write(strContents)
objFile.Close
WScript.Echo "Local GPO Disabled Successfully"


	

	


