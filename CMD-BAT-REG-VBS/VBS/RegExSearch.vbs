'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://sites.google.com/site/assafmiron
' Date : 24/09/2009
' RegExSearch.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Option Explicit
On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim objFSO, objFile
Dim LogFile
Dim regEx
Dim strLine
Dim matches
Set regEx = New RegExp


Set objFSO = CreateObject("Scripting.FileSystemObject")

LogFile = "C:\MyLogFile.txt" ' Set the Log File Path

If objFSO.FileExists(LogFile) Then
	Set objFile = objFSO.OpenTextFile(LogFile,ForReading) ' Reading The Log File
	regEx.IgnoreCase = True
	regEx.Pattern = "(((Fox)?.(jump))|(ABC))|(over)"	
	regEx.Global = True
	
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		If regEx.Test(strLine) Then
			WScript.Echo strLine
			' If you Want to Echo the Matches
'			Set matches = regEx.Execute(strMSg)
'			For Each match In matches
'				WScript.Echo match
'			Next
		End if	
	Loop
	objFile.Close
	Set objFile = Nothing
Else
	WScript.Quit
End If

' Clean up
Set objFSO = Nothing
