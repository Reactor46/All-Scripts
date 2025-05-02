'=*=*=*=*=*=*=*=*=*=*=*=*=
' Author : Assaf Miron 
' Http://assaf.miron.googlepages.com
' Date : 17/08/10
' ParseGPResult.vbs
' Description : This Script will Parse a GPResult output file
' The Script will return when the GPO was Applied 
' and an Array of Applied and Filterd GPO
'=*=*=*=*=*=*=*=*=*=*=*=*=
Option Explicit
On Error Resume Next
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim objFSO, objTextFile
Dim strText, strLine, strFile
Dim arrAppliedGPO(), arrFilteredGPO()
Dim i

Set objFSO = CreateObject("Scripting.FileSystemObject")

strFile = "C:\GPOResult-2003.txt" ' Set the File Path

If strFile <> "" Then ' Check that File Path is Not Empty
	If objFSO.FileExists(strFile) Then ' Check if the File Exsists allready
		Set objTextFile = objFSO.OpenTextFile(strFile,ForReading) ' Reading The Log File
	End If
	Do Until objTextFile.AtEndOfStream
		strLine = objTextFile.ReadLine
		If InStr(strLine,"Last time Group Policy was applied:") Then
			WScript.Echo GetText(strLine,"Last time Group Policy was applied:")
		End If
		' Find Applied GPOs
		If InStr(strLine,"Applied Group Policy Objects") Then
			' Skip the -------- Line
			strLine = objTextFile.ReadLine
			i = 0
			' read all the Applied GPOs
			strLine = objTextFile.ReadLine
			Do Until strLine = "" 
				ReDim Preserve arrAppliedGPO(i)
				arrAppliedGPO(i) = Trim(strLine)
				i = i + 1
				strLine = objTextFile.ReadLine
			Loop
		End If
		' Find Filterd GPOs
		If InStr(strLine,"The following GPOs were not applied because they were filtered out") Then
			' Skip the -------- Line
			strLine = objTextFile.ReadLine
			i = 0
			Dim strFilter
			Dim strNextLine
			' read all the Filtered GPOs
			strLine = objTextFile.ReadLine
			Do Until strLine = "" 
				strNextLine = objTextFile.ReadLine
				If InStr(strNextLine,"-------") = 0 Then
					strFilter = ""
					Do Until strNextLine = ""
						strFilter = strFilter & GetText(strNextLine,"Filtering:")
						strNextLine = objTextFile.ReadLine
					Loop
					ReDim Preserve arrFilteredGPO(i)
					arrFilteredGPO(i) = Trim(strLine) & "(" & strFilter & ")"
					i = i + 1
				Else
					Exit do
				End If
				strLine = objTextFile.ReadLine
			Loop
		End If
	Loop
	WScript.Echo "There are " & UBound(arrAppliedGPO)+1 & " Applied GPOs."
	WScript.Echo "There are " & UBound(arrFilteredGPO)+1 & " GPOs that were Filtered Out."
	' Close the File
	objTextFile.Close
	' Free the Object
	Set objTextFile = Nothing
Else ' Path is Empty
	WScript.Quit ' End Script
End If

'******************************************************************************
' Description	: Checks if a String is Empty and Returns a Default Value if is Empty
' Input			: The String to Check, Default Value
' Output		: Source String if Not Empty, Defualt Value if Empty
Function CheckEmpty(strCheck, defValue)
	If IsNull(strCheck) Then
		CheckEmpty = defValue
	ElseIf strCheck = "" Or strCheck = " " Then
		CheckEmpty = defValue
	Else
		CheckEmpty = Trim(Replace(strCheck,vbCr,""))
	End If
End Function

'******************************************************************************
' Description	: Extracts a String from Text
' Input			: String Text to look into, String to Find
' Output		: The Following Scentence after the String you Searched
Function GetText(strText,strToFind)
	On Error Resume Next ' In Case String is NULL
	Dim objRegEx
	Dim colMatches
	Dim retText
	Dim i : i = 0
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True   
	objRegEx.IgnoreCase = True 
	objRegEx.MultiLine = True
	objRegEx.Pattern = strToFind & "[ -:](.*)"
	Set colMatches = objRegEx.Execute(strText)
	' Check That The RegEx Has Matches
	If colMatches.Count > 0 Then
		' Check if The First Match has Sub Matches
		If colMatches(0).SubMatches.Count > 0 Then
			' Check if Value is not Empty
			While Trim(CheckEmpty(colMatches(0).SubMatches(i),"-"))	= "-"
				i = i + 1
			Wend
			' Get The First Match in the Sub Matches Collection
			retText = Trim(colMatches(0).SubMatches(i))
		Else
			' No Sub Matches - Get The Value of the First Match
			retText = Trim(colMatches(0).Value)
		End If
	Else
		' Return Null - No Matches Found
		retText = Null
	End If
	' Clean Up
	Set objRegEx = Nothing
	GetText = Trim(retText)
End Function