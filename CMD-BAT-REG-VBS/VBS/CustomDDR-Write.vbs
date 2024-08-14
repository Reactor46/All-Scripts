' Script to generate a DDR destined for the ConfigMgr Site servers inbox\DDM.BOX folder to modify an existing Clients Resource record

' By Robert Marshall - SMSMarshall (2014)

' Version 1.0

' Inputs - DDRProperty, DDRPropertyType, DDRClientGUID, DDRSitecode, DDRPropertyLength, DDRDestPath, DDRAgentName, DDRFileName

On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile("c:\temp\rabbit.txt",True)

objFile.Write "test string" & vbCrLf

'objFile.Close

Set colNamedArguments = WScript.Arguments.Named

	Const ADDPROP_NONE = &H0
	Const ADDPROP_GUID = &H2
	Const ADDPROP_KEY = &H8
	Const ADDPROP_NAME = &H44

	If colNamedArguments.Item("DDRProperty") = False Then

		WScript.Quit(2) ' Required argument 'DDRProperty' is missing

	End If

	If colNamedArguments.Item("DDRPropertyType") = False Then

		WScript.Quit(3) ' Required argument 'DDRPropertyType' is missing

	End If

	If colNamedArguments.Item("DDRClientGUID") = False Then

		WScript.Quit(4) ' Required argument 'ClientGUID' is missing

	End If

	If colNamedArguments.Item("DDRSiteCode") = False Then

		WScript.Quit(5) ' Required argument 'SiteCode' is missing

	End If

	If colNamedArguments.Item("DDRPropertyLength") = False Then

		WScript.Quit(6) ' Required argument 'DDRPropertyLength' is missing

	End If

	If colNamedArguments.Item("DDRDestPath") = False Then

		WScript.Quit(7) ' Required argument 'DDRDestPath' is missing

	End If

	If colNamedArguments.Item("DDRAgentName") = False Then

		WScript.Quit(8) ' Required argument 'DDRAgentName' is missing

	End If

	If colNamedArguments.Item("DDRFileName") = False Then

		WScript.Quit(8) ' Required argument 'DDRFileName' is missing

	End If

	If colNamedArguments.Item("DDRPropertyName") = False Then

		WScript.Quit(9) ' Required argument 'DDRPropertyName' is missing

	End If

' Register the smsrsgenctl DLL from the ConfigMgr 2012 R2 SDK

	Err.Clear

	Return = WshShell.Run("regsvr32 -s " & colNamedArguments.Item("DDRCWD") & "\smsrsgenctl.dll", 0, true)

	objFile.Write "No1: " & Err.Description & " - " & Err.Number

' Load an instance of the SMSResGen.dll.

	Err.Clear

	Set newDDR = CreateObject("SMSResGen.SMSResGen.1")

	objFile.Write "No2: " & Err.Description & " - " & Err.Number

	' Create a new DDR using the DDRNew method.

	newDDR.DDRNew "System", colNamedArguments.Item("DDRAgentName"), GetSiteCode

' Create the DDR Property

	if colNamedArguments.Item("DDRPropertyType") = "String" Then ' String

		newDDR.DDRAddString colNamedArguments.Item("DDRPropertyName"), MID(colNamedArguments.Item("DDRProperty"),1,colNamedArguments.Item("DDRPropertyLength")), colNamedArguments.Item("DDRPropertyLength"), ADDPROP_NONE

	ElseIf colNamedArguments.Item("DDRPropertyType") = "Integer" Then ' Integer

		newDDR.DDRAddInteger colNamedArguments.Item("DDRPropertyName"), colNamedArguments.Item("DDRProperty"), ADDPROP_NONE

	ElseIf colNamedArguments.Item("DDRPropertyType") = "Array" Then ' Array of Strings

		modifiedArray = Split(colNamedArguments.Item("DDRProperty"),"$$$$",-1,1)

		newDDR.DDRAddStringArray colNamedArguments.Item("DDRPropertyName"), modifiedArray, colNamedArguments.Item("DDRPropertyLength"), ADDPROP_NONE

	End If

	' Specify the SMS Unique Identifier

	newDDR.DDRAddString "SMS Unique Identifier", colNamedArguments.Item("DDRClientGUID"), 64, ADDPROP_GUID And ADDPROP_KEY

	' Write new DDR to file but check to see if there is a trailing \ and handle it

	If MID(colNamedArguments.Item("DDRDestPath"),LEN(colNamedArguments.Item("DDRDestPath")),1) = "\" Then 

		newDDR.DDRWrite colNamedArguments.Item("DDRDestPath") & colNamedArguments.Item("DDRFileName")

	Else
		newDDR.DDRWrite colNamedArguments.Item("DDRDestPath") & "\" & colNamedArguments.Item("DDRFileName")

	End If

' Unregister the smsrsgenctl DLL

	Set WshShell = WScript.CreateObject("WScript.Shell")

	Return = WshShell.Run("regsvr32 -u -s " & colNamedArguments.Item("DDRCWD") & "\smsrsgenctl.dll", 0, true)

' Simple!