'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages.

Option Explicit
Dim objWMILocator,objNetwork
Set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
Set objNetwork = CreateObject("WScript.Network")

Function ConvertDMTFDateTime(DMTFDateTime)
	'This function converts DMTF Date-Time values 
    ConvertDMTFDateTime = CStr(Mid(DMTFDateTime, 5, 2) & "/" & _
        Mid(DMTFDateTime, 7, 2) & "/" & Left(DMTFDateTime, 4) _
            & " " & Mid (DMTFDateTime, 9, 2) & ":" & _
                Mid(DMTFDateTime, 11, 2) & ":" & Mid(DMTFDateTime, _
                    13, 2))
End Function

Function GetOSCUserLogonDuration(ComputerName,UserName,Password,StartDate,EndDate,IncludeRemoteInteractive)
	On Error Resume Next
	'Declare variables which will be used by this function
	Dim arrReportItem(6),objDicLogonDT,objDicLogoffDT,objWMIService
	Dim objRegExp,startDmtfDate,endDmtfDate,colItems,objItem,strOSVersion,strQueryString,objLogoffEvents
	Dim i,j,strReport,strLocalComputerName,strTargetLogonID,objLogoffEvent,intLogonType,strReportItem,intDiffHours
	'Create two Dictionary objects for saving logon/logoff event entries.
	Set objDicLogonDT = CreateObject("Scripting.Dictionary")
	Set objDicLogoffDT = CreateObject("Scripting.Dictionary")
	'Create Regular Expression object, it will be used to find out the specified logon type.
	Set objRegExp = New RegExp
	objRegExp.Global = True
	objRegExp.IgnoreCase = True
	'Parameters verification
	'If ComputerName is the name of local computer, connect the server without credential.
	'Otherwise, connect the server with the credential.
	strLocalComputerName = objNetwork.ComputerName
	If ComputerName <> strLocalComputerName Then
		Set objWMIService = objWMILocator.ConnectServer(ComputerName,"root\cimv2",UserName,Password)
	Else
		Set objWMIService = objWMILocator.ConnectServer(ComputerName,"root\cimv2")
	End If
	If Err Then
		arrReportItem(0) = ComputerName
		For i = 1 to 6
			arrReportItem(i) = "N/A"
		Next
		strReportItem = Chr(34) + Join(arrReportItem,""",""") + Chr(34)
		strReport = strReportItem + vbCrLf
		GetOSCUserLogonDuration = strReport
	End If	
	'This script need StartDate and EndDate to find out all the logon/logoff events occured
	'in a specific time range.
	If IsEmpty(StartDate) Or IsEmpty(EndDate) Then
		WScript.Echo "Please assign a valid value to StartDate and EndDate."
		WScript.Quit(1)
	Else
		startDmtfDate = Replace(StartDate,"/","") & "000000.000000+000"
		endDmtfDate = Replace(EndDate,"/","") & "235959.000000+000"
	End If
	'Find TimeZone Bias for correcting the report data
	intDiffHours = objWMIService.InstancesOf("Win32_TimeZone").ItemIndex(0).Properties_.Item("Bias")
	'Find the OS version number by WMI
	Set objItem = objWMIService.Get("Win32_OperatingSystem").Instances_.ItemIndex(0)
	strOSVersion = objItem.Properties_.Item("Version")
	'Check OS Version, This function will collect and process the logon/logoff data
	'from Windows Vista,Windows Server 2008,Windows 7 and Windows Server 2008 R2.
	If Left(strOSVersion,1) = "6" Then
		'Find out logon/logoff events
		strQueryString = "Select * from Win32_NTLogEvent Where " & _
		"LogFile='Security' And (TimeGenerated > '" & startDmtfDate & "' And TimeGenerated < '" & endDmtfDate & _
		"') And (EventCode='4624' OR EventCode='4647')"
		Set colItems = objWMIService.ExecQuery(strQueryString)
		'If the WMI query returns valid data, keep on processing.
		'Otherwise, display a message and return a "data not available" report.
		If colItems.Count <> 0 Then
			'Loop each item (Event Log Entry)
			For Each objItem In colItems
				With objItem.Properties_
					'Find Logon Event
					If .Item("EventCode") = "4624" Then
						'Base on function parameter, adjust the Pattern property of the Regular Expression object
						'2 - Interactive Logon; 10 - Remote Interactive Logon
						If (IncludeRemoteInteractive) Then
							objRegExp.Pattern = "2|10"
						Else
							objRegExp.Pattern = "2"
						End If
						'Keep useful logon event only.
						If Len(.Item("InsertionStrings")(4)) > 12 And objRegExp.Test(.Item("InsertionStrings")(8)) Then
							'Add logon events to the specific dictionary object.
							'The key of dictionary object will be Logon ID.
							strTargetLogonID = .Item("InsertionStrings")(7)
							If Not objDicLogonDT.Exists(strTargetLogonID) Then
								Call objDicLogonDT.Add(CStr(strTargetLogonID),objItem)
							End If
						End If
					ElseIf .Item("EventCode") = "4647" Then
						'Add logoff events to the specific dictionary object.
						'The key of dictionary object will be Logon ID.						
						strTargetLogonID = .Item("InsertionStrings")(3)
						If Not objDicLogoffDT.Exists(strTargetLogonID) Then
							Call objDicLogoffDT.Add(CStr(strTargetLogonID),objItem)
						End If
					End If
				End With
			Next
			'Find logon events base on the logoff events.
			objLogoffEvents = objDicLogoffDT.Items
			For i = 0 To objDicLogoffDT.Count - 1
				'Retrieve the object from dictionary
				Set objLogoffEvent = objLogoffEvents(i)
				'Convert LogonID to String
				strTargetLogonID = CStr(objLogoffEvent.Properties_.Item("InsertionStrings")(3))
				'Find logon events in the specific dictionary based on the LogonID
				If objDicLogonDT.Exists(strTargetLogonID) Then
					'Assign computer name to the report item array.
					arrReportItem(0) = ComputerName
					'Assign target domain name to the report item array.
					arrReportItem(1) = objLogoffEvent.Properties_.Item("InsertionStrings")(2)
					'Assign target user name to the report item array.
					arrReportItem(2) = objLogoffEvent.Properties_.Item("InsertionStrings")(1)
					'Translate logon type from number to meaningful words.
					'Then assign logon type to the report item array.
					intLogonType = objDicLogonDT.Item(strTargetLogonID).Properties_.Item("InsertionStrings")(8)
					If intLogonType = 2 Then
						arrReportItem(3) = "Interactive"
					ElseIf intLogonType = 10 Then
						arrReportItem(3) = "RemoteInteractive"
					End If
					'Assign logon date to the report item array.
					If objDicLogonDT.Exists(strTargetLogonID) Then
						arrReportItem(4) = DateAdd("n",intDiffHours,ConvertDMTFDateTime( _
						objDicLogonDT.Item(strTargetLogonID).Properties_.Item("TimeGenerated")))
					Else
						arrReportItem(4) = "N/A"		
					End If
					'Assign logoff date to the report item array.
					arrReportItem(5) = DateAdd("n",intDiffHours,ConvertDMTFDateTime(objLogoffEvent.Properties_.Item("TimeGenerated")))
					'Calculate logon duration (Minutes), then assign it to the report item array.
					If (arrReportItem(4) <> "N/A") Then
						'Convert the result value from day to minute.
						arrReportItem(6) = FormatNumber((CDate(arrReportItem(5)) - CDate(arrReportItem(4))) * 1440,2)
					End If
					'Generate the report item.
					strReportItem = Chr(34) + Join(arrReportItem,""",""") + Chr(34)
					strReport = strReport + strReportItem + vbCrLf
				End If
				'Clear Variables
				strReportItem = ""
				For j = 0 To 6
					arrReportItem(j) = ""
				Next
			Next
		Else
			WScript.Echo "Cannot find logon/logoff event entries on " & ComputerName & " within the specified time."
			arrReportItem(0) = ComputerName
			For i = 1 to 6
				arrReportItem(i) = "N/A"
			Next
			strReportItem = Chr(34) + Join(arrReportItem,""",""") + Chr(34)
			strReport = strReportItem + vbCrLf
		End If
	End if
	GetOSCUserLogonDuration = strReport
End Function


