' ===================================================================================================================================================

' Determine Successfully Authenticated VPN Users Who Comply With NPS Authentication and NAP Client Health Policies
' Execute This Script as a Task on the VPN/NPS Server and this Task MUST run will Full Administrator privilege
' Written By: Monimoy Sanyal
' Contact: monimoys@hotmail.com

' Check the following:
' ---------------------
' Event ID 6272 — NPS Authentication Status
' >> http://technet.microsoft.com/en-us/library/cc735388(v=ws.10).aspx

' Event ID 6278 — NAP Client Health Status
' >> http://technet.microsoft.com/en-us/library/cc735338(v=ws.10).aspx

' ===================================================================================================================================================

Option Explicit

Dim StrComputer, ObjNetwork, ObjWMI, ColEvents, ObjEvent
Dim ObjFSO, WriteHandle, StrFilePath, EventDate, EventTime

Set ObjNetwork = CreateObject("WScript.Network")
StrComputer = Trim(ObjNetwork.ComputerName)
Set ObjNetwork = Nothing

DoThisDeleteJob

Set ObjWMI = GetObject("WinMgmts:{(Security)}\\" & StrComputer & "\Root\CIMV2")
Set ColEvents = ObjWMI.ExecNotificationQuery("Select * From __InstanceCreationEvent Where TargetInstance ISA 'Win32_NTLogEvent' AND TargetInstance.Logfile = 'Security'")
Do
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder) & "\"
	If ObjFSO.FileExists(StrFilePath & "VPNUserList.txt") = False Then
		Set WriteHandle = ObjFSO.OpenTextFile(StrFilePath & "VPNUserList.txt", 8, True, 0)
		WriteHandle.Close:	Set WriteHandle = Nothing
	End If
	Set ObjEvent = ColEvents.NextEvent	
	If StrComp(ObjEvent.TargetInstance.SourceName, "Microsoft-Windows-Security-Auditing", vbTextCompare) = 0 AND ObjEvent.TargetInstance.EventCode = 6272 OR ObjEvent.TargetInstance.EventCode = 6278 Then 
		Set WriteHandle = ObjFSO.OpenTextFile(StrFilePath & "VPNUserList.txt", 8, True, 0)
		WriteHandle.WriteLine "Event ID: " & ObjEvent.TargetInstance.EventCode
		If Trim(ObjEvent.TargetInstance.User) <> vbNullString Then
			WriteHandle.WriteLine "User: " & ObjEvent.TargetInstance.User
		End If
		WriteHandle.WriteLine "Source Name: " & ObjEvent.TargetInstance.SourceName
		WriteHandle.WriteLine "Message: " & ObjEvent.TargetInstance.Message
		EventDate = ObjEvent.TargetInstance.TimeWritten
		EventTime = GetProperTimeInfo(EventDate)
		WriteHandle.WriteLine "Time of Event: " & EventTime
		If Trim(ObjEvent.TargetInstance.EventType) = 0 Then
			WriteHandle.WriteLine "Event Type: 0 -- SUCCESS"
		End If
		WriteHandle.WriteLine " ==================================================================================================="
		WriteHandle.WriteLine vbNullString
		WriteHandle.Close:	Set WriteHandle = Nothing
	End If	
	Set ObjEvent = Nothing:	Set ObjFSO = Nothing:	StrFilePath = vbNullString
Loop
Set ColEvents = Nothing:	Set ObjWMI = Nothing
WScript.Quit

Function GetProperTimeInfo(dtmEventDate)
	GetProperTimeInfo = CDate(Mid(dtmEventDate, 5, 2) & "/" & Mid(dtmEventDate, 7, 2) & "/" & Left(dtmEventDate, 4) _
		& " " & Mid (dtmEventDate, 9, 2) & ":" & Mid(dtmEventDate, 11, 2) & ":" & Mid(dtmEventDate, 13, 2))
End Function

Private Sub DoThisDeleteJob
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder) & "\"
	If ObjFSO.FileExists(StrFilePath & "VPNUserList.txt") = True Then
		ObjFSO.DeleteFile StrFilePath & "VPNUserList.txt", True
	End If
	Set ObjFSO = Nothing:	StrFilePath = vbNullString
End Sub