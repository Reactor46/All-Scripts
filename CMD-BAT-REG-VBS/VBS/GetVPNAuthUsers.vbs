' =======================================================================================================================================
' Determine Successfully Authenticated VPN Users Who Comply With NPS Authentication and NAP Client Health Policies
' Written By: Monimoy Sanyal
' Contact: monimoys@hotmail.com

' Check the following:
' ---------------------
' Event ID 6272 — NPS Authentication Status
' >> http://technet.microsoft.com/en-us/library/cc735388(v=ws.10).aspx

' Event ID 6278 — NAP Client Health Status
' >> http://technet.microsoft.com/en-us/library/cc735338(v=ws.10).aspx

' ======================================================================================================================================

Option Explicit

Dim ObjNetwork, StrComputer, ObjWMI, ColEvents, ObjEvent
Dim EventDate, EventTime

Set ObjNetwork = CreateObject("WScript.Network")
StrComputer = Trim(ObjNetwork.ComputerName)
Set ObjNetwork = Nothing

Set ObjWMI = GetObject("WinMgmts:" & "{impersonationLevel=Impersonate,(Security)}!\\" & StrComputer & "\Root\CIMV2")
Set ColEvents = ObjWMI.ExecQuery("Select * FROM Win32_NTLogEvent WHERE Logfile = 'Security' AND EventCode = 6272 OR EventCode = 6278")
WScript.Echo
For Each ObjEvent In ColEvents
	WScript.Echo "Event ID: " & ObjEvent.EventCode
	If Trim(ObjEvent.User) <> vbNullString Then
		WScript.Echo "User: " & ObjEvent.User
	End If
	WScript.Echo "Source Name: " & ObjEvent.SourceName
	WScript.Echo "Message: " & ObjEvent.Message
	EventDate = ObjEvent.TimeWritten
	EventTime = GetProperTimeInfo(EventDate)
	WScript.Echo "Time of Event: " & EventTime
	WScript.Echo " ==================================================================================================="
	WScript.Echo
Next
Set ColEvents = Nothing:	Set ObjWMI = Nothing
WScript.Quit

Function GetProperTimeInfo(dtmEventDate)
	GetProperTimeInfo = CDate(Mid(dtmEventDate, 5, 2) & "/" & Mid(dtmEventDate, 7, 2) & "/" & Left(dtmEventDate, 4) _
		& " " & Mid (dtmEventDate, 9, 2) & ":" & Mid(dtmEventDate, 11, 2) & ":" & Mid(dtmEventDate, 13, 2))
End Function
