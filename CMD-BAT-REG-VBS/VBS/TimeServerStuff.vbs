' =================================================================================================
' Get Time and Time Server Information of Remotes Machines In The Domain From a Text File
' Usage: CScript TimeServerStuff.vbs
' =================================================================================================

Option Explicit

Dim ObjShell, ObjRootDSE, ObjExec
Dim ObjConn, ObjRS, ObjFSO
Dim ObjWMI, ColItems, ObjItem
Dim WriteHandle, StrFilePath, StrServerName, StrDomain

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
If ObjFSO.FileExists(StrFilePath & "\ServerList.txt") = False Then
	WScript.Echo "Error: Cannot Execute This Script." & VbCrLf & "The File -- ServerList.txt -- Does Not Exist."
	Set ObjFSO = Nothing:	WScript.Quit
Else
	Set ObjFSO = Nothing:	GetTheFilePath
End If

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrFilePath & "\;Extended Properties=""text;HDR=YES;FMT=Delimited"""
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open "Select * From ServerList.txt", ObjConn
ObjRS.MoveFirst
While Not ObjRS.EOF
	StrServerName = Trim(ObjRS.Fields("ServerName").Value)	
	GetTimeServerInfo (StrServerName)	
	ObjRS.MoveNext
Wend
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

WScript.Echo VbCrLf & "Tasks completed. Check log file -- TimeServerLog.txt -- for details."
WScript.Quit

Private Sub GetTimeServerInfo (StrMachineName)

	Dim LocalTime, StrMonth, StrDay, StrYear, DisplayDate
	Dim StrHour, StrMinutes, StrSeconds, DisplayTime, StrAMPM
	
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
	If ObjFSO.FileExists(StrFilePath & "\TimeServerLog.txt") = True Then		
		Set WriteHandle = ObjFSO.OpenTextFile(StrFilePath & "\TimeServerLog.txt", 8, True, 0)
		WriteHandle.WriteLine vbNullString
		WriteHandle.WriteLine "Computer Name: " & StrMachineName		

		Set ObjRootDSE = GetObject("LDAP://RootDSE")
		StrDomain = Trim(ObjRootDSE.Get("DefaultNamingContext"))
		Set ObjRootDSE = Nothing
		StrDomain = Replace(StrDomain, "DC=", vbNullString):	StrDomain = Replace(StrDomain, ",", ".")
		WriteHandle.WriteLine "Domain Name: " & StrDomain
		WriteHandle.WriteLine "Computer FQDN: " & StrMachineName & "." & StrDomain

		Set ObjShell = CreateObject("WScript.Shell")
		Set ObjExec = ObjShell.Exec("%comspec% /c W32tm /Query /Source")  
		WriteHandle.WriteLine "This Computer's Time Server: " & ObjExec.StdOut.ReadAll
		Set ObjExec= Nothing:	Set ObjShell = Nothing

		Set ObjWMI = GetObject("WinMgmts:\\" & StrMachineName & "\Root\CimV2")
		Set ColItems = ObjWMI.ExecQuery("Select * From Win32_TimeZone",,48)
		For Each ObjItem In ColItems
			WriteHandle.WriteLine "Time Server Follows: " & Trim(ObjItem.StandardName)
			WriteHandle.WriteLine "Time Zone on Time Server: " & Trim(ObjItem.Caption)
			WriteHandle.WriteLine "Description: " & Trim(ObjItem.Description)    
		Next
		Set ColItems = Nothing:	Set ObjWMI = Nothing		
		
		Set ObjWMI = GetObject("WinMgmts:\\" & StrMachineName & "\Root\CimV2")
		Set ColItems = ObjWMI.ExecQuery("Select * From Win32_OperatingSystem",,48)
		For Each ObjItem In ColItems
			LocalTime = Trim(ObjItem.LocalDateTime)
			StrMonth = Mid(LocalTime, 5, 2)
			StrDay = Mid(LocalTime, 7, 2)
			StrYear = Left(LocalTime, 4)
			StrHour = Mid(LocalTime, 9, 2)
			StrMinutes = Mid(LocalTime, 11, 2)
			StrSeconds = Mid(LocalTime, 13, 2)
		Next
		DisplayDate = StrDay & "/" & StrMonth & "/" & StrYear
		DisplayTime =  StrHour & ":" & StrMinutes & ":" & StrSeconds
		WriteHandle.WriteLine "Current Time on Machine " & StrMachineName & ": " & DisplayTime & " -- 24 Hour Time Format"
		WriteHandle.WriteLine "Current Date on Machine " & StrMachineName & ": " & DisplayDate & " -- Format DD/MM/YYYY"
		Set ColItems = Nothing:	Set ObjWMI = Nothing		
		
		Set ObjWMI = GetObject("WinMgmts:\\" & StrMachineName & "\Root\CimV2")
		Set ColItems = ObjWMI.ExecQuery("Select * From Win32_LocalTime",,48)
		For Each ObjItem In ColItems
			StrMonth = Trim(ObjItem.Month):	StrDay = Trim(ObjItem.Day):	StrYear = Trim(ObjItem.Year)
			DisplayDate = StrDay & "/" & StrMonth & "/" & StrYear
			StrHour = Trim(ObjItem.Hour)
			If StrHour < 12 Then 
				StrAMPM = "AM"
			Else
				StrHour = StrHour - 12
				StrAMPM = "PM"
			End If		  
			StrMinutes = Trim(ObjItem.Minute)
			If StrMinutes < 10 Then
				StrMinutes = "0" & StrMinutes
			End If
			StrSeconds = Trim(ObjItem.Second)
			If StrSeconds < 10 Then
				StrSeconds = "0" & StrSeconds
			End If
			DisplayTime = StrHour & ":" & StrMinutes & ":" & StrSeconds & " " & StrAMPM
			WriteHandle.WriteLine "Current Time on Machine " & StrMachineName & ": " & DisplayTime & " -- 12 Hour Time Format AM/PM"
			WriteHandle.WriteLine "Current Date on Machine " & StrMachineName & ": " & DisplayDate & " -- Format DD/MM/YYYY"
		Next		
		WriteHandle.Close:	Set WriteHandle = Nothing:	Set ObjFSO = Nothing
	End If
	
End Sub

Private Sub GetTheFilePath
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
	If ObjFSO.FileExists(StrFilePath & "\TimeServerLog.txt") = False Then		
		Set WriteHandle = ObjFSO.OpenTextFile(StrFilePath & "\TimeServerLog.txt", 8, True, 0)
		WriteHandle.Close:	Set WriteHandle = Nothing
	End If
	If ObjFSO.FileExists(StrFilePath & "\TimeServerLog.txt") = True Then
		Set WriteHandle = ObjFSO.OpenTextFile(StrFilePath & "\TimeServerLog.txt", 8, True, 0)
		WriteHandle.WriteLine "# ========== Logging: " & FormatDateTime (Date, 1) & " at " & FormatDateTime (Now, 3) & " =========="
		WriteHandle.Close:	Set WriteHandle = Nothing
	End If	
	Set ObjFSO = Nothing
End Sub
