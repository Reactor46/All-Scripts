Option Explicit

Dim StrComputer, ObjWMI, ColItems, ObjItem, StrIP
Dim ObjConn, ObjRS, ObjFSO, StrFilePath, BlnComp, StrLogFile

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
Set ObjFSO = Nothing

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrFilePath & "\;Extended Properties=""text;HDR=YES;FMT=Delimited"""
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open "Select * From InputFile.txt", ObjConn
If Not ObjRS.EOF Then
	DoThisDeleteJob:	ObjRS.MoveFirst:	PrepareLogFile
	WScript.Echo
	WScript.Echo "Performing Task. Please wait ..."
	WScript.Echo
	While Not ObjRS.EOF
		BlnComp = False		
		DoTheCheckAction
		If BlnComp = True Then
			' -- Rename the Computer
			Call TakeRenameAction(Trim(ObjRS.Fields("TxtIPAddr").Value), Trim(ObjRS.Fields("NewCompName").Value))
		Else
			' -- Write To Log
			Call MakeLogEntry("NotFound", "No", "No")
		End If
		ObjRS.MoveNext
	Wend
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing
WScript.Echo
WScript.Echo "Task Completed !!"
WScript.Echo "Check Log File -- " & StrFilePath & "\LogFile.txt -- For Details"
WScript.Quit

Private Sub DoTheCheckAction

	On Error Resume Next

	StrComputer = Trim(ObjRS.Fields("TxtIPAddr").Value)
	WScript.Echo vbTab & ">> Attempting to Rename Computer Having IP Address: " & StrComputer
	WScript.Echo vbTab & "   New Computer Name: " & Trim(ObjRS.Fields("NewCompName").Value)
	WScript.Echo
	Set ObjWMI = GetObject("WinMgmts:\\" & StrComputer & "\Root\CIMV2")
	Set ColItems = ObjWMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration")
	For Each ObjItem In ColItems
		If Trim(ObjItem.MACAddress) <> vbNullString Then		
			For Each StrIP In ObjItem.TxtIPAddress
				If StrIP <> vbNullString Then
					' -- WScript.Echo "MAC -- " & Trim(ObjItem.MACAddress)
					' -- WScript.Echo "IP  -- " & Trim(StrIP)
					' -- WScript.Echo
					If StrComp(StrComputer, Trim(StrIP), vbBinaryCompare) = 0 AND StrComp(Trim(ObjRS.Fields("MACAddr").Value), Trim(ObjItem.MACAddress), vbBinaryCompare) = 0 Then
						' -- WScript.Echo "IP Equal"
						BlnComp = True
					Else
						BlnComp = False
					End If
				End If
			Next
		End If
	Next
	Set ColItems = Nothing:	Set ObjWMI = Nothing

End Sub

Private Sub TakeRenameAction(StrTargetIP, StrCompName)

	Dim StrUserName, StrPassword, ObjComputer, RetValue
	
	' -- Typically Administrator has the Privilege to Rename Computer
	' -- Put Administrator UserName Here
	StrUsername = "Put Administrator UserName"
	StrPassword = "Put Administrator Password"
	
	Set ObjWMI = GetObject("WinMgmts:\\" & StrTargetIP & "\Root\CIMV2")
	' --- Call always gets only one Win32_ComputerSystem object.
	For Each ObjComputer In ObjWMI.InstancesOf("Win32_ComputerSystem")
		RetValue = ObjComputer.Rename(StrCompName, StrPassword, StrUsername)
		If RetValue <> 0 Then
			Call MakeLogEntry("ERROR", Err.Number, Err.Description)
			Err.Clear
		Else
			' -- Rename Succeeded. Write Success In Log File
			Call MakeLogEntry("SUCCESS", "Yes", "Yes")			
        End If
	Next
	Set ObjWMI = Nothing
	
End Sub

Private Sub PrepareLogFile
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = False Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		StrLogFile.WriteLine "# ========================================================================="
		StrLogFile.WriteLine "# Read Input File AND Rename Multiple Computers"
		StrLogFile.WriteLine "# Logging Started At: " & FormatDateTime (Date, 1) & " -- " & FormatDateTime (Now, 3)
		StrLogFile.WriteLine "# ========================================================================="
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If	
	Set ObjFSO = Nothing
End Sub

Private Sub DoThisDeleteJob
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		ObjFSO.DeleteFile StrFilePath & "\LogFile.txt", True
	End If
	Set ObjFSO = Nothing
End Sub

Private Sub MakeLogEntry(StrWhat, ErrNum, ErrDesc)
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		If StrComp(StrWhat, "SUCCESS", vbTextCompare) = 0 Then
			StrLogFile.WriteLine "SUCCESS -- Computer Having IP Address: " & Trim(ObjRS.Fields("TxtIPAddr").Value)
			StrLogFile.WriteLine "This Computer Has Been Renamed To: " & Trim(ObjRS.Fields("NewCompName").Value)
			StrLogFile.WriteLine "Must Reboot This Computer For New Name To Take Effect."
			StrLogFile.WriteLine vbNullString			
		End If
		If StrComp(StrWhat, "ERROR", vbTextCompare) = 0 Then
			StrLogFile.WriteLine "FAILED --- Computer Having IP Address: " & Trim(ObjRS.Fields("TxtIPAddr").Value)
			StrLogFile.WriteLine "Cannot Rename This Computer. Error: " & ErrNum
			StrLogFile.WriteLine "REASON For Failure --- " & ErrDesc
			StrLogFile.WriteLine vbNullString			
		End If
		If StrComp(StrWhat, "NotFound", vbTextCompare) = 0 Then
			StrLogFile.WriteLine "FAILED --- Computer Having IP Address: " & Trim(ObjRS.Fields("TxtIPAddr").Value)
			StrLogFile.WriteLine "Cannot Rename This Computer."
			StrLogFile.WriteLine "REASON For Failure --- Unable To Connect To This Computer."
			StrLogFile.WriteLine vbNullString			
		End If
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If	
	Set ObjFSO = Nothing
End Sub