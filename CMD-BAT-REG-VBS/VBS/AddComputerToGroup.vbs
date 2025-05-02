' ============================================================================
' Bulk Add Multiple Computer-Accounts To Multiple Groups From Text File
' Usage: CScript AddComputerToGroup.vbs
' ============================================================================

Option Explicit

Dim StrDomain, StrCompName, StrGroupName
Dim StrLogFile, StrFilePath, ChkComp, ChkGroup
Dim ObjRootDSE, ObjConn, ObjRS, ObjFSO, ObjGroup
Dim ObjNewConn, ObjNewRS, CompNameLDAPPath, GroupNameLDAPPath

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomain = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
If ObjFSO.FileExists(StrFilePath & "\ComputerList.txt") = False Then
	WScript.Echo "Error. The required file ComputerList.txt is missing."
	Set ObjFSO = Nothing:	WScript.Quit
Else
	StrFilePath = vbNullString:	Set ObjFSO = Nothing
End If

LogFileAction

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrFilePath & "\;Extended Properties=""text;HDR=YES;FMT=Delimited"""
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open "Select * From ComputerList.txt", ObjConn
If ObjRS.EOF Then
	ObjRS.Close:	Set ObjRS = Nothing
	ObjConn.Close:	Set ObjConn = Nothing
	WScript.Echo "Error. The file ComputerList.txt has no input parameter."
	WScript.Quit
Else
	ObjRS.MoveFirst
End If
WScript.Echo "Working. Please wait ..." & VbCrLf
While Not ObjRS.EOF	
	If Trim(ObjRS.Fields("ComputerName").Value) <> vbNullString Then
		StrCompName = Trim(ObjRS.Fields("ComputerName").Value)
	End If
	If Trim(ObjRS.Fields("GroupName").Value) <> vbNullString Then
		StrGroupName = Trim(ObjRS.Fields("GroupName").Value)
	End If
	ChkComp = False:	ChkGroup = False
	If Trim(ObjRS.Fields("ComputerName").Value) <> vbNullString AND Trim(ObjRS.Fields("GroupName").Value) <> vbNullString Then
		GetLDAPPath("Computer"):	GetLDAPPath("Group")
	End If
	If ChkComp = False Then
		Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
		If ObjFSO.FileExists(StrFilePath & "\ComputerLogFile.txt") = True Then
			Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\ComputerLogFile.txt", 8, True, 0)
			StrLogFile.WriteLine "Error: Computer " & StrCompName & " Not Found in Active Directory"
			StrLogFile.Close:	Set StrLogFile = Nothing
		End If
		Set ObjFSO = Nothing
	End If
	If ChkGroup = False Then
		Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
		If ObjFSO.FileExists(StrFilePath & "\ComputerLogFile.txt") = True Then
			Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\ComputerLogFile.txt", 8, True, 0)
			StrLogFile.WriteLine "Error: Group " & StrGroupName & " Not Found in Active Directory"
			StrLogFile.Close:	Set StrLogFile = Nothing
		End If
		Set ObjFSO = Nothing
	End If
	If ChkComp = True AND ChkGroup = True Then
		AddToGroup
	End If
	ObjRS.MoveNext
Wend
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

WScript.Echo VbCrLf & "Tasks completed. Check log file for details."
WScript.Quit

Private Sub GetLDAPPath(StrWhat)
	Set ObjNewConn = CreateObject("ADODB.Connection")
	ObjNewConn.Provider = "ADsDSOObject":	ObjNewConn.Open "Active Directory Provider"
	Select Case StrWhat
		Case "Group"
			Set ObjNewRS = ObjNewConn.Execute("SELECT ADsPath FROM 'LDAP://" & StrDomain & "' Where ObjectCategory = 'Group' AND Name = '" & StrGroupName & "'")
		Case "Computer"
			Set ObjNewRS = ObjNewConn.Execute("SELECT ADsPath FROM 'LDAP://" & StrDomain & "' Where ObjectCategory = 'Computer' AND Name = '" & StrCompName & "'")
	End Select	
	If Not ObjNewRS.EOF Then
		ObjNewRS.MoveFirst
		Select Case StrWhat
			Case "Group"
				GroupNameLDAPPath = Trim(ObjNewRS.Fields("ADsPath").Value)
				ChkGroup = True
			Case "Computer"
				CompNameLDAPPath = Trim(ObjNewRS.Fields("ADsPath").Value)
				ChkComp = True
		End Select
	Else
		Select Case StrWhat
			Case "Group"
				ChkGroup = False
				WScript.Echo StrGroupName & " -- " & "Group not found in Active Directory"
			Case "Computer"
				ChkComp = False
				WScript.Echo StrCompName & " -- " & "Computer Object not found in Active Directory"
		End Select		
	End If
	ObjNewRS.Close:	Set ObjNewRS = Nothing
	ObjNewConn.Close:	Set ObjNewConn = Nothing
End Sub

Private Sub AddToGroup
	Dim GroupMember
	Set ObjGroup = GetObject(GroupNameLDAPPath)
	For Each GroupMember in ObjGroup.Members
		If LCase(GroupMember.ADsPath) = LCase(CompNameLDAPPath) Then
			Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
			If ObjFSO.FileExists(StrFilePath & "\ComputerLogFile.txt") = True Then
				Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\ComputerLogFile.txt", 8, True, 0)
				StrLogFile.WriteLine "Object Not Added. Computer " & StrCompName & " is already a Member of the Group: " & StrGroupName
				StrLogFile.Close:	Set StrLogFile = Nothing
			End If
			Set ObjFSO = Nothing:	Exit Sub
		End If
	Next
	ObjGroup.Add(CompNameLDAPPath)
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FileExists(StrFilePath & "\ComputerLogFile.txt") = True Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\ComputerLogFile.txt", 8, True, 0)
		If Err.Number = 0 Then
			StrLogFile.WriteLine "Success. Computer " & StrCompName & " added to the Group: " & StrGroupName
		Else
			StrLogFile.WriteLine "Failed with Error: " & Err.Number & " For User: " & StrCompName
			StrLogFile.WriteLine "User Cannot be added Group: " & StrGroupName
			StrLogFile.WriteLine Err.Description
		End If				
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If
	Set ObjFSO = Nothing	
End Sub

Private Sub LogFileAction
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
	If ObjFSO.FileExists(StrFilePath & "\ComputerLogFile.txt") = False Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\ComputerLogFile.txt", 8, True, 0)
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If
	If ObjFSO.FileExists(StrFilePath & "\ComputerLogFile.txt") = True Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\ComputerLogFile.txt", 8, True, 0)
		StrLogFile.WriteLine "# ========== Logging: " & FormatDateTime (Date, 1) & " at " & FormatDateTime (Now, 3) & " =========="
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If	
	Set ObjFSO = Nothing
End Sub