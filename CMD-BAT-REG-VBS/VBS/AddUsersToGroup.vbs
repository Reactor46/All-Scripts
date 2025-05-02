' ============================================================================
' Bulk Add Multiple User-Accounts To Multiple Groups From Text File
' Usage: CScript AddUserToGroup.vbs
' ============================================================================

Option Explicit

Dim StrDomain, StrUserName, StrGroupName
Dim StrLogFile, StrFilePath, ChkUser, ChkGroup
Dim ObjRootDSE, ObjConn, ObjRS, ObjFSO, ObjGroup
Dim ObjNewConn, ObjNewRS, UserNameLDAPPath, GroupNameLDAPPath

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomain = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
If ObjFSO.FileExists(StrFilePath & "\NameList.txt") = False Then
	WScript.Echo "Error. The required file NameList.txt is missing."
	Set ObjFSO = Nothing:	WScript.Quit
Else
	StrFilePath = vbNullString:	Set ObjFSO = Nothing
End If

LogFileAction

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrFilePath & "\;Extended Properties=""text;HDR=YES;FMT=Delimited"""
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open "Select * From NameList.txt", ObjConn
If ObjRS.EOF Then
	ObjRS.Close:	Set ObjRS = Nothing
	ObjConn.Close:	Set ObjConn = Nothing
	WScript.Echo "Error. The file NameList.txt has no input parameter."
	WScript.Quit
Else
	ObjRS.MoveFirst
End If
WScript.Echo "Working. Please wait ..." & VbCrLf
While Not ObjRS.EOF
	If Trim(ObjRS.Fields("UserName").Value) <> vbNullString Then
		StrUserName = Trim(ObjRS.Fields("UserName").Value)
	End If
	If Trim(ObjRS.Fields("GroupName").Value) <> vbNullString Then
		StrGroupName = Trim(ObjRS.Fields("GroupName").Value)
	End If
	ChkUser = False:	ChkGroup = False
	If Trim(ObjRS.Fields("UserName").Value) <> vbNullString AND Trim(ObjRS.Fields("GroupName").Value) <> vbNullString Then
		GetLDAPPath("User"):	GetLDAPPath("Group")
	End If
	If ChkUser = False Then
		Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
		If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
			Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
			If Trim(ObjRS.Fields("UserName").Value) <> vbNullString Then
				StrLogFile.WriteLine "Error: User " & StrUserName & " Not Found in Active Directory"
			End If
			If Trim(ObjRS.Fields("UserName").Value) = vbNullString Then
				StrLogFile.WriteLine "Error: UserName is NULL"
			End If
			StrLogFile.Close:	Set StrLogFile = Nothing
		End If
		Set ObjFSO = Nothing
	End If
	If ChkGroup = False Then
		Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
		If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
			Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
			If Trim(ObjRS.Fields("GroupName").Value) <> vbNullString Then
				StrLogFile.WriteLine "Error: Group " & StrGroupName & " Not Found in Active Directory"
			End If
			If Trim(ObjRS.Fields("GroupName").Value) = vbNullString Then
				StrLogFile.WriteLine "Error: GroupName is NULL"
			End If
			StrLogFile.Close:	Set StrLogFile = Nothing
		End If
		Set ObjFSO = Nothing
	End If
	If ChkUser = True AND ChkGroup = True Then
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
		Case "User"
			Set ObjNewRS = ObjNewConn.Execute("SELECT ADsPath FROM 'LDAP://" & StrDomain & "' Where ObjectCategory = 'User' AND SamAccountName = '" & StrUserName & "'")
	End Select	
	If Not ObjNewRS.EOF Then
		ObjNewRS.MoveFirst
		Select Case StrWhat
			Case "Group"
				GroupNameLDAPPath = Trim(ObjNewRS.Fields("ADsPath").Value)
				ChkGroup = True
			Case "User"
				UserNameLDAPPath = Trim(ObjNewRS.Fields("ADsPath").Value)
				ChkUser = True
		End Select
	Else
		Select Case StrWhat
			Case "Group"
				ChkGroup = False
				WScript.Echo StrGroupName & ": " & "This Group not found in Active Directory"
			Case "User"
				ChkUser = False
				WScript.Echo StrUserName & ": " & "This User-Account not found in Active Directory"
		End Select		
	End If
	ObjNewRS.Close:	Set ObjNewRS = Nothing
	ObjNewConn.Close:	Set ObjNewConn = Nothing
End Sub

Private Sub LogFileAction
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	StrFilePath = Trim(ObjFSO.GetFile(WScript.ScriptFullName).ParentFolder)
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = False Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		StrLogFile.WriteLine "# ========== Logging: " & FormatDateTime (Date, 1) & " at " & FormatDateTime (Now, 3) & " =========="
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If	
	Set ObjFSO = Nothing
End Sub

Private Sub AddToGroup
	Dim GroupMember
	Set ObjGroup = GetObject(GroupNameLDAPPath)
	For Each GroupMember in ObjGroup.Members
		WScript.Echo vbTab & ">> Currently working on User-Account: " & StrUserName
		If LCase(GroupMember.ADsPath) = LCase(UserNameLDAPPath) Then
			Set ObjFSO = CreateObject("Scripting.FileSystemObject")	
			If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
				Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
				StrLogFile.WriteLine StrUserName & " Not Added. User-Account is already a Member of the Group: " & StrGroupName
				StrLogFile.Close:	Set StrLogFile = Nothing
			End If
			Set ObjFSO = Nothing:	Exit Sub
		End If
	Next
	ObjGroup.Add(UserNameLDAPPath)
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		Set StrLogFile = ObjFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		If Err.Number = 0 Then
			StrLogFile.WriteLine "Success. User " & StrUserName & " added to the Group: " & StrGroupName
		Else
			StrLogFile.WriteLine "Failed with Error: " & Err.Number & " For User: " & StrUserName
			StrLogFile.WriteLine "User Cannot be added Group: " & StrGroupName
			StrLogFile.WriteLine Err.Description
		End If				
		StrLogFile.Close:	Set StrLogFile = Nothing
	End If
	Set ObjFSO = Nothing	
End Sub