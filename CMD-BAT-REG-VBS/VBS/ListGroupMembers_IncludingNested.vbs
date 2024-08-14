' =============================================================
' List All Members of a Group; Including Nested Members
' =============================================================

Option Explicit

Dim ObjRootDSE, ObjConn, ObjRS, ObjCustom
Dim StrDomainName, StrGroupName, StrSQL
Dim StrGroupDN, StrEmptySpace

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomainName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

' -- Mention any AD Group Name Here. Also works for Domain Admins, Enterprise Admins etc.
StrGroupName = "Domain Admins"
StrSQL = "Select ADsPath From 'LDAP://" & StrDomainName & "' Where ObjectCategory = 'Group' AND Name = '" & StrGroupName & "'"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL, ObjConn
If ObjRS.EOF Then
	WScript.Echo VbCrLf & "This Group: " & StrGroupName & " does not exist in Active Directory"
End If
If Not ObjRS.EOF Then	
	WScript.Echo vbNullString
	ObjRS.MoveLast:	ObjRS.MoveFirst
	WScript.Echo "Total No of Groups Found: " & ObjRS.RecordCount
	WScript.Echo "List of Members In " & StrGroupName & " are: " & VbCrLf
	While Not ObjRS.EOF		
		StrGroupDN = Trim(ObjRS.Fields("ADsPath").Value)
		Set ObjCustom = CreateObject("Scripting.Dictionary")
		StrEmptySpace = " "
		GetAllNestedMembers StrGroupDN, StrEmptySpace, ObjCustom
		Set ObjCustom = Nothing
		ObjRS.MoveNext
	Wend
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

Private Function GetAllNestedMembers (StrGroupADsPath, StrEmptySpace, ObjCustom)
	Dim ObjGroup, ObjMember
	Set ObjGroup = GetObject(StrGroupADsPath)
	For Each ObjMember In ObjGroup.Members		
		WScript.Echo Trim(ObjMember.CN) & " --- " & Trim(ObjMember.DisplayName) & " (" & Trim(ObjMember.Class) & ")"
		If Strcomp(Trim(ObjMember.Class), "Group", vbTextCompare) = 0 Then
			If ObjCustom.Exists(ObjMember.ADsPath) Then	
				WScript.Echo StrEmptySpace & " -- Already Checked Group-Member " & "(Stopping Here To Escape Loop)"
			Else
				ObjCustom.Add ObjMember.ADsPath, 1	
				GetFromHere ObjMember.ADsPath, StrEmptySpace & " ", ObjCustom
			End If
		End If
	Next
End Function

Private Sub GetFromHere(StrGroupADsPath, StrEmptySpace, ObjCustom)
	Dim ObjThisGroup, ObjThisMember
	Set ObjThisGroup = GetObject(StrGroupADsPath)
	WScript.Echo vbNullString
	WScript.Echo "  ** Members of this Group are:"
	For Each ObjThisMember In ObjThisGroup.Members		
		WScript.Echo "    >> " & Trim(ObjThisMember.CN) & " --- " & Trim(ObjThisMember.DisplayName) & " (" & Trim(ObjThisMember.Class) & ")"
	Next
	WScript.Echo vbNullString
End Sub