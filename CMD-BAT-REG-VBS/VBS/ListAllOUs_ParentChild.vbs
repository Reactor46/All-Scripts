' ==================================================================================
' List All the OUs and Sub-OUs From Active Directory in Parent-Child Order
' Usage -- CScript /NoLogo ListAllOUs_ParentChild.vbs
' ==================================================================================

Option Explicit

Const ADS_SCOPE_SUBTREE = 2
Dim ObjConn, ObjRS, ObjRootDSE
Dim StrSQL, StrDomName, ObjOU

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing
StrSQL = "Select Name, ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'OrganizationalUnit' And Name <> 'Domain Controllers'"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL, ObjConn
If Not ObjRS.EOF Then
	ObjRS.MoveLast:	ObjRS.MoveFirst
	WScript.Echo vbNullString
	WScript.Echo "Total OU: " & Trim(ObjRS.RecordCount) & VbCrLf & "==================="
	WScript.Echo vbNullString
	While Not ObjRS.EOF
		Set ObjOU = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
		If StrComp(Right(Trim(ObjOU.Parent), Len(Trim(ObjOU.Parent)) - 7), StrDomName, VbTextCompare) = 0 Then
			WScript.Echo "Parent OU: " & Trim(ObjRS.Fields("Name").Value)
			GetChild(ObjOU)
		End If		
		ObjRS.MoveNext
		Set ObjOU = Nothing
	Wend
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

Private Sub GetChild(ThisObject)
	Dim ObjChild, StrThisParent
	For Each ObjChild In ThisObject
		If StrComp(Trim(ObjChild.Class), "OrganizationalUnit", VbTextCompare) = 0 Then
			WScript.Echo vbTab & ">> Child OU: " & Right(Trim(ObjChild.Name), Len(Trim(ObjChild.Name)) - 3)
			GetGrandChild (ObjChild.ADsPath)
		End If		
	Next
End Sub

Private Sub GetGrandChild (ThisADsPath)
	Dim ObjGrand, ObjItem
	Set ObjGrand = GetObject(ThisADsPath)
	For Each ObjItem In ObjGrand
		If StrComp(Trim(ObjItem.Class), "OrganizationalUnit", VbTextCompare) = 0 Then
			WScript.Echo vbTab & vbTab & ">> Child OU: " & Right(Trim(ObjItem.Name), Len(Trim(ObjItem.Name)) - 3)
		End If
		GetGrandChild Trim(ObjItem.ADsPath)
	Next	
	Set ObjGrand = Nothing
End Sub