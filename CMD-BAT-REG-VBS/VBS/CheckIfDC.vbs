' =============================================================
' How to determine if the Server is a Domain Controller
' Usage: CScript CheckIfDC.vbs
' =============================================================

Option Explicit

Dim ObjRootDSE, ObjThisObject, ObjConn, ObjRS
Dim StrDomName, StrMachineName, StrSQL

' -- Replace the value of StrMachineName with the name of actual Domain Controller
StrMachineName = "OurDomainDC"

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where Name = '" & StrMachineName & "'"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject"
ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL, ObjConn
If ObjRS.EOF Then
	WScript.Echo "This machine does not exist in Active Directory Domain"
End If
If Not ObjRS.EOF Then
	ObjRS.MoveFirst
	While Not ObjRS.EOF
		Set ObjThisObject = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
		If StrComp(Trim(ObjThisObject.Class), "COMPUTER", vbTextCompare) = 0 Then
			WScript.Echo "Machine: " & Trim(ObjThisObject.DNSHostName)
			WScript.Echo "Primary Group ID: " & Trim(ObjThisObject.primaryGroupID)
			If Trim(ObjThisObject.primaryGroupID) = 516 Then
				WScript.Echo "This machine is a DC"
			Else
				WScript.Echo "The machine: " & StrMachineName & " is NOT a DC"
			End If
		End If
		ObjRS.MoveNext
	Wend
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing