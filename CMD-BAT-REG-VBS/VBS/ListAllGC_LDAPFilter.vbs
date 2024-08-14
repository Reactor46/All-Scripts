' =====================================================================================================================
' Get List Of All Global Catalogue Servers with LDAP Filter
'
' Reference For LDAP Filters:
' >> http://social.technet.microsoft.com/wiki/contents/articles/5392.active-directory-ldap-syntax-filters.aspx
' =====================================================================================================================

Option Explicit

Dim ObjConn, ObjRS, ObjRootDSE, ObjDC, ObjServer
Dim StrDomName, StrSQL, StrFilter

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("ConfigurationNamingContext"))
WScript.Echo vbNullString
WScript.Echo "Please wait. Listing all Global Catalog Servers ... "
WScript.Echo vbNullString
Set ObjRootDSE = Nothing

StrSQL = "<LDAP://" & StrDomName & ">;"
StrFilter = "(&(objectCategory=nTDSDSA)(options:1.2.840.113556.1.4.803:=1))"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL & StrFilter & ";ADsPath;SubTree", ObjConn
If Not ObjRS.EOF Then
	ObjRS.MoveLast:	ObjRS.MoveFirst
	WScript.Echo "Total No of Global Catalog Servers: " & ObjRS.RecordCount
	WScript.Echo vbNullString
	While Not ObjRS.EOF
		Set ObjServer = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
		Set ObjDC = GetObject( ObjServer.Parent )
		WScript.Echo ">> DC " & ObjRS.AbsolutePosition & " -- " & Trim(ObjDC.Get("DNSHostName"))
		Set ObjDC = Nothing:	Set ObjServer = Nothing
		ObjRS.MoveNext
	Wend
End If

ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing