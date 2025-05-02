' ===============================================================================================================
' All Attributes in the Schema Partition That Are Replicated to the Global Catalog Server
' Use LDAP Filter -- (&(objectCategory=attributeSchema)(isMemberOfPartialAttributeSet=TRUE)) 
' Usage: CScript /NoLogo ReplSchemaAttributes.vbs > SchemaReport.txt
' ===============================================================================================================

Option Explicit

Dim ObjRootDSE, StrSQL, StrSchema
Dim ObjConn, ObjRS, Counter

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrSchema = Trim(ObjRootDSE.Get("SchemaNamingContext"))
Set ObjRootDSE = Nothing

StrSQL = "<LDAP://" & StrSchema & ">;(&(objectCategory=attributeSchema)(isMemberOfPartialAttributeSet=TRUE));Name;Subtree"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL, ObjConn
ObjRS.MoveLast:	ObjRS.MoveFirst
WScript.Echo "All the following attributes in the Schema Partition is replicated to the Global Catalog Server."
WScript.Echo "Total No of Attributes: " & ObjRS.RecordCount
WScript.Echo "===================================================================================================" & VbCrLf
While Not ObjRS.EOF
	WScript.Echo "Attribute Name: " & vbTab & Trim(ObjRS.Fields("Name").Value)
	Counter = Counter + 1
	ObjRS.MoveNext
Wend
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

WScript.Echo vbNullString
WScript.Echo "Total No of Attributes: " & Counter
WScript.Echo "============================="
WScript.Quit